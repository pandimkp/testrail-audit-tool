require('dotenv').config();
const express = require('express');
const fetch = require('node-fetch');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const {
  TESTRAIL_BASE_URL,
  TESTRAIL_API_KEY,
  TESTRAIL_PROJECT_ID,
  PORT = 3000
} = process.env;

const TESTRAIL_API = `${TESTRAIL_BASE_URL}/index.php?/api/v2`;

// Known user ID -> name mapping
const USER_MAP = {
  896: 'achnna',
  778: 'maniikk',
};

// Test type filters (matched against plan name)
const TEST_TYPES = ['EBAT', 'Regression', 'Sanity', 'UBat', 'Bat'];

// Helper to make authenticated TestRail API requests
async function testrailRequest(endpoint) {
  const url = `${TESTRAIL_API}${endpoint}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      'Authorization': 'Basic ' + Buffer.from(`pandimkp@amazon.com:${TESTRAIL_API_KEY}`).toString('base64'),
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`TestRail API error: ${response.status} - ${errorText}`);
  }

  return response.json();
}

// Convert date string to Unix timestamp
function dateToTimestamp(dateStr, endOfDay = false) {
  const date = new Date(dateStr + 'T00:00:00Z');
  if (endOfDay) {
    date.setUTCHours(23, 59, 59, 999);
  }
  return Math.floor(date.getTime() / 1000);
}

// Get user name by ID
function getUserName(userId) {
  if (!userId) return '';
  if (USER_MAP[userId]) return USER_MAP[userId];
  return `User_${userId}`;
}

// Map priority ID to name
function getPriorityName(priorityId) {
  const priorities = { 1: 'P4', 2: 'P3', 3: 'P2', 4: 'P1', 5: 'P0' };
  return priorities[priorityId] || `P${priorityId}`;
}

function getMonthName(dateStr) {
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const date = new Date(dateStr);
  return months[date.getMonth()];
}

// Fetch paginated data from TestRail
async function fetchPaginated(endpoint, key) {
  let allResults = [];
  let offset = 0;
  const limit = 250;

  while (true) {
    const sep = endpoint.includes('&') ? '&' : '&';
    const data = await testrailRequest(`${endpoint}${sep}limit=${limit}&offset=${offset}`);

    if (Array.isArray(data)) {
      allResults = allResults.concat(data);
      if (data.length < limit) break;
    } else if (data[key] && Array.isArray(data[key])) {
      allResults = allResults.concat(data[key]);
      if (data[key].length < limit) break;
    } else {
      break;
    }
    offset += limit;
  }
  return allResults;
}

// Core audit logic
async function getAuditData(from, to, projectId, testType) {
  const fromTimestamp = dateToTimestamp(from);
  const toTimestamp = dateToTimestamp(to, true);
  const pid = projectId || TESTRAIL_PROJECT_ID;
  const typeFilter = (testType || 'ebat').toLowerCase();

  console.log(`[Audit] Fetching from ${from} to ${to}, project=${pid}, type=${typeFilter}`);

  // 1. Get plans in date range
  const plans = await fetchPaginated(
    `/get_plans/${pid}&created_after=${fromTimestamp}&created_before=${toTimestamp}`,
    'plans'
  );

  // 2. Filter plans by test type (match ":TYPE:" in plan name)
  const ksPlans = plans.filter(p => {
    const name = (p.name || '').toLowerCase();
    return name.includes(`:${typeFilter}:`);
  });

  console.log(`[Audit] ${plans.length} total plans, ${ksPlans.length} matching "${typeFilter}" filter`);

  // Get project name for team column
  let teamName = '';
  try {
    const project = await testrailRequest(`/get_project/${pid}`);
    teamName = project.name || `Project ${pid}`;
  } catch (err) {
    teamName = `Project ${pid}`;
  }

  // 3. Get runs from matching plans
  const planRuns = [];
  for (const plan of ksPlans) {
    try {
      const planDetail = await testrailRequest(`/get_plan/${plan.id}`);
      for (const entry of (planDetail.entries || [])) {
        for (const run of (entry.runs || [])) {
          planRuns.push({ run, planName: plan.name });
        }
      }
    } catch (err) {
      console.error(`[Audit] Error fetching plan ${plan.id}: ${err.message}`);
    }
  }

  console.log(`[Audit] ${planRuns.length} runs to process`);

  // 4. Process each run sequentially
  const auditRows = [];
  const seenUserIds = new Set();

  for (let i = 0; i < planRuns.length; i++) {
    const { run, planName } = planRuns[i];
    try {
      // Get tests
      const tests = await fetchPaginated(`/get_tests/${run.id}`, 'tests');
      if (tests.length === 0) continue;

      // Get results (only in date range)
      const results = await fetchPaginated(
        `/get_results_for_run/${run.id}&created_after=${fromTimestamp}&created_before=${toTimestamp}`,
        'results'
      );
      if (results.length === 0) continue;

      // Filter: only passed, no automation comments
      const resultsByTest = {};
      for (const result of results) {
        if (result.status_id !== 1) continue;
        const comment = (result.comment || '').toLowerCase();
        if (comment.includes('based on automation result')) continue;

        if (!resultsByTest[result.test_id] || result.created_on > resultsByTest[result.test_id].created_on) {
          resultsByTest[result.test_id] = result;
        }
      }

      const matchCount = Object.keys(resultsByTest).length;
      if (matchCount > 0) {
        console.log(`[Audit] Run ${run.id} (${i + 1}/${planRuns.length}): ${matchCount} passed manual results`);
      }

      // Build rows
      for (const testId of Object.keys(resultsByTest)) {
        const result = resultsByTest[testId];
        const test = tests.find(t => t.id === parseInt(testId));
        if (!test) continue;

        const userId = result.created_by || test.assignedto_id;
        seenUserIds.add(userId);

        auditRows.push({
          da_user_id: getUserName(userId),
          priority: getPriorityName(test.priority_id),
          test_case_id: `T${test.case_id}`,
          test_plan_name: planName,
          team: teamName,
          ambiguity_1a: 'Yes',
          ambiguity_1b: 'Yes',
          ambiguity_2a: 'Yes',
          ambiguity_2b: 'Yes',
          execution_check: 'OK',
          comments: result.comment || 'Passed the TC with proper Build number',
          audit_date: from,
          audited_by: 'pandimkp',
          dispute: 'NO',
          da_manager: 'priyadp',
          month: getMonthName(from),
          audit_type: 'Ambiguity',
          test_id: test.id,
          run_id: run.id,
          test_url: `${TESTRAIL_BASE_URL}/index.php?/tests/view/${test.id}`
        });
      }
    } catch (err) {
      console.error(`[Audit] Error on run ${run.id}: ${err.message}`);
    }
  }

  console.log(`[Audit] Done: ${auditRows.length} rows. User IDs seen: ${[...seenUserIds].join(', ')}`);

  return {
    project_id: TESTRAIL_PROJECT_ID,
    date_range: { from, to },
    total_rows: auditRows.length,
    total_runs: planRuns.length,
    total_plans: ksPlans.length,
    user_ids_seen: [...seenUserIds],
    rows: auditRows
  };
}

// GET /api/audit
app.get('/api/audit', async (req, res) => {
  try {
    const { from, to, project, testType } = req.query;
    if (!from || !to) {
      return res.status(400).json({ error: 'Both "from" and "to" dates required (YYYY-MM-DD)' });
    }
    const data = await getAuditData(from, to, project, testType);
    res.json(data);
  } catch (error) {
    console.error('[Audit] Error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/export
app.get('/api/export', async (req, res) => {
  try {
    const { from, to, project, testType } = req.query;
    if (!from || !to) {
      return res.status(400).json({ error: 'Both "from" and "to" dates required' });
    }

    const data = await getAuditData(from, to, project, testType);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Sheet1');

    sheet.columns = [
      { header: 'DA User ID', key: 'da_user_id', width: 15 },
      { header: 'Priority', key: 'priority', width: 10 },
      { header: 'Test case id (if applicable )', key: 'test_case_id', width: 18 },
      { header: 'Test plan name (if applicable )', key: 'test_plan_name', width: 45 },
      { header: 'Team', key: 'team', width: 18 },
      { header: 'Ambiguity-1A', key: 'ambiguity_1a', width: 14 },
      { header: 'Ambiguity-1B', key: 'ambiguity_1b', width: 14 },
      { header: 'Ambiguity-2A', key: 'ambiguity_2a', width: 14 },
      { header: 'Ambiguity-2B', key: 'ambiguity_2b', width: 14 },
      { header: 'Execution Check', key: 'execution_check', width: 16 },
      { header: 'Comments(provide detailed description)', key: 'comments', width: 40 },
      { header: 'Audit date', key: 'audit_date', width: 12 },
      { header: 'Auditted by', key: 'audited_by', width: 14 },
      { header: 'Dispute', key: 'dispute', width: 10 },
      { header: 'DA Manager', key: 'da_manager', width: 14 },
      { header: 'Month', key: 'month', width: 10 },
      { header: 'Audit type', key: 'audit_type', width: 12 }
    ];

    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E2F3' } };

    for (const row of data.rows) {
      sheet.addRow(row);
    }

    const filename = `TestRail_Audit_${from}_to_${to}.xlsx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    await workbook.xlsx.write(res);
    res.end();
  } catch (error) {
    console.error('[Export] Error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/projects - Get list of projects (teams)
app.get('/api/projects', async (req, res) => {
  try {
    const data = await testrailRequest('/get_projects&is_completed=0');
    const projects = Array.isArray(data) ? data : (data.projects || []);
    res.json(projects.map(p => ({ id: p.id, name: p.name })));
  } catch (error) {
    console.error('[Projects] Error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// GET /api/test-types - Get available test types
app.get('/api/test-types', (req, res) => {
  res.json(TEST_TYPES);
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`\n  TestRail Audit Tool running at http://localhost:${PORT}`);
  console.log(`  Project ID: ${TESTRAIL_PROJECT_ID}`);
  console.log(`  TestRail: ${TESTRAIL_BASE_URL}\n`);
});
