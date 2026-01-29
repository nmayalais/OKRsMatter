const APP_CONFIG = {
  SPREADSHEET_NAME: 'OKRsMatter Data',
  SHEETS: {
    OBJECTIVES: 'Objectives',
    KEY_RESULTS: 'KeyResults',
    USERS: 'Users',
    DEPARTMENTS: 'Departments',
    TEAMS: 'Teams',
    QUARTERS: 'Quarters',
    CHECK_INS: 'CheckIns',
    COACH_LOGS: 'CoachLogs'
  },
  SETTINGS_DEFAULTS: {
    KR_SOFT_CAP: 5,
    EXEC_OBJECTIVE_SOFT_CAP: 5,
    TEAM_OBJECTIVE_SOFT_CAP: 3
  }
};

function getBootstrapData(options) {
  const user = Session.getActiveUser();
  const email = user.getEmail();
  const profile = getOrCreateUserProfile(email);
  const settings = options || {};
  const includeKeyResults = settings.includeKeyResults !== false;
  const includeCheckIns = settings.includeCheckIns !== false;
  return {
    user: profile,
    aiConfig: getAiConfig_(),
    canManageAiSettings: canManageAiSettings_(profile),
    canManageAdmin: canManageAdmin_(profile),
    appSettings: getAppSettings_(),
    objectives: listObjectives(),
    keyResults: includeKeyResults ? listKeyResults() : [],
    checkIns: includeCheckIns ? listCheckIns() : [],
    quarters: listQuarters(),
    departments: listDepartments(),
    teams: listTeams()
  };
}

function getBootstrapDataRaw(options) {
  return JSON.stringify(getBootstrapData(options));
}

function setApiKeys(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAiSettings_(profile)) {
    throw new Error('You do not have permission to manage AI settings.');
  }
  const props = PropertiesService.getScriptProperties();
  if (payload && payload.openaiKey) {
    props.setProperty('OPENAI_API_KEY', payload.openaiKey);
  }
  if (payload && payload.anthropicKey) {
    props.setProperty('ANTHROPIC_API_KEY', payload.anthropicKey);
  }
  if (payload && payload.openaiModel) {
    props.setProperty('OPENAI_MODEL', payload.openaiModel);
  }
  if (payload && payload.anthropicModel) {
    props.setProperty('ANTHROPIC_MODEL', payload.anthropicModel);
  }
  return getAiConfig_();
}

function getAiConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    hasOpenAiKey: !!props.getProperty('OPENAI_API_KEY'),
    hasAnthropicKey: !!props.getProperty('ANTHROPIC_API_KEY'),
    openaiModel: props.getProperty('OPENAI_MODEL') || 'gpt-4o-mini',
    anthropicModel: props.getProperty('ANTHROPIC_MODEL') || 'claude-3-5-sonnet-20240620'
  };
}

function canManageAiSettings_(profile) {
  const role = normalizeRole_(profile && profile.role);
  return role === 'exec' || role === 'operations';
}

function normalizeRole_(role) {
  return (role || '').toString().trim().toLowerCase();
}

function canManageAdmin_(profile) {
  const role = normalizeRole_(profile && profile.role);
  return role === 'exec' || role === 'operations';
}

function getAppSettings_() {
  const props = PropertiesService.getScriptProperties();
  return {
    krSoftCap: Number(props.getProperty('KR_SOFT_CAP')) || APP_CONFIG.SETTINGS_DEFAULTS.KR_SOFT_CAP,
    execObjectiveSoftCap: Number(props.getProperty('EXEC_OBJECTIVE_SOFT_CAP')) || APP_CONFIG.SETTINGS_DEFAULTS.EXEC_OBJECTIVE_SOFT_CAP,
    teamObjectiveSoftCap: Number(props.getProperty('TEAM_OBJECTIVE_SOFT_CAP')) || APP_CONFIG.SETTINGS_DEFAULTS.TEAM_OBJECTIVE_SOFT_CAP
  };
}

function updateAppSettings(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage app settings.');
  }
  const props = PropertiesService.getScriptProperties();
  if (payload && payload.krSoftCap) {
    props.setProperty('KR_SOFT_CAP', String(payload.krSoftCap));
  }
  if (payload && payload.execObjectiveSoftCap) {
    props.setProperty('EXEC_OBJECTIVE_SOFT_CAP', String(payload.execObjectiveSoftCap));
  }
  if (payload && payload.teamObjectiveSoftCap) {
    props.setProperty('TEAM_OBJECTIVE_SOFT_CAP', String(payload.teamObjectiveSoftCap));
  }
  return getAppSettings_();
}

function getOrCreateUserProfile(email) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS, [
    'email', 'name', 'role', 'department', 'team', 'createdAt', 'updatedAt'
  ]);
  const values = getSheetValues_(sheet);
  const rows = values.rows;
  const existing = rows.find(row => row.email === email);
  const now = new Date().toISOString();

  if (existing) {
    return existing;
  }

  const profile = {
    email,
    name: email.split('@')[0],
    role: 'member',
    department: 'Unassigned',
    team: 'Unassigned',
    createdAt: now,
    updatedAt: now
  };

  appendRow_(sheet, values.headers, profile);
  return profile;
}

function listUsers() {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to view users.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS, [
    'email', 'name', 'role', 'department', 'team', 'createdAt', 'updatedAt'
  ]);
  return getSheetValues_(sheet).rows;
}

function updateUser(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage users.');
  }
  if (!payload || !payload.email) {
    throw new Error('User email is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS, [
    'email', 'name', 'role', 'department', 'team', 'createdAt', 'updatedAt'
  ]);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.email === payload.email);
  const now = new Date().toISOString();
  if (rowIndex === -1) {
    const record = Object.assign({
      email: payload.email,
      name: payload.name || payload.email.split('@')[0],
      role: payload.role || 'member',
      department: payload.department || 'Unassigned',
      team: payload.team || 'Unassigned',
      createdAt: now,
      updatedAt: now
    });
    appendRow_(sheet, values.headers, record);
    return record;
  }
  const updated = Object.assign({}, values.rows[rowIndex], payload, { updatedAt: now });
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function deleteUser(email) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage users.');
  }
  if (!email) {
    throw new Error('User email is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.USERS, [
    'email', 'name', 'role', 'department', 'team', 'createdAt', 'updatedAt'
  ]);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.email === email);
  if (rowIndex === -1) {
    throw new Error('User not found.');
  }
  sheet.deleteRow(rowIndex + 2);
  return { ok: true };
}

function listObjectives(filters) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.OBJECTIVES, [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
  ]);

  let rows = getSheetValues_(sheet).rows;

  if (filters) {
    if (filters.level) {
      rows = rows.filter(row => row.level === filters.level);
    }
    if (filters.quarter) {
      rows = rows.filter(row => row.quarter === filters.quarter);
    }
    if (filters.department) {
      rows = rows.filter(row => row.department === filters.department);
    }
    if (filters.team) {
      rows = rows.filter(row => row.team === filters.team);
    }
  }

  return rows;
}

function createObjective(payload) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.OBJECTIVES, [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
  ]);
  const values = getSheetValues_(sheet);
  const now = new Date().toISOString();
  validateObjectivePayload_(payload);
  const record = Object.assign({}, payload, {
    id: Utilities.getUuid(),
    createdAt: now,
    updatedAt: now
  });
  appendRow_(sheet, values.headers, record);
  return record;
}

function updateObjective(payload) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.OBJECTIVES, [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
  ]);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.id === payload.id);
  if (rowIndex === -1) {
    throw new Error('Objective not found.');
  }
  const now = new Date().toISOString();
  const updated = Object.assign({}, values.rows[rowIndex], payload, { updatedAt: now });
  validateObjectivePayload_(updated);
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function listKeyResults(objectiveId) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.KEY_RESULTS, [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
  ]);
  let rows = getSheetValues_(sheet).rows;
  if (objectiveId) {
    rows = rows.filter(row => row.objectiveId === objectiveId);
  }
  return rows;
}

function createKeyResult(payload) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.KEY_RESULTS, [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
  ]);
  const values = getSheetValues_(sheet);
  const now = new Date().toISOString();
  validateKeyResultPayload_(payload);
  const record = Object.assign({}, payload, {
    id: Utilities.getUuid(),
    createdAt: now,
    updatedAt: now
  });
  appendRow_(sheet, values.headers, record);
  return record;
}

function updateKeyResult(payload) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.KEY_RESULTS, [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
  ]);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.id === payload.id);
  if (rowIndex === -1) {
    throw new Error('Key Result not found.');
  }
  const now = new Date().toISOString();
  const updated = Object.assign({}, values.rows[rowIndex], payload, { updatedAt: now });
  validateKeyResultPayload_(updated);
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function listDepartments() {
  const sheet = getSheet_(APP_CONFIG.SHEETS.DEPARTMENTS, ['name', 'leadEmail']);
  return getSheetValues_(sheet).rows;
}

function upsertDepartment(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage departments.');
  }
  if (!payload || !payload.name) {
    throw new Error('Department name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.DEPARTMENTS, ['name', 'leadEmail']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === payload.name);
  if (rowIndex === -1) {
    appendRow_(sheet, values.headers, payload);
    return payload;
  }
  const updated = Object.assign({}, values.rows[rowIndex], payload);
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function deleteDepartment(name) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage departments.');
  }
  if (!name) {
    throw new Error('Department name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.DEPARTMENTS, ['name', 'leadEmail']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === name);
  if (rowIndex === -1) {
    throw new Error('Department not found.');
  }
  sheet.deleteRow(rowIndex + 2);
  return { ok: true };
}

function listTeams() {
  const sheet = getSheet_(APP_CONFIG.SHEETS.TEAMS, ['name', 'department', 'leadEmail']);
  return getSheetValues_(sheet).rows;
}

function upsertTeam(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage teams.');
  }
  if (!payload || !payload.name) {
    throw new Error('Team name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.TEAMS, ['name', 'department', 'leadEmail']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === payload.name);
  if (rowIndex === -1) {
    appendRow_(sheet, values.headers, payload);
    return payload;
  }
  const updated = Object.assign({}, values.rows[rowIndex], payload);
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function deleteTeam(name) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage teams.');
  }
  if (!name) {
    throw new Error('Team name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.TEAMS, ['name', 'department', 'leadEmail']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === name);
  if (rowIndex === -1) {
    throw new Error('Team not found.');
  }
  sheet.deleteRow(rowIndex + 2);
  return { ok: true };
}

function listQuarters() {
  const sheet = getSheet_(APP_CONFIG.SHEETS.QUARTERS, ['name', 'start', 'end', 'status']);
  const rows = getSheetValues_(sheet).rows;
  if (rows.length) {
    return rows;
  }

  // Pre-populate the current year quarters on first run.
  const year = new Date().getFullYear();
  const defaults = [
    { name: `Q1 ${year}`, start: `${year}-01-01`, end: `${year}-03-31`, status: 'active' },
    { name: `Q2 ${year}`, start: `${year}-04-01`, end: `${year}-06-30`, status: 'upcoming' },
    { name: `Q3 ${year}`, start: `${year}-07-01`, end: `${year}-09-30`, status: 'upcoming' },
    { name: `Q4 ${year}`, start: `${year}-10-01`, end: `${year}-12-31`, status: 'upcoming' }
  ];
  defaults.forEach(item => appendRow_(sheet, ['name', 'start', 'end', 'status'], item));
  return defaults;
}

function upsertQuarter(payload) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage quarters.');
  }
  if (!payload || !payload.name) {
    throw new Error('Quarter name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.QUARTERS, ['name', 'start', 'end', 'status']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === payload.name);
  if (rowIndex === -1) {
    appendRow_(sheet, values.headers, payload);
    return payload;
  }
  const updated = Object.assign({}, values.rows[rowIndex], payload);
  writeRow_(sheet, values.headers, rowIndex + 2, updated);
  return updated;
}

function deleteQuarter(name) {
  const profile = getOrCreateUserProfile(Session.getActiveUser().getEmail());
  if (!canManageAdmin_(profile)) {
    throw new Error('You do not have permission to manage quarters.');
  }
  if (!name) {
    throw new Error('Quarter name is required.');
  }
  const sheet = getSheet_(APP_CONFIG.SHEETS.QUARTERS, ['name', 'start', 'end', 'status']);
  const values = getSheetValues_(sheet);
  const rowIndex = values.rows.findIndex(row => row.name === name);
  if (rowIndex === -1) {
    throw new Error('Quarter not found.');
  }
  sheet.deleteRow(rowIndex + 2);
  return { ok: true };
}

function listCheckIns(krId) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.CHECK_INS, [
    'id', 'objectiveId', 'keyResultId', 'current', 'confidence', 'notes', 'weekStart', 'ownerEmail', 'createdAt'
  ]);
  let rows = getSheetValues_(sheet).rows;
  if (krId) {
    rows = rows.filter(row => row.keyResultId === krId);
  }
  return rows;
}

function createCheckIn(payload) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.CHECK_INS, [
    'id', 'objectiveId', 'keyResultId', 'current', 'confidence', 'notes', 'weekStart', 'ownerEmail', 'createdAt'
  ]);
  const values = getSheetValues_(sheet);
  const now = new Date().toISOString();
  validateCheckInPayload_(payload);
  const record = Object.assign({}, payload, {
    id: Utilities.getUuid(),
    createdAt: now
  });
  appendRow_(sheet, values.headers, record);

  if (payload.keyResultId) {
    const updated = updateKeyResult({
      id: payload.keyResultId,
      current: payload.current,
      confidence: payload.confidence
    });
    return { checkIn: record, keyResult: updated };
  }

  return { checkIn: record };
}

function setupApp() {
  const spreadsheet = getOrCreateSpreadsheet_();
  Object.keys(APP_CONFIG.SHEETS).forEach(key => {
    getSheet_(APP_CONFIG.SHEETS[key]);
  });
  return { spreadsheetId: spreadsheet.getId(), url: spreadsheet.getUrl() };
}

function getOrCreateSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  const storedId = props.getProperty('OKR_SPREADSHEET_ID');
  if (storedId) {
    try {
      return SpreadsheetApp.openById(storedId);
    } catch (error) {
      // Fall through and recreate.
    }
  }

  const files = DriveApp.getFilesByName(APP_CONFIG.SPREADSHEET_NAME);
  if (files.hasNext()) {
    const file = files.next();
    props.setProperty('OKR_SPREADSHEET_ID', file.getId());
    return SpreadsheetApp.openById(file.getId());
  }

  const spreadsheet = SpreadsheetApp.create(APP_CONFIG.SPREADSHEET_NAME);
  props.setProperty('OKR_SPREADSHEET_ID', spreadsheet.getId());
  return spreadsheet;
}

function getSheet_(name, headers) {
  const spreadsheet = getOrCreateSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
  }
  if (headers && sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (headers) {
    ensureHeaders_(sheet, headers);
  }
  return sheet;
}

function getSheetValues_(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    return { headers: [], rows: [] };
  }
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = row[index] === undefined ? '' : row[index];
    });
    return record;
  });
  return { headers, rows };
}

function appendRow_(sheet, headers, record) {
  const row = headers.map(header => record[header] || '');
  sheet.appendRow(row);
}

function writeRow_(sheet, headers, rowIndex, record) {
  const row = headers.map(header => record[header] || '');
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function ensureHeaders_(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const missing = headers.filter(header => existing.indexOf(header) === -1);
  if (!missing.length) return;
  const startCol = existing.length + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
}

function validateObjectivePayload_(payload) {
  if (!payload.title) {
    throw new Error('Objective title is required.');
  }
  if (!payload.level) {
    throw new Error('Objective level is required.');
  }
  if (!payload.quarter) {
    throw new Error('Objective quarter is required.');
  }
  if (!payload.description) {
    throw new Error('Objective description is required.');
  }
  if (!payload.rationale) {
    throw new Error('Objective rationale is required.');
  }
  if (payload.level === 'Department' && !payload.parentId) {
    throw new Error('Department objectives must map to an executive objective.');
  }
  if (payload.level === 'Team' && !payload.parentId) {
    throw new Error('Team objectives must map to a department objective.');
  }
}

function validateKeyResultPayload_(payload) {
  if (!payload.objectiveId) {
    throw new Error('Key Result must be linked to an objective.');
  }
  if (!payload.title) {
    throw new Error('Key Result title is required.');
  }
  if (!payload.metric) {
    throw new Error('Key Result metric is required.');
  }
  if (!payload.baseline && payload.baseline !== 0) {
    throw new Error('Key Result baseline is required.');
  }
  if (!payload.target && payload.target !== 0) {
    throw new Error('Key Result target is required.');
  }
  if (!payload.timeline) {
    throw new Error('Key Result timeline is required.');
  }
}

function validateCheckInPayload_(payload) {
  if (!payload.keyResultId) {
    throw new Error('Check-in must be linked to a key result.');
  }
  if (!payload.weekStart) {
    throw new Error('Check-in week start is required.');
  }
  if (!payload.current && payload.current !== 0) {
    throw new Error('Check-in current value is required.');
  }
  if (!payload.confidence) {
    throw new Error('Check-in confidence is required.');
  }
}

function coachObjective(payload) {
  const requesterEmail = Session.getActiveUser().getEmail();
  let objective = null;
  let keyResults = [];
  let latestCheckIns = [];
  let prompt = '';
  let result = null;
  let provider = payload && payload.provider ? payload.provider : '';
  let model = payload && payload.model ? payload.model : '';

  try {
    if (!payload || !payload.objectiveId) {
      throw new Error('Objective ID is required for coaching.');
    }
    if (!provider) {
      throw new Error('Provider is required for coaching.');
    }
    objective = listObjectives().find(obj => obj.id === payload.objectiveId);
    if (!objective) {
      throw new Error('Objective not found.');
    }
    keyResults = listKeyResults(objective.id);
    const checkIns = listCheckIns();
    latestCheckIns = keyResults.map(kr => {
      const latest = checkIns.reduce((acc, item) => {
        if (item.keyResultId !== kr.id) return acc;
        if (!acc) return item;
        return (item.createdAt || '') > (acc.createdAt || '') ? item : acc;
      }, null);
      return latest
        ? {
            keyResultId: kr.id,
            current: latest.current,
            confidence: latest.confidence,
            weekStart: latest.weekStart
          }
        : null;
    }).filter(Boolean);

    prompt = buildCoachPrompt_(objective, keyResults, latestCheckIns);
    if (provider === 'openai') {
      result = callOpenAiCoach_(prompt, model);
    } else if (provider === 'anthropic') {
      result = callAnthropicCoach_(prompt, model);
    } else {
      throw new Error('Unsupported provider.');
    }
  } catch (error) {
    result = { ok: false, error: error.message || String(error) };
  }

  logCoachEvent_(requesterEmail, provider, model, objective, keyResults, result);
  return result;
}

function callOpenAiCoach_(prompt, modelOverride) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'OPENAI_API_KEY is not set in Script Properties.' };
  }
  const model = modelOverride || props.getProperty('OPENAI_MODEL') || 'gpt-4o-mini';
  const payload = {
    model,
    input: prompt,
    temperature: 0.2
  };
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${apiKey}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return parseAiResponse_('openai', response, model);
}

function callAnthropicCoach_(prompt, modelOverride) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) {
    return { ok: false, error: 'ANTHROPIC_API_KEY is not set in Script Properties.' };
  }
  const model = modelOverride || props.getProperty('ANTHROPIC_MODEL') || 'claude-3-5-sonnet-20240620';
  const payload = {
    model,
    max_tokens: 700,
    temperature: 0.2,
    messages: [
      { role: 'user', content: prompt }
    ]
  };
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return parseAiResponse_('anthropic', response, model);
}

function parseAiResponse_(provider, response, model) {
  const status = response.getResponseCode();
  const text = response.getContentText();
  let data;
  try {
    data = JSON.parse(text);
  } catch (error) {
    return { ok: false, error: `Invalid JSON response (${status}).`, raw: text };
  }

  if (status < 200 || status >= 300) {
    return { ok: false, error: data.error ? JSON.stringify(data.error) : `Request failed (${status}).`, raw: text };
  }

  if (provider === 'openai') {
    const outputText = extractOpenAiText_(data);
    return { ok: true, provider, model, text: outputText };
  }
  if (provider === 'anthropic') {
    const outputText = extractAnthropicText_(data);
    return { ok: true, provider, model, text: outputText };
  }
  return { ok: false, error: 'Unsupported provider.', raw: text };
}

function extractOpenAiText_(data) {
  if (data.output_text) return data.output_text;
  if (!data.output || !data.output.length) return '';
  const content = data.output
    .map(item => item.content || [])
    .flat()
    .map(part => part.text || '')
    .join('');
  return content;
}

function extractAnthropicText_(data) {
  if (!data.content || !data.content.length) return '';
  return data.content.map(part => part.text || '').join('');
}

function buildCoachPrompt_(objective, keyResults, checkIns) {
  return [
    'You are an OKR coach. Provide concise, actionable coaching as JSON with keys:',
    'summary, clarity_gaps, alignment, prioritization, risks, suggestions, rewritten_objective, kr_rewrites.',
    'Rules:',
    '- summary: 2 sentences max.',
    '- clarity_gaps: list of missing elements (metric/baseline/target/timeline).',
    '- alignment: assess if the objective is mapped properly and if it is outcome-focused.',
    '- prioritization: comment on focus and tradeoffs.',
    '- risks: list top 3 risks.',
    '- suggestions: list 3-5 actionable improvements.',
    '- rewritten_objective: rewrite the objective in an outcome-focused way.',
    '- kr_rewrites: list of rewritten KRs with metric, baseline, target, timeline.',
    'Objective:',
    JSON.stringify(objective),
    'KeyResults:',
    JSON.stringify(keyResults),
    'LatestCheckIns:',
    JSON.stringify(checkIns)
  ].join('\n');
}

function logCoachEvent_(requesterEmail, provider, model, objective, keyResults, result) {
  const sheet = getSheet_(APP_CONFIG.SHEETS.COACH_LOGS, [
    'id', 'objectiveId', 'objectiveTitle', 'requesterEmail', 'provider', 'model',
    'ok', 'error', 'responseText', 'keyResultCount', 'createdAt'
  ]);
  const values = getSheetValues_(sheet);
  const now = new Date().toISOString();
  const record = {
    id: Utilities.getUuid(),
    objectiveId: objective ? objective.id : '',
    objectiveTitle: objective ? objective.title : '',
    requesterEmail,
    provider,
    model: (result && result.model) || model || '',
    ok: result && result.ok ? 'true' : 'false',
    error: result && result.error ? result.error : '',
    responseText: result && result.text ? result.text : '',
    keyResultCount: keyResults ? keyResults.length : 0,
    createdAt: now
  };
  appendRow_(sheet, values.headers, record);
}

function exportRollupReport(filters) {
  const spreadsheet = getOrCreateSpreadsheet_();
  const objectives = listObjectives(filters);
  const keyResults = listKeyResults();
  const checkIns = listCheckIns();
  const weekStart = getWeekStartString_(new Date());

  const departmentRows = buildRollupRows_(objectives, keyResults, checkIns, 'department', weekStart);
  const teamRows = buildRollupRows_(objectives, keyResults, checkIns, 'team', weekStart);

  let sheet = spreadsheet.getSheetByName('Rollup Report');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Rollup Report');
  } else {
    sheet.clearContents();
  }

  const header = ['Group', 'Objectives', 'Alignment %', 'KR Coverage %', 'Check-ins %'];
  let row = 1;
  sheet.getRange(row, 1, 1, header.length).setValues([header]);
  row += 1;
  if (departmentRows.length) {
    sheet.getRange(row, 1, departmentRows.length, header.length).setValues(departmentRows);
    row += departmentRows.length + 2;
  } else {
    row += 1;
  }
  sheet.getRange(row, 1, 1, header.length).setValues([header.map(item => `Team ${item}`)]);
  row += 1;
  if (teamRows.length) {
    sheet.getRange(row, 1, teamRows.length, header.length).setValues(teamRows);
  }

  return { sheetName: sheet.getName(), url: spreadsheet.getUrl() };
}

function buildRollupRows_(objectives, keyResults, checkIns, groupKey, weekStart) {
  const groups = {};
  objectives.forEach(obj => {
    const name = obj[groupKey] || 'Unassigned';
    if (!groups[name]) {
      groups[name] = { name, objectives: [], objectiveIds: new Set() };
    }
    groups[name].objectives.push(obj);
    groups[name].objectiveIds.add(obj.id);
  });

  const checkInsThisWeek = checkIns.filter(item => item.weekStart === weekStart);

  return Object.values(groups).map(group => {
    const objectiveIds = Array.from(group.objectiveIds);
    const groupKrs = keyResults.filter(kr => objectiveIds.indexOf(kr.objectiveId) !== -1);
    const krCounts = objectiveIds.reduce((acc, id) => {
      acc[id] = 0;
      return acc;
    }, {});
    groupKrs.forEach(kr => {
      if (krCounts[kr.objectiveId] !== undefined) {
        krCounts[kr.objectiveId] += 1;
      }
    });
    const objectivesWithKrs = Object.values(krCounts).filter(count => count > 0).length;
    const alignedObjectives = group.objectives.filter(obj => obj.parentId && obj.level !== 'Executive').length;
    const alignEligible = group.objectives.filter(obj => obj.level !== 'Executive').length;
    const krCheckedInSet = checkInsThisWeek.reduce((acc, item) => {
      if (groupKrs.some(kr => kr.id === item.keyResultId)) {
        acc[item.keyResultId] = true;
      }
      return acc;
    }, {});
    const krsCheckedIn = Object.keys(krCheckedInSet).length;

    return [
      group.name,
      group.objectives.length,
      percent_(alignedObjectives, alignEligible),
      percent_(objectivesWithKrs, group.objectives.length),
      percent_(krsCheckedIn, groupKrs.length)
    ];
  }).sort((a, b) => b[1] - a[1]);
}

function getWeekStartString_(date) {
  const now = new Date(date);
  const day = now.getDay();
  const diff = now.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(now);
  monday.setDate(diff);
  return monday.toISOString().slice(0, 10);
}

function percent_(value, total) {
  if (!total) return 0;
  return Math.round((value / total) * 100);
}

function normalizeImportedData() {
  const objectivesSheet = getSheet_(APP_CONFIG.SHEETS.OBJECTIVES, [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
  ]);
  const keyResultsSheet = getSheet_(APP_CONFIG.SHEETS.KEY_RESULTS, [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
  ]);
  const quartersSheet = getSheet_(APP_CONFIG.SHEETS.QUARTERS, ['name', 'start', 'end', 'status']);

  const objectivesValues = getSheetValues_(objectivesSheet);
  const keyResultsValues = getSheetValues_(keyResultsSheet);
  const quarterValues = getSheetValues_(quartersSheet);

  const quartersByName = new Set(quarterValues.rows.map(row => row.name).filter(Boolean));
  const objectiveQuarterById = {};
  let objectivesUpdated = 0;
  let keyResultsUpdated = 0;

  objectivesValues.rows.forEach((row, index) => {
    let updated = false;
    const normalizedQuarter = normalizeQuarterName_(row.quarter);
    if (normalizedQuarter && normalizedQuarter !== row.quarter) {
      row.quarter = normalizedQuarter;
      updated = true;
    }
    const normalizedProgress = normalizeProgressValue_(row.progress);
    if (normalizedProgress !== row.progress) {
      row.progress = normalizedProgress;
      updated = true;
    }
    objectiveQuarterById[row.id] = row.quarter;
    if (row.quarter) {
      quartersByName.add(row.quarter);
    }
    if (updated) {
      objectivesUpdated += 1;
      writeRow_(objectivesSheet, objectivesValues.headers, index + 2, row);
    }
  });

  keyResultsValues.rows.forEach((row, index) => {
    let updated = false;
    let timeline = row.timeline;
    if (!timeline && row.objectiveId && objectiveQuarterById[row.objectiveId]) {
      timeline = objectiveQuarterById[row.objectiveId];
    }
    const normalizedTimeline = normalizeQuarterName_(timeline);
    if (normalizedTimeline && normalizedTimeline !== row.timeline) {
      row.timeline = normalizedTimeline;
      updated = true;
    }
    if (row.timeline) {
      quartersByName.add(row.timeline);
    }
    if (updated) {
      keyResultsUpdated += 1;
      writeRow_(keyResultsSheet, keyResultsValues.headers, index + 2, row);
    }
  });

  const existingQuarterNames = new Set(quarterValues.rows.map(row => row.name).filter(Boolean));
  let quartersAdded = 0;
  quartersByName.forEach(name => {
    if (!existingQuarterNames.has(name)) {
      const range = quarterDateRange_(name);
      if (range) {
        appendRow_(quartersSheet, ['name', 'start', 'end', 'status'], {
          name,
          start: range.start,
          end: range.end,
          status: 'upcoming'
        });
        quartersAdded += 1;
      }
    }
  });

  return { objectivesUpdated, keyResultsUpdated, quartersAdded };
}

function normalizeQuarterName_(value) {
  if (!value) return '';
  const text = value.toString().trim();
  const match = text.match(/Q([1-4])\\s*(?:'| )?(\\d{2}|\\d{4})/i);
  if (!match) return text;
  const quarter = match[1];
  let year = match[2];
  if (year.length === 2) {
    year = `20${year}`;
  }
  return `Q${quarter} ${year}`;
}

function quarterDateRange_(quarterName) {
  const match = quarterName.match(/Q([1-4])\\s+(\\d{4})/i);
  if (!match) return null;
  const quarter = Number(match[1]);
  const year = Number(match[2]);
  const startMonth = (quarter - 1) * 3;
  const start = new Date(year, startMonth, 1);
  const end = new Date(year, startMonth + 3, 0);
  return { start, end };
}

function normalizeProgressValue_(value) {
  if (value === '' || value === null || value === undefined) return value;
  const num = Number(value);
  if (Number.isNaN(num)) return value;
  if (num <= 1) {
    return Math.round(num * 10000) / 100;
  }
  return num;
}

function repairImportedData() {
  const objectivesSheet = getSheet_(APP_CONFIG.SHEETS.OBJECTIVES, [
    'id', 'title', 'level', 'parentId', 'ownerEmail', 'ownerName', 'department', 'team',
    'quarter', 'status', 'priority', 'impact', 'progress', 'description', 'rationale',
    'createdAt', 'updatedAt'
  ]);
  const keyResultsSheet = getSheet_(APP_CONFIG.SHEETS.KEY_RESULTS, [
    'id', 'objectiveId', 'title', 'metric', 'baseline', 'target', 'current',
    'timeline', 'status', 'confidence', 'ownerEmail', 'updatedAt', 'createdAt'
  ]);

  const objectivesValues = getSheetValues_(objectivesSheet);
  const keyResultsValues = getSheetValues_(keyResultsSheet);

  const objectiveById = {};
  const duplicateObjectiveRows = [];
  objectivesValues.rows.forEach((row, index) => {
    if (!row.id) return;
    if (!objectiveById[row.id]) {
      objectiveById[row.id] = { row, index: index + 2 };
      return;
    }

    const existing = objectiveById[row.id].row;
    const existingIsPlaceholder = existing.title === '[Missing Objective]';
    const currentIsPlaceholder = row.title === '[Missing Objective]';
    if (existingIsPlaceholder && !currentIsPlaceholder) {
      duplicateObjectiveRows.push(objectiveById[row.id].index);
      objectiveById[row.id] = { row, index: index + 2 };
    } else {
      duplicateObjectiveRows.push(index + 2);
    }
  });

  if (duplicateObjectiveRows.length) {
    duplicateObjectiveRows.sort((a, b) => b - a).forEach(rowIndex => {
      objectivesSheet.deleteRow(rowIndex);
    });
  }

  const updatedObjectivesValues = getSheetValues_(objectivesSheet);
  const updatedObjectiveById = {};
  updatedObjectivesValues.rows.forEach(row => {
    if (row.id) updatedObjectiveById[row.id] = row;
  });

  const orphanedByObjectiveId = {};
  keyResultsValues.rows.forEach(row => {
    if (!row.objectiveId) return;
    if (updatedObjectiveById[row.objectiveId]) return;
    if (!orphanedByObjectiveId[row.objectiveId]) {
      orphanedByObjectiveId[row.objectiveId] = row;
    }
  });

  const now = new Date().toISOString();
  let placeholdersCreated = 0;
  Object.keys(orphanedByObjectiveId).forEach(objectiveId => {
    const kr = orphanedByObjectiveId[objectiveId];
    const placeholder = {
      id: objectiveId,
      title: '[Missing Objective]',
      level: 'Executive',
      parentId: '',
      ownerEmail: '',
      ownerName: '',
      department: '',
      team: '',
      quarter: kr.timeline || '',
      status: '',
      priority: '',
      impact: '',
      progress: 0,
      description: 'Auto-created to repair orphaned key results.',
      rationale: 'Imported data had key results referencing missing objectives.',
      createdAt: now,
      updatedAt: now
    };
    appendRow_(objectivesSheet, updatedObjectivesValues.headers, placeholder);
    placeholdersCreated += 1;
  });

  let metricsFilled = 0;
  keyResultsValues.rows.forEach((row, index) => {
    if (!row.metric) {
      row.metric = 'value';
      writeRow_(keyResultsSheet, keyResultsValues.headers, index + 2, row);
      metricsFilled += 1;
    }
  });

  const normalized = normalizeImportedData();
  return {
    duplicateObjectivesRemoved: duplicateObjectiveRows.length,
    placeholdersCreated,
    metricsFilled,
    normalized
  };
}
