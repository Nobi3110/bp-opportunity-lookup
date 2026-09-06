/**
 * Google Apps Script backend for BP Name / Opportunity ID / BP Code lookup.
 *
 * WITH ADMIN & EMPLOYEE AUTHENTICATION
 *
 * SETUP:
 * 1. Open your Google Sheet with the customer data.
 * 2. Go to Extensions -> Apps Script.
 * 3. Delete any starter code, paste this whole file in.
 * 4. Update SHEET_NAME / BP_CODE_COLUMN / BP_NAME_COLUMN / OPPORTUNITY_ID_COLUMN
 *    below to match your actual column headers exactly.
 * 5. Change ADMIN_PASSWORD below to something private before deploying.
 * 6. Deploy -> Manage deployments -> edit -> Version: New version -> Deploy.
 *    (Or Deploy -> New deployment the first time.)
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 7. Copy the Web app URL -- paste it into API_URL in both index.html and auth.html.
 *
 * ENDPOINTS:
 *   ?suggest=partial-text   -> returns suggestions for dropdown
 *   ?id=exact-bp-name-or-opportunity-id-or-bp-code -> returns the single full matching record
 *   ?login=1&username=X&password=Y -> authenticate user
 *   ?action=create_user&username=X&password=Y&role=admin|employee -> create new user (admin only)
 *   ?action=list_users -> get all users (admin only)
 *   ?action=delete_user&username=X -> delete user (admin only)
 */

// ---- CONFIG: change these to match your sheet/column names ----
var SHEET_NAME = "06 Sept. 2026";               // name of the tab with customer records — update if you rename the tab
var BP_CODE_COLUMN = "BP Code";
var BP_NAME_COLUMN = "BP Name";
var OPPORTUNITY_ID_COLUMN = "Opportunity ID";
var SUGGEST_LIMIT = 15; // max rows returned by the suggest endpoint
var ADMIN_USERNAME = "admin";
var ADMIN_PASSWORD = "#m0t0r0L@$";
var USERS_PROPERTY_KEY = "BP_LOOKUP_USERS";
// -----------------------------------------------------------------

function normalize_(s) {
  return String(s).trim().replace(/\s+/g, " ").toLowerCase();
}

function findColumnIndex_(headers, targetName) {
  var normalizedTarget = normalize_(targetName);
  for (var h = 0; h < headers.length; h++) {
    if (normalize_(headers[h]) === normalizedTarget) {
      return h;
    }
  }
  return -1;
}

function getSheetData_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // Fallback: SHEET_NAME didn't match any tab (e.g. the tab was renamed).
    // Use the first tab in the spreadsheet instead of failing outright.
    sheet = ss.getSheets()[0];
  }
  if (!sheet) {
    throw new Error("No sheets found in this spreadsheet.");
  }
  return sheet.getDataRange().getValues();
}

function doGet(e) {
  var output;
  try {
    if (e.parameter.login !== undefined) {
      output = handleLogin_(e.parameter.username || "", e.parameter.password || "");
    } else if (e.parameter.action !== undefined) {
      output = handleAdminAction_(e.parameter);
    } else if (e.parameter.suggest !== undefined) {
      output = handleSuggest_(e.parameter.suggest || "");
    } else {
      output = handleExactLookup_(e.parameter.id || "");
    }
  } catch (err) {
    output = { error: err.message };
  }

  return ContentService
    .createTextOutput(JSON.stringify(output))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Returns up to SUGGEST_LIMIT rows whose BP Name, Opportunity ID, or BP Code
 * contains the given text (case-insensitive "contains" match), for
 * populating a dropdown while the user is still typing.
 */
function handleSuggest_(query) {
  var q = normalize_(query);
  if (!q) {
    return { matches: [] };
  }

  var data = getSheetData_();
  var headers = data[0];
  var bpCodeIdx = findColumnIndex_(headers, BP_CODE_COLUMN);
  var bpNameIdx = findColumnIndex_(headers, BP_NAME_COLUMN);
  var oppIdIdx = findColumnIndex_(headers, OPPORTUNITY_ID_COLUMN);

  if (bpCodeIdx === -1 || bpNameIdx === -1 || oppIdIdx === -1) {
    return {
      error: "Column not found. Actual headers: " + headers.join(" | "),
      matches: []
    };
  }

  var matches = [];
  for (var i = 1; i < data.length && matches.length < SUGGEST_LIMIT; i++) {
    var bpCodeVal = String(data[i][bpCodeIdx]).trim();
    var bpNameVal = String(data[i][bpNameIdx]).trim();
    var oppIdVal = String(data[i][oppIdIdx]).trim();
    if (
      normalize_(bpCodeVal).indexOf(q) !== -1 ||
      normalize_(bpNameVal).indexOf(q) !== -1 ||
      normalize_(oppIdVal).indexOf(q) !== -1
    ) {
      matches.push({ bpCode: bpCodeVal, bpName: bpNameVal, opportunityId: oppIdVal });
    }
  }

  return { matches: matches };
}

/**
 * Returns the single full record whose BP Name, Opportunity ID, or BP Code
 * exactly matches (case-insensitive) the given value.
 */
function handleExactLookup_(searchId) {
  searchId = String(searchId).trim();
  if (!searchId) {
    return { found: false, error: "No search value provided" };
  }

  var data = getSheetData_();
  var headers = data[0];
  var bpCodeIdx = findColumnIndex_(headers, BP_CODE_COLUMN);
  var bpNameIdx = findColumnIndex_(headers, BP_NAME_COLUMN);
  var oppIdIdx = findColumnIndex_(headers, OPPORTUNITY_ID_COLUMN);

  if (bpCodeIdx === -1 || bpNameIdx === -1 || oppIdIdx === -1) {
    return {
      found: false,
      error: "Required search column not found. Actual headers: " + headers.join(" | ")
    };
  }

  for (var i = 1; i < data.length; i++) {
    var bpCodeVal = String(data[i][bpCodeIdx]).trim();
    var bpNameVal = String(data[i][bpNameIdx]).trim();
    var oppIdVal = String(data[i][oppIdIdx]).trim();
    if (
      bpCodeVal.toLowerCase() === searchId.toLowerCase() ||
      bpNameVal.toLowerCase() === searchId.toLowerCase() ||
      oppIdVal.toLowerCase() === searchId.toLowerCase()
    ) {
      var record = {};
      for (var j = 0; j < headers.length; j++) {
        var key = String(headers[j]).trim();
        if (key) {
          record[key] = data[i][j];
        }
      }
      return { found: true, data: record };
    }
  }

  return { found: false };
}

// ============ AUTHENTICATION FUNCTIONS ============

function getUsers_() {
  var properties = PropertiesService.getScriptProperties();
  var storedUsers = properties.getProperty(USERS_PROPERTY_KEY);
  if (storedUsers) {
    return JSON.parse(storedUsers);
  }

  var users = [];
  if (!users.some(function(user) { return user.username === ADMIN_USERNAME; })) {
    users.push({
      username: ADMIN_USERNAME,
      password: ADMIN_PASSWORD,
      role: "admin",
      status: "active",
      created: new Date().toISOString()
    });
  }
  properties.setProperty(USERS_PROPERTY_KEY, JSON.stringify(users));
  return users;
}

function saveUsers_(users) {
  PropertiesService.getScriptProperties().setProperty(USERS_PROPERTY_KEY, JSON.stringify(users));
}

function handleLogin_(username, password) {
  var users = getUsers_();
  username = String(username).trim().toLowerCase();
  for (var i = 0; i < users.length; i++) {
    var user = users[i];
    if (user.username === username && user.password === password && user.status === "active") {
      return {
        success: true,
        username: username,
        role: user.role,
        message: "Login successful"
      };
    }
  }
  return { success: false, error: "Invalid username or password" };
}

function handleAdminAction_(params) {
  var action = params.action;
  var username = (params.username || "").toString().trim().toLowerCase();
  var password = (params.password || "").toString().trim();
  var role = (params.role || "employee").toString().trim().toLowerCase();
  var adminUser = (params.adminUser || "").toString().trim().toLowerCase();
  var adminPass = (params.adminPass || "").toString().trim();

  var loginResult = handleLogin_(adminUser, adminPass);
  if (!loginResult.success || loginResult.role !== "admin") {
    return { success: false, error: "Unauthorized. Admin credentials required." };
  }

  if (action === "create_user") {
    return createUser_(username, password, role);
  } else if (action === "list_users") {
    return listUsers_();
  } else if (action === "delete_user") {
    return deleteUser_(username);
  } else if (action === "update_user") {
    return updateUser_(username, password, role);
  }

  return { success: false, error: "Unknown action" };
}

function createUser_(username, password, role) {
  if (!username || !password) {
    return { success: false, error: "Username and password required" };
  }
  if (role !== "admin" && role !== "employee") {
    return { success: false, error: "Role must be 'admin' or 'employee'" };
  }

  var users = getUsers_();
  for (var i = 0; i < users.length; i++) {
    if (users[i].username === username) {
      return { success: false, error: "User already exists" };
    }
  }

  users.push({ username: username, password: password, role: role, status: "active", created: new Date().toISOString() });
  saveUsers_(users);
  return {
    success: true,
    message: "User '" + username + "' created successfully as " + role
  };
}

function listUsers_() {
  var users = [];
  var storedUsers = getUsers_();
  for (var i = 0; i < storedUsers.length; i++) {
    if (storedUsers[i].username) {
      users.push({
        username: storedUsers[i].username,
        role: storedUsers[i].role,
        status: storedUsers[i].status,
        created: storedUsers[i].created
      });
    }
  }
  return { success: true, users: users };
}

function deleteUser_(username) {
  if (username === "admin") {
    return { success: false, error: "Cannot delete default admin user" };
  }
  var users = getUsers_();
  for (var i = 0; i < users.length; i++) {
    if (users[i].username === username) {
      users.splice(i, 1);
      saveUsers_(users);
      return { success: true, message: "User '" + username + "' deleted" };
    }
  }
  return { success: false, error: "User not found" };
}

function updateUser_(username, password, role) {
  var users = getUsers_();
  for (var i = 0; i < users.length; i++) {
    if (users[i].username === username) {
      if (password) {
        users[i].password = password;
      }
      if (role) {
        users[i].role = role;
      }
      saveUsers_(users);
      return { success: true, message: "User '" + username + "' updated" };
    }
  }
  return { success: false, error: "User not found" };
}