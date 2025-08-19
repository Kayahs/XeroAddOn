var CLIENT_ID = '';
var CLIENT_SECRET = '';
const REDIRECT_URI = 'https://script.google.com/macros/library/d/1cA_ih-QnoCrWI3z5rB1rkDxEhkSseZo8qy4Qf3E2M0j7FdnRlzPktCsq/2';

function onInstall(e) {
  onOpen(e)
}

function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Show Sidebar', 'showSidebar')
    .addItem('Reset Service', 'reset')
    .addItem('Get Balance Sheet', 'getBalanceSheet')
    .addItem('Get Profit and Loss', 'getProfitAndLoss')
    .addItem('Get Accounts Payable', 'getAccountsPayable')
    .addItem('Get Accounts Receivable', 'getAccountsReceivable')
    .addToUi()
}

function showSidebar() {
  var service = getService(getUserKey());
  var template = HtmlService.createTemplateFromFile("Sidebar.html");
  var date = new Date();
  var currentDate = `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2, '0')}`
  template.currentDate = currentDate;
  if (service.hasAccess()) {
    template.hasAccess = true;
    template.authorizationUrl = null;
  } else {
    template.hasAccess = false;
    template.authorizationUrl = service.getAuthorizationUrl();
  }
  var page = template.evaluate();
  SpreadsheetApp.getUi().showSidebar(page);
}

function closeSidebar() {
  var html = HtmlService.createHtmlOutput("<script>google.script.host.close();</script>");
  SpreadsheetApp.getUi().showSidebar(html);
}

function refreshSidebar() {
  showSidebar();
}

function getUserKey() {
  return Session.getActiveUser().getEmail() || Session.getTemporaryActiveUserKey();
}


function loadReportFromForm(formData) {
  Logger.log("FormData: %s", formData);
  Logger.log("Query String: %s", buildStringFromForm(formData));
  switch(formData.report) {
    case "balanceSheet":
      getBalanceSheet(formData);
      break;
    case "profitLoss":
      getProfitAndLoss(formData);
      break;
    case "accountsPayable":
      getAccountsPayable();
      break;
    case "accountsReceivable":
      getAccountsReceivable();
      break;
    default:
      throw new Error("How did you get here.");
  }
}

/**
 * Reset the authorization state, so that it can be re-tested.
 */
function reset() {
  getService(getUserKey()).reset();
  refreshSidebar();
}

/**
 * Configures the service.
 */
function getService(userKey) {
  Logger.log('Getting Service');
  Logger.log('User Key: %s', userKey);
  return OAuth2.createService('Xero-' + userKey)
    // Set the endpoint URLs.
    .setAuthorizationBaseUrl(
        'https://login.xero.com/identity/connect/authorize')
    .setTokenUrl('https://identity.xero.com/connect/token')

    // Set the client ID and secret.
    .setClientId(CLIENT_ID)
    .setClientSecret(CLIENT_SECRET)

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow.
    .setCallbackFunction('authCallback')

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(PropertiesService.getUserProperties())

    // Set the scopes to request from the user. The scope "offline_access" is
    // required to refresh the token. The full list of scopes is available here:
    // https://developer.xero.com/documentation/oauth2/scopes
    .setScope('accounting.reports.read accounting.settings offline_access')
};

/**
 * Handles the OAuth callback.
 */
function authCallback(request) {
  Logger.log('Auth Callback triggered');
  var service = getService(getUserKey());
  var authorized = service.handleCallback(request);
  if (authorized) {
    // Retrieve the connected tenants.
    var response = UrlFetchApp.fetch('https://api.xero.com/connections', {
      headers: {
        Authorization: 'Bearer ' + service.getAccessToken()
      },
    });
    var connections = JSON.parse(response.getContentText());
    // Store the first tenant ID in the service's storage. If you want to
    // support multiple tenants, store the full list and then let the user
    // select which one to operate against.
    Logger.log(connections[0].tenantId);
    service.getStorage().setValue('tenantId', connections[0].tenantId);
    // Logger.log("Service: %s", service.getStorage());
    // Logger.log("Stored tenantId: %s", service.getStorage().getValue('tenantId'));
    refreshSidebar();
    return HtmlService.createHtmlOutput('Success! You can close this window.');
  } else {
    return HtmlService.createHtmlOutput('Denied.');
  }
}

/**
 * Get the Balance Sheet
 */
function getBalanceSheet(input) {
  Logger.log("Get Balance userKey: %s", getUserKey());
  var service = getService(getUserKey());
  // Logger.log("Service has access: %s", service.hasAccess());
  // Logger.log('Token: ' + JSON.stringify(service.getToken()));
  // Logger.log('Last Error: ' + service.getLastError());
  if (service.hasAccess()) {
    var queryString = buildStringFromForm(input);
    Logger.log("Query String: %s", queryString);
    var serviceStorage = service.getStorage();
    Logger.log("Storage: %s", serviceStorage);
    var tenantId = serviceStorage.getValue('tenantId');
    Logger.log("Tenant ID: %s", tenantId);
    const url = 'https://api.xero.com/api.xro/2.0/Reports/BalanceSheet' + queryString;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Xero-tenant-id': tenantId
      }
    })
    var result = JSON.parse(response.getContentText());
    var report = result.Reports[0];
    // Logger.log(result);
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.clearContents();
    var curRow = 1;
    for (const title of report.ReportTitles) {
      sheet.getRange(curRow,1,1,1).setValue(title);
      curRow++;
    }

    var cleanRows = report.Rows.map(row => {
      if (row.Cells) {
        row.Cells = row.Cells.map(cell => {
          Logger.log("Cell: %s", cell)
          Logger.log("Cell Value: %s", cell.Value)
          return cell.Value
      });
      }
      if (row.Rows) {
        for (const nestRow of row.Rows) {
          if(nestRow.Cells) {
            nestRow.Cells = nestRow.Cells.map(cell => {
              Logger.log("Cell: %s", cell)
              Logger.log("Cell Value: %s", cell.Value)
              return cell.Value
            });
          }
        }
      }
      return row;
    })
    Logger.log(cleanRows);

    for (const row of cleanRows) {
      Logger.log(row);
      switch(row.RowType) {
        case "Header":
          sheet.getRange(curRow, 1, 1, row.Cells.length).setValues([row.Cells]);
          curRow++;
          break;
        case "Section":
          if(row.Title != ""){
            sheet.getRange(curRow, 1).setValue(row.Title);
            curRow++;
          }
          for (const secRow of row.Rows) {
            switch(secRow.RowType) {
              case "Row":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
              case "SummaryRow":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
            }
          }
      }
    }

    return "Balance Sheet Loaded.";
    // Logger.log(connections[0].reports);
  }

}

/**
 * Get Profit and Loss Report
 */
function getProfitAndLoss(input) {
  var service = getService(Session.getActiveUser().getEmail());
  Logger.log("Service has access: %s", service.hasAccess());
  if (service.hasAccess()) {
    var queryString = buildStringFromForm(input);
    Logger.log("Query String: %s", queryString);
    var serviceStorage = service.getStorage();
    Logger.log("Storage: %s", serviceStorage);
    var tenantId = service.getStorage().getValue('tenantId');
    Logger.log("Tenant ID: %s", tenantId);
    const url = 'https://api.xero.com/api.xro/2.0/Reports/ProfitAndLoss' + queryString;
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Xero-tenant-id': tenantId
      }
    })
    var result = JSON.parse(response.getContentText());
    var report = result.Reports[0];
    // Logger.log(result);
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.clearContents();
    var curRow = 1;
    for (const title of report.ReportTitles) {
      sheet.getRange(curRow,1,1,1).setValue(title);
      curRow++;
    }

    var cleanRows = report.Rows.map(row => {
      if (row.Cells) {
        row.Cells = row.Cells.map(cell => {
          Logger.log("Cell: %s", cell)
          Logger.log("Cell Value: %s", cell.Value)
          return cell.Value
      });
      }
      if (row.Rows) {
        for (const nestRow of row.Rows) {
          if(nestRow.Cells) {
            nestRow.Cells = nestRow.Cells.map(cell => {
              Logger.log("Cell: %s", cell)
              Logger.log("Cell Value: %s", cell.Value)
              return cell.Value
            });
          }
        }
      }
      return row;
    })
    Logger.log(cleanRows);

    for (const row of cleanRows) {
      Logger.log(row);
      switch(row.RowType) {
        case "Header":
          sheet.getRange(curRow, 1, 1, row.Cells.length).setValues([row.Cells]);
          curRow++;
          break;
        case "Section":
          if(row.Title != ""){
            sheet.getRange(curRow, 1).setValue(row.Title);
            curRow++;
          }
          for (const secRow of row.Rows) {
            switch(secRow.RowType) {
              case "Row":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
              case "SummaryRow":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
            }
          }
      }
    }
  }
}

/**
 * Get Accounts Payable
 */
function getAccountsPayable() {
  var service = getService(Session.getActiveUser().getEmail());
  Logger.log("Service has access: %s", service.hasAccess());
  if (service.hasAccess()) {
    var serviceStorage = service.getStorage();
    Logger.log("Storage: %s", serviceStorage);
    var tenantId = service.getStorage().getValue('tenantId');
    Logger.log("Tenant ID: %s", tenantId);
    const url = 'https://api.xero.com/api.xro/2.0/Reports/AgedPayablesByContact?ContactID=2a320b2c-4190-4c58-afe7-69ebfe241988';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Xero-tenant-id': tenantId
      }
    })
    var result = JSON.parse(response.getContentText());
    var report = result.Reports[0];
    Logger.log(result);
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.clearContents();
    var curRow = 1;
    for (const title of report.ReportTitles) {
      sheet.getRange(curRow,1,1,1).setValue(title);
      curRow++;
    }

    var cleanRows = report.Rows.map(row => {
      if (row.Cells) {
        row.Cells = row.Cells.map(cell => {
          Logger.log("Cell: %s", cell)
          Logger.log("Cell Value: %s", cell.Value)
          return cell.Value
      });
      }
      if (row.Rows) {
        for (const nestRow of row.Rows) {
          if(nestRow.Cells) {
            nestRow.Cells = nestRow.Cells.map(cell => {
              Logger.log("Cell: %s", cell)
              Logger.log("Cell Value: %s", cell.Value)
              return cell.Value
            });
          }
        }
      }
      return row;
    })
    Logger.log(cleanRows);

    for (const row of cleanRows) {
      Logger.log(row);
      switch(row.RowType) {
        case "Header":
          sheet.getRange(curRow, 1, 1, row.Cells.length).setValues([row.Cells]);
          curRow++;
          break;
        case "Section":
          if(row.Title != ""){
            sheet.getRange(curRow, 1).setValue(row.Title);
            curRow++;
          }
          for (const secRow of row.Rows) {
            switch(secRow.RowType) {
              case "Row":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
              case "SummaryRow":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
            }
          }
      }
    }
  }
}

/**
 * Get Accounts Receivable
 */
function getAccountsReceivable() {
  var service = getService(Session.getActiveUser().getEmail());
  Logger.log("Service has access: %s", service.hasAccess());
  if (service.hasAccess()) {
    var serviceStorage = service.getStorage();
    Logger.log("Storage: %s", serviceStorage);
    var tenantId = service.getStorage().getValue('tenantId');
    Logger.log("Tenant ID: %s", tenantId);
    const url = 'https://api.xero.com/api.xro/2.0/Reports/AgedReceivablesByContact?ContactID=2a320b2c-4190-4c58-afe7-69ebfe241988';
    var response = UrlFetchApp.fetch(url, {
      headers: {
        'Authorization': 'Bearer ' + service.getAccessToken(),
        'Xero-tenant-id': tenantId
      }
    })
    var result = JSON.parse(response.getContentText());
    var report = result.Reports[0];
    Logger.log(result);
    const sheet = SpreadsheetApp.getActiveSheet();
    sheet.clearContents();
    var curRow = 1;
    for (const title of report.ReportTitles) {
      sheet.getRange(curRow,1,1,1).setValue(title);
      curRow++;
    }

    var cleanRows = report.Rows.map(row => {
      if (row.Cells) {
        row.Cells = row.Cells.map(cell => {
          Logger.log("Cell: %s", cell)
          Logger.log("Cell Value: %s", cell.Value)
          return cell.Value
      });
      }
      if (row.Rows) {
        for (const nestRow of row.Rows) {
          if(nestRow.Cells) {
            nestRow.Cells = nestRow.Cells.map(cell => {
              Logger.log("Cell: %s", cell)
              Logger.log("Cell Value: %s", cell.Value)
              return cell.Value
            });
          }
        }
      }
      return row;
    })
    Logger.log(cleanRows);

    for (const row of cleanRows) {
      Logger.log(row);
      switch(row.RowType) {
        case "Header":
          sheet.getRange(curRow, 1, 1, row.Cells.length).setValues([row.Cells]);
          curRow++;
          break;
        case "Section":
          if(row.Title != ""){
            sheet.getRange(curRow, 1).setValue(row.Title);
            curRow++;
          }
          for (const secRow of row.Rows) {
            switch(secRow.RowType) {
              case "Row":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
              case "SummaryRow":
                sheet.getRange(curRow, 1, 1, secRow.Cells.length).setValues([secRow.Cells]);
                curRow++;
                break;
            }
          }
      }
    }
  }
}

function buildStringFromForm(input) {
  const validInput = Object.keys(input).filter(key => input[key] != "" && key != "report");
  let queryString = "";
  if (validInput.length > 0) {    
    queryString += "?";
    let queryArray = validInput.map(key => {
      return `${key}=${input[key]}`
    })
    queryString += queryArray.join('&')
  }
  return queryString
}


/**
 * Logs the redict URI to register in the Dropbox application settings.
 */
function logRedirectUri() {
  Logger.log(OAuth2.getRedirectUri());
}