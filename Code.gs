/**
 * Copyright Google LLC
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @OnlyCurrentDoc  Limits the script to only accessing the current spreadsheet.
 */

const REPORT_SIDEBAR_TITLE = "Reports";
const CONFIGURATION_SIDEBAR_TITLE = "Configuration";

/**
 * Adds a custom menu with items to show the sidebar and dialog.
 *
 * @param {Object} e The event parameter for a simple onOpen trigger.
 */
function onOpen(e) {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem("Configuration", "showConfigurationSidebar")
    .addItem("Reports", "showReportsSidebar")
    .addToUi();
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure menu creation and
 * any other initializion work is done immediately.
 *
 * @param {Object} e The event parameter for a simple onInstall trigger.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Opens a sidebar. The sidebar structure is described in the Sidebar.html
 * project file.
 */
function showConfigurationSidebar() {
	var ui = HtmlService.createTemplateFromFile("Configuration Sidebar")
		.evaluate()
		.setTitle(CONFIGURATION_SIDEBAR_TITLE);
	SpreadsheetApp.getUi().showSidebar(ui);
}

function showReportsSidebar() {
  var ui = HtmlService.createTemplateFromFile("Reports Sidebar")
    .evaluate()
    .setTitle(REPORT_SIDEBAR_TITLE);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Opens a dialog. The dialog structure is described in the Dialog.html
 * project file.
 */
function showDialog() {
  var ui = HtmlService.createTemplateFromFile("Dialog")
    .evaluate()
    .setWidth(400)
    .setHeight(190);
  SpreadsheetApp.getUi().showModalDialog(ui, DIALOG_TITLE);
}

function loadConfiguration() {
  try {
    const props = PropertiesService.getScriptProperties();
    const config = props.getProperties();
    return config;
  } catch (e) {
    console.error(`Failed to load properties: ${e.toString()}`);
    return {};
  }
}

function saveConfiguration(formObject) {
  try {
    let props = PropertiesService.getScriptProperties();
    props.setProperty('APPFOLIO_CLIENT_ID', formObject.clientID);
    props.setProperty('APPFOLIO_CLIENT_SECRET', formObject.clientSecret);
    props.setProperty('APPFOLIO_API_VERSION', formObject.apiVersion);
    props.setProperty('COMPANY_NAME', formObject.companyName);
  } catch (e) {
    console.error(`Failed to save properties: ${e.toString()}`);
    return {};
  }
}

function runSelectedReport(reportName, apiPayload) {
    //const reportEndpoint = formObject.report;
    Logger.log("Received Report Name: " + reportName);
    Logger.log("Received API Payload: " + JSON.stringify(apiPayload));

    try {
      callApi(reportName, apiPayload);
    } catch (e) {
      console.log(`Error occured running API: ${e.toString}`);
    }
}

function callApi(reportName, apiPayload) {
  let props = PropertiesService.getScriptProperties();
  let url = `https://${props.getProperty('COMPANY_NAME')}.appfolio.com/api/${props.getProperty('APPFOLIO_API_VERSION')}/reports/${reportName}.json`;
  
  //Determine credentials from sheet
  let options = {
    method: 'post', // Use GET for fetching data
    contentType: 'application/json',
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(
          props.getProperty('APPFOLIO_CLIENT_ID') + ":" + props.getProperty('APPFOLIO_CLIENT_SECRET')
      )
    },
    payload: JSON.stringify(apiPayload),
    muteHttpExceptions: true // Prevents script from stopping on API errors
  };


  let response = UrlFetchApp.fetch(url, options);

  let responseCode = response.getResponseCode();
  let responseText = response.getContentText();

  Logger.log("External API Response Code: " + responseCode);
  Logger.log("External API Raw Response Content (first 500 chars): " + responseText.substring(0, 500)); // Log only part if it's huge
  Logger.log("External API Raw Response Content (FULL): " + responseText); // Log full for complete inspection

  if (responseCode >= 400) {
    throw new Error(`API returned an error: Status ${responseCode}, Response: ${responseText}`);
  }

  Logger.log(`External API Parsed Response: ${JSON.stringify(JSON.parse(response.getContentText()))}`)

  try {
    routeAPICallback(reportName, JSON.parse(response.getContentText()));
  } catch (e) {
    console.log(`Error requesting API response for URL=${url}: Error ${e.toString()}`);
  }
}

function showReportSpecificSidebar(reportName) {
  let sidebarTitle = "";
  let htmlFile = "";
  let templateData = {}; // Object to hold data for the template

  // Use a switch statement to handle different reports
  switch (reportName) {
    case "account_totals":
      htmlFile = "Account Totals Sidebar";
      sidebarTitle = "Account Totals Parameters";
      templateData.glAccounts = getGeneralLedgerData();
      break;

    default:
      // Optional: show a generic message if no specific sidebar exists
      SpreadsheetApp.getUi().alert(`No specific parameter sidebar has been created for the "${reportName}" report.`);
      return;
  }

  const template = HtmlService.createTemplateFromFile(htmlFile);
  // Pass the report name to the template for the hidden input field
  template.data = templateData;
  template.reportName = reportName; 
  
  const ui = template.evaluate().setTitle(sidebarTitle);
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Fetches GL accounts from the AppFolio API to populate a dropdown.
 * @returns {Array} A list of objects, each with an id and name.
 */
function getGeneralLedgerData() {
  let props = PropertiesService.getScriptProperties();
  let gl_url = `https://${props.getProperty('COMPANY_NAME')}.appfolio.com/api/${props.getProperty('APPFOLIO_API_VERSION')}/reports/general_ledger.json`;

  let options = {
    method: 'post', // Use GET for fetching data
    contentType: 'application/json',
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(
          props.getProperty('APPFOLIO_CLIENT_ID') + ":" + props.getProperty('APPFOLIO_CLIENT_SECRET')
      )
    },
    muteHttpExceptions: true // Prevents script from stopping on API errors
  };

  try {
    let response = UrlFetchApp.fetch(gl_url, options);
    let gl_data = JSON.parse(response.getContentText());

    let gl_account_ids = [];

    if (gl_data && gl_data.results && Array.isArray(gl_data.results)) {
      gl_data.results.forEach(item => {
        if (item.account_id !== undefined && item.account_id !== null) {
          gl_account_ids.push(item.account_id);
        }
      });

      const uniqueAccountIds = [...new Set(gl_account_ids)];

      uniqueAccountIds.sort((a, b) => Number(a) - Number(b));

      return uniqueAccountIds;
    } else {
      console.error("gl_data.account_id is not an array or does not exist:", gl_data);
      return [];
    }
  } catch (e) {
    console.error(`Error fetching GL data: ${e.toString()}`);
    return [];
  }
}
