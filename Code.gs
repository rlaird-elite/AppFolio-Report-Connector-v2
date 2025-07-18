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

function runSelectedReport(formObject) {
    const reportEndpoint = formObject.report;
    console.log(reportEndpoint);

    try {
      callApi(reportEndpoint);
    } catch (e) {
      console.log(`Error occured running API: ${e.toString}`);
    }
}

function callApi(endpoint) {
  let props = PropertiesService.getScriptProperties();
  let url = `https://${props.getProperty('COMPANY_NAME')}.appfolio.com/api/${props.getProperty('APPFOLIO_API_VERSION')}/reports/${endpoint}.json`;
  
  //Determine credentials from sheet
  let options = {};
  options.method = 'post';
  options.headers = {
      "Authorization": "Basic " + Utilities.base64Encode(
          props.getProperty('APPFOLIO_CLIENT_ID')
          + ":" +
          props.getProperty('APPFOLIO_CLIENT_SECRET')
      ),
      "Content-Type": "application/json"
  };


  let response = UrlFetchApp.fetch(url, options);

  try {
    routeAPICallback(endpoint, JSON.parse(response.getContentText()));
  } catch (e) {
    console.log(`Error requesting API response for URL=${url}: Error ${e.toString()}`);
  }
}
