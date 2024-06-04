function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Hunter.io')
    .addItem('Set API Key', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Hunter.io API Key');
  SpreadsheetApp.getUi().showSidebar(html);
}

function setApiKey(apiKey) {
  PropertiesService.getUserProperties().setProperty('HUNTER_API_KEY', apiKey);
}

function findEmailByCompany(firstName, lastName, company) {
  if (!firstName) {
    return "Error: First name is required.";
  }
  if (!lastName) {
    return "Error: Last name is required.";
  }
  if (!company) {
    return "Error: Company name is required.";
  }

  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  if (!apiKey) {
    return "Error: API key not set. Please set it using the custom menu.";
  }

  var url = `https://api.hunter.io/v2/email-finder?first_name=${firstName}&last_name=${lastName}&company=${company}&api_key=${apiKey}`;

  Logger.log(`Request URL: ${url}`);

  try {
    var response = UrlFetchApp.fetch(url);
    var result = JSON.parse(response.getContentText());

    Logger.log(`Response: ${JSON.stringify(result)}`);

    if (result.data && result.data.email) {
      return result.data.email;
    } else if (result.errors && result.errors.length > 0) {
      var errorDetails = result.errors[0].details;
      if (result.errors[0].id === "email_not_found") {
        return "Error: Email not found for the given details.";
      } else if (result.errors[0].id === "quota_exceeded") {
        return "Error: API quota exceeded.";
      } else if (result.errors[0].id === "wrong_params") {
        return "Error: Invalid parameters provided.";
      } else if (result.errors[0].id === "invalid_company") {
        return "Error: Invalid company name.";
      } else if (result.errors[0].id === "claimed_email") {
        return "Error: The person behind this email address has requested their information not be disclosed.";
      } else {
        return `Error: ${errorDetails}`;
      }
    } else {
      throw new Error('Unknown error occurred');
    }
  } catch (e) {
    Logger.log('Error: ' + e.message);
    return `Error: ${e.message}`;
  }
}

function FindEmail(firstName, lastName, company) {
  try {
    return findEmailByCompany(firstName, lastName, company);
  } catch (e) {
    return `Error: ${e.message}`;
  }
}

function logApiKey() {
  var apiKey = PropertiesService.getUserProperties().getProperty('HUNTER_API_KEY');
  Logger.log(apiKey);
}
