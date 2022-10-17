/*
  Script constants
 */
const DISPLAY_AIRCALL_CTI = "displayAircallCTI";
const CLICK_TO_DIAL = "clickToDial";
const CREATE_DIALER_CAMPAIGN = "createDialerCampaign";

// List of GS User Emails and their associated Aircall User IDs
const AIRCALL_USER_IDS = {
  "user1@gmail.com": "111111",
  "user2@gmail.com": "222222",
  "user3@gmail.com": "333333"
};

/*
  Function to call when the spreadsheet opens, which sets up the environment.
 */
function onOpen(e) {
  // Create the special menu on the GS toolbar
  SpreadsheetApp.getUi()
    .createMenu('Aircall')
    .addItem('Display Aircall CTI', DISPLAY_AIRCALL_CTI)
    .addItem('Click-to-Dial', CLICK_TO_DIAL)
    .addItem('Create Dialer Campaign', CREATE_DIALER_CAMPAIGN)
    .addToUi();
}

/*
  Function that simulates Aircall's click-to-dial.
 */
async function clickToDial() {
  // Retrieve the value from the selected cell
  const ss = SpreadsheetApp.getActiveSheet();
  const activeCellValue = ss.getActiveCell().getValue().toString();
  const ui = SpreadsheetApp.getUi();

  // If the value does not contain a '+' at the start, add it
  const phoneTo = (activeCellValue.includes("+") == true) ? activeCellValue : '+' + activeCellValue;

  // Quick check to see if the selected value is even a phone number...
  if (quickNumberCheck(phoneTo)) {
    // Create the payload for the API request
    const data = { to: phoneTo };

    // Make Aircall API request
    await aircallAPIRequest(CLICK_TO_DIAL, 'POST', data, 204)

  // ...else, the cell value is not even a phone number
  } else {
    ui.alert(CLICK_TO_DIAL + '(): Number format is not valid...');
    Logger.log(CLICK_TO_DIAL + '(): Number format is not valid...');
  }
};

/*
  Function that displays the Aircall CTI in the GS sidebar.
 */
function displayAircallCTI() {
  const aircallSoftphone = HtmlService.createHtmlOutputFromFile('CTI');
  aircallSoftphone.setTitle('Aircall CTI');
  SpreadsheetApp.getUi().showSidebar(aircallSoftphone);
};

/*
  Function that creates an Aircall Dialer Campaign.
 */
async function createDialerCampaign() {
  // Retrieve the values from all the selected cells
  const ss = SpreadsheetApp.getActiveSheet();
  let selectedCellValues = ss.getActiveRange().getValues();
  const ui = SpreadsheetApp.getUi();

  // List to keep track of the potentially well-formatted phone numbers
  let cleanPhoneNumbers = [];

  // For each row...
  for (var i = 0; i < selectedCellValues.length; i++) {
    // For each column (from the current row)...
    for (var j = 0; j < selectedCellValues[i].length; j++) {
      // If the current cell has a value and passes a quick validation...
      if (selectedCellValues[i][j] && quickNumberCheck(selectedCellValues[i][j])) {
        // Add value to current list
        cleanPhoneNumbers.push(selectedCellValues[i][j]);
      }
    }
  }

  // If there is at least one potentially well-formatted phone number...
  if (cleanPhoneNumbers.length) {

    // Create the payload for the API request
    const data = {
      phone_numbers: cleanPhoneNumbers
    };

    // Make Aircall API request
    await aircallAPIRequest(CREATE_DIALER_CAMPAIGN, 'POST', data, 204)

  // ...else, we do not have any potentially valid phone numbers
  } else {
    ui.alert(CREATE_DIALER_CAMPAIGN + '(): Could not find any valid numbers...');
    Logger.log(CREATE_DIALER_CAMPAIGN + '(): Could not find any valid numbers...');
  }
};

/*
  Function that makes an Aircall API request.
  @param  {String} request_type     Type of Aircall API request
          {String} verb             API verb to use
          {JSON} payload            Payload of API request (if required)
          {Integer} success_code    Code return if API request is successful
 */
async function aircallAPIRequest(request_type, verb, payload, success_code) {
  const ui = SpreadsheetApp.getUi();

  // Try the code below...
  try {
    // Retrieve GS Script Properties related to this spreadsheet
    const scriptProperties = PropertiesService.getScriptProperties();
    const baseUrl = scriptProperties.getProperty('BASE_URL');
    const apiID = scriptProperties.getProperty('API_ID');
    const apiToken = scriptProperties.getProperty('API_TOKEN');
    
    // Retrieve the current GS User's Aircall User ID
    const userEmail = Session.getActiveUser();
    const aircallUserID = AIRCALL_USER_IDS[userEmail];

    let aircallAPIURI = "";

    // Define the url for the Aircall Click-to-Dial API request
    if (request_type == CLICK_TO_DIAL) {
      aircallAPIURI = baseUrl + 'users/' + aircallUserID + '/dial';
    } else if (request_type == CREATE_DIALER_CAMPAIGN) {
      aircallAPIURI = baseUrl + 'users/' + aircallUserID + '/dialer_campaign';
    }

    // If we have the Aircall User's ID...
    if (aircallUserID) {
      // Execute the API request
      let req = await UrlFetchApp.fetch(aircallAPIURI, {
        method: verb,
        headers: {
          'Authorization': 'Basic ' + Utilities.base64Encode(apiID + ':' + apiToken),
          'Content-Type': 'application/json',
        },
        'muteHttpExceptions': true,
        payload: JSON.stringify(payload)
      });

      // If the API request was not successful...
      if (req.getResponseCode() != success_code) {
        if (request_type == CLICK_TO_DIAL) {
          ui.alert(CLICK_TO_DIAL + '(): Invalid number to call!');
        } else if (request_type == CREATE_DIALER_CAMPAIGN) {
          ui.alert(CREATE_DIALER_CAMPAIGN + '(): ' + req.getContentText());
        }
        
        Logger.log(request_type + '(): ' + req.getContentText());
      }

    // ...else, we do not have the Aircall User's ID
    } else {
      ui.alert(request_type + '(): Could not retrieve Aircall User ID...');
      Logger.log(request_type + '(): Could not retrieve Aircall User ID...');
    }
  
  // ...else, an error occurred, so print the error
  } catch (err) {
    Logger.log(request_type + '(): Failed to run this function with error: %s', err.message);
  }
}

/*
  Function that does a quick validation on a cell value to determine if it has the
  potential to be a well-formatted phone number.
 */
function quickNumberCheck(str) {
  // Match on a cell value if it...
  //    - [OPTIONAL]: Starts with a '+'
  //    - Contains only numbers
  var rx = new RegExp("^\\+*\\d+$","gm"); 
  let checkStr = rx.exec(str);

  // If the cell value matches, return true
  if (checkStr) {
    return true;
  } 
  
  // Else, return false
  return false;
};
