// Credit: Elder Scoresby, Philippines Manila Mission 2023-2025 

const contactMasterURL = "https://script.google.com/macros/s/AKfycby9sGeAGqsdg6TuEwF9VKJPSOR5aLKb-KB0usuHYu2blLPwFJ39qG5t0Us0OgdT4AYz/exec";
const labelName = "IMOS Roster";
const calendarURL = "c_f61d2b5c52f1f0a6bc03a805cff44db56b963d9a9bca6bca1420d56b523a589e@group.calendar.google.com";

// ---------- Utility Functions ----------
/**
 * Recursively removes unwanted keys from objects or arrays.
 * @param {Object|Array} obj - The object or array to clean.
 * @param {Array} keys - The keys to remove (default: ["metadata"]).
 */

/**
 * Logs messages and optionally sends email alerts for errors.
 * @param {string} message - The message to log
 * @param {string} level - "INFO" or "ERROR" (default: "INFO")
 * @param {boolean} notifyUser - Whether to send an email alert (default: false)
 */

function removeKeys(obj, keys = ["metadata"]) {
  if (!obj) return;

  if (Array.isArray(obj)) {
    obj.forEach(item => removeKeys(item, keys));
  } else if (typeof obj === 'object') {
    keys.forEach(key => {
      if (obj.hasOwnProperty(key)) delete obj[key];
    });
    // Recursively clean nested objects
    Object.values(obj).forEach(val => {
      if (typeof val === 'object') removeKeys(val, keys);
    });
  }
}

// ---------- Constants / Config Options ----------


// Retry / API Config
const MAX_RETRIES = 6;              // number of retry attempts for API calls
const BASE_DELAY_MS = 1000;         // base delay for exponential backoff

// Script Behavior
const DRY_RUN = false;              // if true, simulates updates without changing contacts
const SEND_ERROR_EMAILS = true;     // if true, sends error alerts to current user



/** 1/22/2026 Simplifed 
 * Logs errors and optionally notifies via email
 * @param {string} message - Error message to log
 * @param {boolean} notifyUser - Whether to send an email alert
 */
function handleError(message, notifyUser = true) {
  Logger.log("ERROR: " + message);
  console.error(message);

  if (notifyUser) {
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(), // sends to current user
        subject: "IMOS Script Error!!",
        body: message
      });
    } catch (e) {
      Logger.log("Failed to send error email: " + e + " Elder Jenkins get ya butt in the office and fix this.");
    }
  }
}

function logMessage(message, level = "INFO", notifyUser = false) {
  const timestamp = new Date().toISOString();
  const fullMessage = `[${timestamp}] [${level}] ${message}`;

  if (level === "ERROR") {
    console.error(fullMessage);
  } else {
    Logger.log(fullMessage);
  }

  if (notifyUser && level === "ERROR") {
    try {
      MailApp.sendEmail({
        to: Session.getEffectiveUser().getEmail(),
        subject: "IMOS Script Error",
        body: fullMessage
      });
    } catch (e) {
      Logger.log(`[${timestamp}] Failed to send error email: ${e}`);
    }
  }
}

function backupContacts(contacts) { const sheet = SpreadsheetApp.openById("111jdbU_NmFZEYaUcB6bhLvNzj-BmmadQjOQr8gNPOqM").getSheetByName("Backup"); sheet.clear(); contacts.forEach((c, i) => sheet.getRange(i + 1, 1).setValue(JSON.stringify(c))); }

/**
 *-------------- MAIN FUNCTIONS --------------
 */
// This function runs once everyday (when the trigger is set correctly)
function dailyUpdate() {
  try {
    Logger.log("--- DAILY UPDATE START ---");

    let contactList = getData();
    if (contactList && contactList.length > 0) {
      if (deleteAllContacts()) {
        uploadContacts(contactList);
      } else {
        handleError("Failed to delete existing contacts.");
      }
    } else {
      handleError("No contacts retrieved from master list.");
    }

    // Optional: joinCalendar();
    Logger.log("--- DAILY UPDATE COMPLETE ---");
  } catch (e) {
    handleError("Daily update failed: " + e);
  }
}


function testCompareContacts() {
  let currUserContacts = getContacts();
  let currNameList = currUserContacts.map(x=>x.names[0].displayName).filter(n => n);
  let newContacts = getData();
  let newNameList = newContacts.map(x=>x.names[0].displayName).filter(n => n);
  currNameList.sort();
  newNameList.sort();
  console.log(currNameList.join() == newNameList.join());
  console.log(JSON.stringify(currUserContacts[47], null, 2));
  for (let i=0;i<newNameList.length;i++){
    if (newNameList[i] != currNameList[i]) {
      console.log(newNameList[i] + " | " + currNameList[i] + " | " + i);
    }
  }
  console.log()
  //console.log(JSON.stringify(currUserContacts.map(x=>x.names[0].displayName)))
}

function testCalendar() {
  const calendars = CalendarApp.getAllCalendars();
  let prettyCalendars = calendars.map(x=>(x.getName() + " | " + x.getId())).join(", ");
  console.log(prettyCalendars);
  if (calendars.some(x=>(x.getId() === calendarURL))) {
    console.log("This User has the Calendar!")
  }
}

function joinCalendar() {
  const calendars = CalendarApp.getAllCalendars();
  if (!calendars.some(x=>(x.getId() === calendarURL))) {
    var calendar = CalendarApp.subscribeToCalendar(calendarURL);
    Logger.log("User was not on the calendar, but is now subscribed to " + calendar.getName());
  }
}

// Gets the Contacts from other Main Mission account
function getData() {

  // Gets the current user's email address
  var ownEmail = Session.getEffectiveUser().getEmail();
  Logger.log("My email is: " + ownEmail);
  
  // Payload is the information to be sent
  var payload = {
    email: ownEmail
  };

  // Sets up the POST request
  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  // Gets data from POST request
  Logger.log("Sending POST Request");
  var responseData = retryAttempts(function() {
    var response = UrlFetchApp.fetch(contactMasterURL, options);
    Logger.log("Response Recieved: " + response.getContentText());
    return JSON.parse(response.getContentText());
  });
  Logger.log("Recieved response data from POST request with status: " + responseData.status);

  if (responseData.status == "Success") {
    
    // Returns the contact list
    return responseData.contacts; 

  } else if (responseData.status == "Fail") {

    // Throws the error from the POST Request
    console.error(responseData);
    if (responseData.message == "Did not find own Email") {
      return [];
    } else {
      throw responseData.error;
    }

  } else {

    // Unknown response... throws error
    console.error(responseData);
    throw new "Unknown response";
  }

}

// Deletes the full IMOS Roster from the current user's account
function deleteAllContacts() {

  // Gets the IMOS Roster of the current user
  var allContacts = retryAttempts(function() {
    return getContacts();
  });

  try {

    var contactIds = [];
    // Collects Ids of each Contact
    for (var i = 0; i < allContacts.length; i++) {
      contactIds.push(allContacts[i].resourceName);
    }
  
    // Deletes Contacts
    try {
      Logger.log("Deleting Contacts with the Following Ids: " + contactIds);
      retryAttempts(function() {
        People.People.batchDeleteContacts({resourceNames: contactIds});
      });
      return true;
    } catch (e) {
      console.error("Failed to Delete Contacts");
      return false;
    }

  } catch (e) {
    Logger.log("No Contacts Deleted");
    return true;
  }
}

function getContacts() {
  // Credit: https://stackoverflow.com/questions/77531546/api-people-from-google-in-apps-script-search-by-label

  // 1. Retrieve the label list using People.ContactGroups.list.
  const { contactGroups } = People.ContactGroups.list({ groupFields: "memberCount,name", pageSize: 1000 });

  // 2. Retrieve the resource name of the label.
  const obj1 = contactGroups.find(({ name }) => name == labelName);
  if (!obj1) {
    Logger.log("Label " + labelName + " was not found. Creating Label now");
    People.ContactGroups.create({contactGroup: {name: "IMOS Roster"}});
    return;
  }
  const { resourceName, memberCount } = obj1;

  // 3. Retrieve the member resource names using the retrieved resource name with People.ContactGroups.get.
  const { memberResourceNames } = People.ContactGroups.get(resourceName, { maxMembers: memberCount });

  // 4. Retrieve all contacts.
  const contacts = People.People.Connections.list('people/me', { personFields: 'names', pageSize: 1000 });

  // 5. Filter the retrieved contacts using the member resource names.
  try {
    const res = contacts.connections.filter(({ resourceName }) => memberResourceNames.includes(resourceName));
    return res;
  } catch (e) {
    Logger.log("No IMOS Roster")
  }

  
}
// Sends each of the contacts to current user's account
function uploadContacts(contactList) {
  
  // Runs for each contact given
  for (let i = 0; i < contactList.length; i++) {
    
    // Deletes information that will throw error if given
    try{
      for (j = 0; j < contactList[i].names.length; j++) {
        delete contactList[i].names[j].metadata;
      }
    } catch (e) {}
    try {
      for (j = 0; j < contactList[i].emailAddresses.length; j++) {
        delete contactList[i].emailAddresses[j].metadata;
      }
    } catch (e) {}
    try {
      for (j = 0; j < contactList[i].phoneNumbers.length; j++) {
        delete contactList[i].phoneNumbers[j].metadata;
      }
    } catch (e) {}
    try {
      for (j = 0; j < contactList[i].externalIds.length; j++) {
        delete contactList[i].externalIds[j].metadata;
      }
    } catch (e) {}
    try {
      for (j = 0; j < contactList[i].memberships.length; j++) {
        delete contactList[i].memberships[j].metadata;
      }
    } catch (e) {}
    try {
      for (j = 0; j < contactList[i].addresses.length; j++) {
        delete contactList[i].addresses[j].metadata;
      }
    } catch (e) {}
    try {
      delete contactList[i].phoneNumbers.metadata;
    } catch (e) {
      try {
        for (j = 0; j < contactList[i].phoneNumbers.length; j++) {
          delete contactList[i].phoneNumbers[j].metadata;
        }
      } catch (e) {}
    }
    try {
      for (j = 0; j < contactList[i].biographies.length; j++) {
        delete contactList[i].biographies[j].metadata;
      }
    } catch (e) {}

    // Creates a new contact with only the usable data
    let contactTemp = {
      names: contactList[i].names[0],
      emailAddresses: contactList[i].emailAddresses,
      biographies: contactList[i].biographies,
      phoneNumbers: contactList[i].phoneNumbers,
      externalIds: contactList[i].externalIds,
      addresses: contactList[i].addresses,
      memberships: [{contactGroupMembership: {contactGroupResourceName: getContactLabelId("IMOS Roster")}}]
     };

    // Create the contact using retryAttempts
    let createdContact = retryAttempts(() => People.People.createContact(contactTemp));
    Logger.log("Created Contact: " + (createdContact.names ? createdContact.names[0].displayName : "Unnamed"));

    
  }
}


// Conducts retry attempts on API calls
function retryAttempts(APICall, description = "API Call") {
  const MAX_ATTEMPTS = 6;
  const BASE_DELAY_MS = 1000;

  for (let attempt = 0; attempt < MAX_ATTEMPTS; attempt++) {
    try {
      return APICall();
    } catch (e) {
      const msg = `${description} failed on attempt ${attempt + 1}: ${e}`;
      Logger.log(msg);
      if (attempt >= 2) Utilities.sleep(BASE_DELAY_MS * Math.pow(2, attempt)); // exponential backoff
      if (attempt === MAX_ATTEMPTS - 1) handleError(msg); // notify user on final failure
    }
  }
  throw new Error(`${description} failed after ${MAX_ATTEMPTS} attempts`);
}




// Returns the Id of the Label
function getContactLabelId(labelName) {
  // Get the list of contact groups (labels)
  var response = People.ContactGroups.list({
    pageSize: 1000
  });

  var contactGroups = response.contactGroups;

  // Loop through the contact groups to find the ID of the specified label
  for (var i = 0; i < contactGroups.length; i++) {
    var contactGroup = contactGroups[i];
    if (contactGroup.formattedName == labelName) {
      var labelId = contactGroup.resourceName; // This is the ID of the contact label
      return labelId;
    }
  }

  // If label is not found, log an error
  console.error('Label not found: ' + labelName);
  return null;
}