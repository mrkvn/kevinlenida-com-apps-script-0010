function logSharedSheets() {
  const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  const lastCheck = PropertiesService.getScriptProperties().getProperty('lastCheck');
  const currentTime = Math.floor(Date.now() / 1000);
  const lastCheckTime = lastCheck ? parseInt(lastCheck) : currentTime;

  const query = 'to:me after:' + lastCheckTime + ' subject:"shared with you"';
  const threads = GmailApp.search(query);

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      const body = message.getBody();
      const urls = findGoogleSheetsLinks(body);
      if (urls.length === 0) {
        return
      }
      const shared_by = message.getFrom()
      const receivedDate = message.getDate();
      urls.forEach(url => {
          const fileId = extractFileId(url);
          const file = DriveApp.getFileById(fileId);
          const fileName = file.getName();
          logSheet.appendRow([shared_by, fileName, url, receivedDate, new Date()]);
        });
    });
  });

  PropertiesService.getScriptProperties().setProperty('lastCheck', currentTime.toString());

}

function findGoogleSheetsLinks(body) {
  const regex = /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/g;
  let match;
  const urlsSet = new Set();
  while ((match = regex.exec(body)) !== null) {
    urlsSet.add(match[0]);
  }
  return Array.from(urlsSet);
}

function extractFileId(url) {
  const regex = /https:\/\/docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/;
  const match = url.match(regex);
  return match ? match[1] : null;
}

function setupTrigger() {
  ScriptApp.newTrigger('logSharedSheets')
    .timeBased()
    .everyDays(1)
    .create();
}

// Run this function once to set up the trigger
function initialize() {
  setupTrigger();
}
