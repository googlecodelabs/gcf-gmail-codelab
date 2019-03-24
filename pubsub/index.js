const clientSecrets = require('./client_secret.json');
const {Datastore} = require('@google-cloud/datastore');
const {google} = require('googleapis');
const gmail = google.gmail('v1');
const {Storage} = require('@google-cloud/storage');
const googleSheets = google.sheets('v4');

const datastoreClient = new Datastore();
const storageClient = new Storage();

const TAG = process.env.TAG;
const BUCKET = process.env.CLOUD_SaTORAGE_BUCKET;
const SHEET = process.env.GOOGLE_SHEET_ID;

const getMostRecentMessageWithTag = async (email, historyId) => {
  // Look up the most recent message using the history ID in the push
  // notification. The API call returns a message ID.
  var listMessagesRes = await gmail.users.history.list({
    userId: email,
    maxResults: 1,
    startHistoryId: historyId
  });
  var messageId = listMessagesRes.data.history[0].messages[0].id;

  // Get the message using the message ID.
  var message = await gmail.users.messages.get({
    userId: email,
    id: messageId
  });

  // Proceed only when the message has the keyword [SUBMISSION] in the subject.
  var headers = message.data.payload.headers;
  for (var x in headers) {
    if (headers[x].name === 'Subject' && headers[x].value.indexOf(TAG) > -1) {
      return message;
    }
  }
};

// Extract message ID, sender, attachment filename and attachment ID
// from the message.
const extractInfoFromMessage = (message) => {
  var messageId = message.data.id;
  var from;
  var filename;
  var attachmentId;

  var headers = message.data.payload.headers;
  for (var i in headers) {
    if (headers[i].name === 'From') {
      from = headers[i].value;
    }
  }

  var payloadParts = message.data.payload.parts;
  for (var j in payloadParts) {
    if (payloadParts[j].body.attachmentId) {
      filename = payloadParts[j].filename;
      attachmentId = payloadParts[j].body.attachmentId;
    }
  }

  return {
    messageId: messageId,
    from: from,
    attachmentFilename: filename,
    attachmentId: attachmentId
  };
};

// Get attachment of a message..
const extractAttachmentFromMessage = async (email, messageId, attachmentId) => {
  await gmail.users.messages.attachments.get({
    id: attachmentId,
    messageId: messageId,
    userId: email
  });
};

// Upload attachment of a message to Cloud Storage.
const uploadAttachment = async (data, filename) => {
  let file = storageClient.bucket(BUCKET).file(filename);
  let writeableStream = file.createWriteStream({resumable: false});
  writeableStream.write(data);
  writeableStream.end();
};

// Write sender, attachment filename, and download link to a Google Sheet.
const updateReferenceSheet = async (from, filename) => {
  let link = `https://storage.cloud.google.com/${BUCKET}/${filename}`;
  await googleSheets.spreadsheets.values.append({
    spreadsheetId: SHEET,
    range: 'Sheet1!A1:D1',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      range: 'Sheet1!A1:D1',
      majorDimension: 'ROWS',
      values: [
        [from, filename, link]
      ]
    }
  });
};

exports.watchGmailMessages = async (event) => {
  // Decode the incoming Gmail push notification.
  const data = Buffer.from(event.data, 'base64').toString();
  const newMessageNotification = JSON.parse(data);
  var email = newMessageNotification.emailAddress;
  var historyId = newMessageNotification.historyId;

  // Connect to Gmail API and Google Sheets API using the access tokens
  // from the authorization process.
  var credentialKey = datastoreClient.key(['oauth2token', `${email}`]);
  var [credentials] = await datastoreClient.get(credentialKey);
  var token = credentials.token;

  var OAuth2Client = new google.auth.OAuth2(clientSecrets.GOOGLE_CLIENT_ID,
    clientSecrets.GOOGLE_CLIENT_SECRET, clientSecrets.GOOGLE_CALLBACK_URL);
  OAuth2Client.setCredentials(token);
  google.options({auth: OAuth2Client});

  // Process the incoming message.
  var message = await getMostRecentMessageWithTag(email, historyId);
  var messageInfo = extractInfoFromMessage(message);
  if (messageInfo) {
    var attachment = await extractAttachmentFromMessage(email, messageInfo.messageId, messageInfo.attachmentId);
    await uploadAttachment(attachment.data.data, `${messageInfo.messageId}_${messageInfo.attachmentFilename}`);
    await updateReferenceSheet(messageInfo.from, messageInfo.attachmentFilename);
  }
};
