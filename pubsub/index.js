const Auth = require('@google-cloud/express-oauth2-handlers');
const {google} = require('googleapis');
const gmail = google.gmail('v1');
const googleSheets = google.sheets('v4');
const vision = require('@google-cloud/vision');

const visionClient = new vision.ImageAnnotatorClient();

const TAG = process.env.TAG;
const SHEET = process.env.GOOGLE_SHEET_ID;

const requiredScopes = [
  'profile',
  'email',
  'https://www.googleapis.com/auth/gmail.modify',
  'https://www.googleapis.com/auth/spreadsheets'
];

const auth = Auth('datastore', requiredScopes, 'email', true);

const getMostRecentMessageWithTag = async (email, historyId) => {
  // Look up the most recent message using the history ID in the push
  // notification. The API call returns a message ID.
  const listMessagesRes = await gmail.users.history.list({
    userId: email,
    maxResults: 1,
    startHistoryId: historyId
  });
  const messageId = listMessagesRes.data.history[0].messages[0].id;

  // Get the message using the message ID.
  const message = await gmail.users.messages.get({
    userId: email,
    id: messageId
  });

  // Proceed only when the message has the keyword [SUBMISSION] in the subject.
  const headers = message.data.payload.headers;
  for (var x in headers) {
    if (headers[x].name === 'Subject' && headers[x].value.indexOf(TAG) > -1) {
      return message;
    }
  }
};

// Extract message ID, sender, attachment filename and attachment ID
// from the message.
const extractInfoFromMessage = (message) => {
  const messageId = message.data.id;
  var from;
  var filename;
  var attachmentId;

  const headers = message.data.payload.headers;
  for (var i in headers) {
    if (headers[i].name === 'From') {
      from = headers[i].value;
    }
  }

  const payloadParts = message.data.payload.parts;
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

// Get attachment of a message.
const extractAttachmentFromMessage = async (email, messageId, attachmentId) => {
  return gmail.users.messages.attachments.get({
    id: attachmentId,
    messageId: messageId,
    userId: email
  });
};

// Tag the attachment using Cloud Vision API
const analyzeAttachment = async (data, filename) => {
  var topLabels = ['', '', ''];
  if (filename.endsWith('.png') || filename.endsWith('.jpg')) {
    const [analysis] = await visionClient.labelDetection({
      image: {
        content: Buffer.from(data, 'base64')
      }
    });
    const labels = analysis.labelAnnotations;
    for (var i = 0; i <= 2; i++) {
      if (labels[i] && labels[i].description) {
        topLabels[i] = labels[i].description;
      }
    }
  }

  return topLabels;
};

// Write sender, attachment filename, and download link to a Google Sheet.
const updateReferenceSheet = async (from, filename, topLabels) => {
  await googleSheets.spreadsheets.values.append({
    spreadsheetId: SHEET,
    range: 'Sheet1!A1:F1',
    valueInputOption: 'USER_ENTERED',
    requestBody: {
      range: 'Sheet1!A1:F1',
      majorDimension: 'ROWS',
      values: [
        [from, filename].concat(topLabels)
      ]
    }
  });
};

exports.watchGmailMessages = async (event) => {
  // Decode the incoming Gmail push notification.
  const data = Buffer.from(event.data, 'base64').toString();
  const newMessageNotification = JSON.parse(data);
  const email = newMessageNotification.emailAddress;
  const historyId = newMessageNotification.historyId;
  console.log(historyId);

  try {
    await auth.auth.requireAuth(null, null, email);
  } catch (err) {
    console.log('An error has occurred in the auth process.');
    throw err;
  }
  const authClient = await auth.auth.authedUser.getClient();
  google.options({auth: authClient});

  // Process the incoming message.
  const message = await getMostRecentMessageWithTag(email, historyId);
  const messageInfo = extractInfoFromMessage(message);
  if (messageInfo) {
    const attachment = await extractAttachmentFromMessage(email, messageInfo.messageId, messageInfo.attachmentId);
    const topLabels = await analyzeAttachment(attachment.data.data, messageInfo.attachmentFilename);
    await updateReferenceSheet(messageInfo.from, messageInfo.attachmentFilename, topLabels);
  }
};
