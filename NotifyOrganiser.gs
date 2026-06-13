// David's script — runs under davidmurphy1088@gmail.com
// Sends organiser notification to Janet when someone registers.
// Deploy as Web App: Execute as: Me, Who has access: Anyone
// Paste the deployment URL into index.html as NOTIFY_ORGANISER_URL.

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const now   = new Date();
    const day   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd');
    const month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
    const time  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');

    MailApp.sendEmail({
      to:      'janetmckean1088@gmail.com',
      subject: 'Seminar Registration, ' + day + ' ' + month + ' ' + time + ' — ' + (data.name || ''),
      body:
        'New registration received ' + day + ' ' + month + ' at ' + time + '\n\n'
        + 'Name:          ' + (data.name          || '') + '\n'
        + 'Email:         ' + (data.email         || '') + '\n'
        + 'Phone:         ' + (data.phone         || '') + '\n'
        + 'Participation: ' + (data.participation || '') + '\n'
        + 'Accommodation: ' + (data.accommodation || '') + '\n'
        + 'Meals:         ' + (data.meals         || '') + '\n'
        + 'Gluten free:   ' + (data.gluten_free   || '') + '\n'
        + 'Special needs: ' + (data.special_needs || '') + '\n\n'
        + '▶️ Added to the Google Sheet.',
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Run this once from the editor to grant MailApp permission before deploying.
function testEmail() {
  MailApp.sendEmail({
    to:      'janetmckean1088@gmail.com',
    subject: 'Organiser notification test — Who, What, Where is God?',
    body:    'If Janet received this in her inbox, David\'s notification script is working correctly.',
  });
}
