// Paste this into Google Apps Script opened from within the Google Sheet
// (Extensions > Apps Script), then deploy as a Web App:
//   Execute as: Me
//   Who has access: Anyone
// Paste the deployment URL into index.html as the APPS_SCRIPT_URL value.

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Timestamp', 'Name', 'Email', 'Phone',
        'Participation', 'Accommodation', 'Meals',
        'Gluten Free', 'Special Needs'
      ]);
    }

    const now = new Date();
    sheet.appendRow([
      Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
      data.name          || '',
      data.email         || '',
      data.phone         || '',
      data.participation || '',
      data.accommodation || '',
      data.meals         || '',
      data.gluten_free   || '',
      data.special_needs || '',
    ]);

    const mealCosts = { 'Friday dinner ($8)': 8, 'Saturday lunch ($25)': 25, 'Saturday dinner ($32)': 32 };
    const selectedMeals = data.meals && data.meals !== 'None selected'
      ? data.meals.split(', ').filter(m => mealCosts[m] !== undefined)
      : [];
    const total = selectedMeals.reduce((sum, m) => sum + mealCosts[m], 0);

    const mealsSection = selectedMeals.length > 0
      ? 'You have selected the following meals:\n'
        + selectedMeals.map(m => `  ${m}`).join('\n')
        + `\n  ──────────────\n  Total: $${total}\n\n`
        + 'Please include this amount in your payment.\n\n'
      : 'You have not selected any meals.\n\n';

    // Confirmation email to registrant
    MailApp.sendEmail({
      to:      data.email,
      subject: 'Your registration — Who, What, Where is God?',
      body:
        `Dear ${data.name.split(' ')[0]},\n\n`
        + 'Thank you for registering for our seminar "Who, What, Where is God?" '
        + 'at Wellington Meeting House, 11–13 September 2026. We look forward to seeing you there.\n\n'
        + mealsSection
        + 'Payment details:\n'
        + '  Account name:   Wellington Monthly Meeting Of The Religious Society Of Friends\n'
        + '  Account number: 38-9015-0775731-06\n'
        + '  Reference:      Your name\n\n'
        + 'If you have any questions, please contact Murray at mandns@xtra.co.nz.\n\n'
        + 'Warm regards,\n'
        + 'Murray Short & Janet McKean',
    });

    // Notification email to organiser
    const day   = Utilities.formatDate(now, Session.getScriptTimeZone(), 'd');
    const month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM');
    const time  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
    MailApp.sendEmail({
      to:      'quakerregistration6@gmail.com',
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

// Run this once from the editor to grant MailApp permission:
// select testEmail from the dropdown and click Run
function testEmail() {
  MailApp.sendEmail({
    to:      'quakerregistration6@gmail.com',
    subject: 'Apps Script authorisation test',
    body:    'If you received this, MailApp is authorised and forwarding is working.',
  });
}
