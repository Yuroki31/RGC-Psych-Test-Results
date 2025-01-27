function sendFormattedEmails() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  // Get headers
  const headers = data[0];
  const emailIndex = headers.indexOf("Email Address");
  const nameIndex = headers.indexOf("Name");
  const aqProfileIndex = headers.indexOf("Verbal Interpretation");
  const controlIndex = headers.indexOf("Control");
  const controlIndexverb = headers.indexOf("Control Verbal");
  const ownershipIndex = headers.indexOf("Ownership");
  const ownershipIndexverb = headers.indexOf("Ownership Verbal");
  const reachIndex = headers.indexOf("Reach");
  const reachIndexverb = headers.indexOf("Reach Verbal");
  const enduranceIndex = headers.indexOf("Endurance");
  const enduranceIndexverb = headers.indexOf("Endurance Verbal");

  // Your personal email for testing
  const testEmail = "sps.guidancecollege@uphsl.edu.ph";
  
  // Loop through the first 5 rows (students)
  for (let i = 1; i <= 5; i++) {  // Loop from 1 to 5 to limit to the first 5 students
    const row = data[i];
    
    const name = row[nameIndex];
    const aqProfile = row[aqProfileIndex];
    const control = row[controlIndex] + " " + row[controlIndexverb];
    const ownership = row[ownershipIndex] + " " + row[ownershipIndexverb];
    const reach = row[reachIndex] + " " + row[reachIndexverb];
    const endurance = row[enduranceIndex] + " " + row[enduranceIndexverb];

    // HTML Content
    const htmlBody = `
      <div style="font-family: Arial, sans-serif; line-height: 1.6;">
        <p>Good day ${name},</p>
        <p>I am Sir Del Ocampo, one of the UPHSL-RGC. I am sending this email to share with you your psycho-social test results:</p>
        <p>If you answered each item in all tests with full honesty, then you will truly benefit from these results. To wit, here are your scores and their respective verbal interpretations:</p>

        <table style="border-collapse: collapse; width: 100%; margin: 20px 0;">
          <thead>
            <tr style="background-color: #ffd700; text-align: center;">
              <th style="border: 1px solid #ddd; padding: 8px;">Adversity Quotient (AQ)</th>
              <th style="border: 1px solid #ddd; padding: 8px;">Control</th>
              <th style="border: 1px solid #ddd; padding: 8px;">Ownership</th>
              <th style="border: 1px solid #ddd; padding: 8px;">Reach</th>
              <th style="border: 1px solid #ddd; padding: 8px;">Endurance</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td style="border: 1px solid #ddd; padding: 8px;">AQ Profile = ${aqProfile}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${control}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${ownership}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${reach}</td>
              <td style="border: 1px solid #ddd; padding: 8px;">${endurance}</td>
            </tr>
          </tbody>
        </table>

        <p style="color: #d2691e;"><b>NOTE:</b> These psycho-social test results are vital in understanding your personal growth, leadership potential, and overall well-being. Keep working on these areas to improve your skills and achieve success!</p>

        <p>God bless you in conquering your career milestones and reaching your life goals!</p>
      </div>
    `;

    // Send the email to your personal email for testing
    GmailApp.sendEmail(testEmail, "Your Psycho-Social Test Results", "This is the plain text version of the email.", {
      htmlBody: htmlBody
    });

    Logger.log(`Test email sent to ${testEmail}`);
  }
}
