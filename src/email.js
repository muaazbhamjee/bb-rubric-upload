// ─── Configuration ────────────────────────────────────────────────────────
const FOLDER_ID = "PASTE_YOUR_FOLDER_ID_HERE";
const STUDENT_EMAIL_DOMAIN = "tuks.co.za";
const SENDER_NAME = "Prof M. Bhamjee";
const EMAIL_SUBJECT = "MKM411 Semester Test 2 — Marking Rubric";
const EMAIL_BODY =
  "Dear Student,\n\n" +
  "Please find attached your marking rubric for MKM411 Semester Test 2.\n\n" +
  "If you have any queries regarding your marks, please contact me during consultation hours.\n\n" +
  "Please take note of the below critical information regarding the marking rubric versus the mark and GA5 Competency Assessment on clickUP:\n" +
  "1. The marks were adjusted up by 8%. The rubric reflects the mark pre-adjustment. The mark on clickUP reflects the adjusted mark.\n" +
  "2. In terms of assessing GA5 Competency, the rubric provides the original assessment criteria. If you met those criteria which is 4 or greater then you have passed the GA5 component.\n" +
  "3. We removed the GA5-4 and GA5-7 criteria from the rubric. Thus, the revised criteria requires you to obtain 3 or greater then you have passed the GA5 component.\n\" +
  "4. The rubric reflects the original GA5 assessment criteria. The ST2 GA5 column in clickUP reflects the adjusted criteria.\n" +
  "5. On clickUP, the GA5 Competency Assessment is listed as 'ST2 GA5' and is separate from the mark. The mark reflects your overall mark for the test, while the 'ST2 GA5' column indicates whether you passed the GA5 competency based on the adjusted criteria.\n\n" +
  "6. On clickUP for the ST2 GA5 column, a mark of 1/1 means you are deemed 'Competent' and that means you passed the GA5 component based on the adjusted criteria, while a mark of 0/1 means you are deemed 'Not Yet Competent' which means you did not meet the adjusted criteria for passing the GA5 component.\n\n" +
  "Please review your rubric and the information above carefully. 
  "Kind regards,\n" +
  "Prof M. Bhamjee\n" +
  "Department of Mechanical and Aeronautical Engineering\n" +
  "University of Pretoria";

// Set to true for first run — sends only to yourself
const TEST_MODE = true;
const TEST_EMAIL = "muaaz.bhamjee@up.ac.za";
// ──────────────────────────────────────────────────────────────────────────

/**
 * DRY RUN — logs what would be sent without sending anything.
 * Run this first to verify the matching is correct.
 */
function dryRun() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let count = 0;

  Logger.log("=== DRY RUN ===\n");

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const match = name.match(/^feedback_(u\d+)\.pdf$/i);

    if (!match) {
      Logger.log("SKIP: " + name + " (no username found)");
      continue;
    }

    const username = match[1];
    const studentEmail = username + "@" + STUDENT_EMAIL_DOMAIN;
    Logger.log(username + " -> " + studentEmail + "  [" + name + "]");
    count++;
  }

  Logger.log("\nTotal: " + count + " emails would be sent");
}

/**
 * SEND EMAILS — respects TEST_MODE flag.
 * Set TEST_MODE = true to send one email to yourself first.
 * Set TEST_MODE = false to send to all students.
 */
function sendRubrics() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();
  let success = 0;
  let failed = 0;
  let skipped = 0;

  const mode = TEST_MODE ? "TEST" : "LIVE";
  Logger.log("=== " + mode + " MODE ===\n");

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const match = name.match(/^feedback_(u\d+)\.pdf$/i);

    if (!match) {
      Logger.log("SKIP: " + name);
      skipped++;
      continue;
    }

    const username = match[1];
    const recipient = TEST_MODE ? TEST_EMAIL : username + "@" + STUDENT_EMAIL_DOMAIN;
    const subject = EMAIL_SUBJECT + " (" + username + ")";

    try {
      GmailApp.sendEmail(recipient, subject, EMAIL_BODY, {
        attachments: [file.getBlob()],
        name: SENDER_NAME
      });
      Logger.log("[OK] " + username + " -> " + recipient);
      success++;
    } catch (e) {
      Logger.log("[FAIL] " + username + " -> " + recipient + " (" + e.message + ")");
      failed++;
    }

    // In test mode, only send one
    if (TEST_MODE) {
      Logger.log("\nTest email sent to " + TEST_EMAIL);
      Logger.log("Check your inbox. If it looks good, set TEST_MODE = false and run again.");
      return;
    }

    // Small delay to avoid rate limits
    Utilities.sleep(500);
  }

  Logger.log("\nDone: " + success + " sent, " + failed + " failed, " + skipped + " skipped");
}

/**
 * VERIFICATION — creates a Google Sheet listing every email-rubric pairing.
 * Run this before sending to confirm the mapping is correct.
 * Check that:
 *   1. Every student number maps to the right email
 *   2. File sizes vary (confirms each PDF is unique, not duplicates)
 *   3. The total count is correct
 */
function createVerificationSheet() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = folder.getFiles();

  const ss = SpreadsheetApp.create("MKM411_ST2_Email_Verification");
  const sheet = ss.getActiveSheet();
  sheet.appendRow(["Username", "Email", "Filename", "File Size (KB)"]);
  sheet.getRange("A:A").setNumberFormat("@");

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const match = name.match(/^feedback_(u\d+)\.pdf$/i);
    if (!match) continue;

    const username = match[1];
    const email = username + "@" + STUDENT_EMAIL_DOMAIN;
    const sizeKB = (file.getSize() / 1024).toFixed(1);

    sheet.appendRow([username, email, name, sizeKB]);
  }

  Logger.log("Verification sheet created: " + ss.getUrl());
}
