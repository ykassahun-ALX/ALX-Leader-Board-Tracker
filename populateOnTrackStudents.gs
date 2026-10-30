/**
 * Automatically populates program tabs with On-Track students.
 *
 * Output:
 * Column A: Full Name
 * Column B: eHub Class Name
 *
 * Selection rules:
 * - Prioritizes the highest stages.
 * - Selects complete eHub class groups.
 * - Targets 10-20 students.
 * - Can exceed 20 to include an entire class.
 * - Never splits students sharing the same eHub Class Name.
 * - Some programs have classes that are always left out (see EXCLUDE_CLASSES).
 * - Each student appears only once, under their highest class.
 */

// Classes to leave out, per program (program name = tab name without the "_Oct 23" date).
// Founder Academy and Freelancer Academy: skip the WALX class and the first-level class.
// "(?!\d)" stops FA-1 from also matching FA-10, FA-12, etc.
const EXCLUDE_CLASSES = {
  "Founder Academy":    [/^WALX/i, /^FA-1(?!\d)/i],
  "Freelancer Academy": [/^WALX/i, /^FLA-1(?!\d)/i]
};

function populateOnTrackStudents() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName("Raw Data");

  if (!rawSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Raw Data' sheet not found!");
    return;
  }

  const data = rawSheet.getDataRange().getValues();

  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert("No data found in Raw Data sheet.");
    return;
  }

  // 0-indexed column positions
  const EMAIL_IDX = 0;           // Column A
  const LAST_NAME_IDX = 1;       // Column B
  const FIRST_NAME_IDX = 2;      // Column C
  const EHUB_CLASS_IDX = 24;     // Column Y
  const PROGRAM_NAME_IDX = 25;   // Column Z

  // Group records by Program Name
  const programData = {};

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    const program = String(row[PROGRAM_NAME_IDX] || "").trim();
    const firstName = String(row[FIRST_NAME_IDX] || "").trim();
    const lastName = String(row[LAST_NAME_IDX] || "").trim();
    const classCode = String(row[EHUB_CLASS_IDX] || "").trim();
    const email = String(row[EMAIL_IDX] || "").trim().toLowerCase();

    if (!program || !classCode) continue;

    if (!programData[program]) {
      programData[program] = [];
    }

    // Extract stage number
    // Examples: AiCE-6_rolling = 6, GD-10_rolling = 10
    const match = classCode.match(/\d+/);
    const stageNum = match ? parseInt(match[0], 10) : 0;

    const fullName = (firstName + " " + lastName).trim();

    programData[program].push({
      // Identifies the person: email if there is one, otherwise their name
      key: email || fullName.toLowerCase(),
      fullName: fullName,
      className: classCode,
      stage: stageNum
    });

  }

  // Process all program tabs
  const sheets = ss.getSheets();

  sheets.forEach(sheet => {

    const sheetName = sheet.getName();

    if (sheetName === "Raw Data") return;

    // Remove date suffix to identify program
    // Example: AiCE_Oct 23 = AiCE
    const programKey = sheetName
      .replace(/_[A-Za-z]+ \d+$/, "")
      .trim();

    let records = programData[programKey] || programData[sheetName];

    if (!records || records.length === 0) return;

    // Drop excluded classes for this program (e.g. WALX and FA-1 for Founder Academy)
    const excludeRules = EXCLUDE_CLASSES[programKey] || EXCLUDE_CLASSES[sheetName];
    if (excludeRules) {
      records = records.filter(student =>
        !excludeRules.some(rule => rule.test(student.className))
      );
    }

    // Group students by their EXACT eHub Class Name
    const classGroups = {};

    records.forEach(student => {

      const className = student.className;

      if (!classGroups[className]) {
        classGroups[className] = {
          stage: student.stage,
          students: []
        };
      }

      classGroups[className].students.push(student);

    });

    // Convert groups into an array and sort by stage descending
    const sortedClasses = Object.entries(classGroups)
      .sort((a, b) => b[1].stage - a[1].stage);

    // Select complete class groups
    let selectedStudents = [];

    for (let i = 0; i < sortedClasses.length; i++) {

      const classGroup = sortedClasses[i][1];

      // Add EVERY student in this eHub class
      selectedStudents = selectedStudents.concat(classGroup.students);

      // Stop once we reach at least 10 students
      // Never split a class to enforce the 20-student target
      if (selectedStudents.length >= 10) {
        break;
      }

    }

    // One row per student: if someone is in two selected classes,
    // keep only their highest class (e.g. PF-5 beats DA-4).
    const best = {};

    selectedStudents.forEach(student => {
      if (!best[student.key] || student.stage > best[student.key].stage) {
        best[student.key] = student;
      }
    });

    selectedStudents = selectedStudents.filter(student => best[student.key] === student);

    // Prepare output: Full Name and eHub Class Name only
    const outputValues = selectedStudents.map(student => [
      student.fullName,
      student.className
    ]);

    // Clear previous data in Columns A, B and C
    const lastRow = Math.max(sheet.getLastRow(), 1);

    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
    }

    // Clear old third-column header
    sheet.getRange(1, 3).clearContent();

    // Set headers
    sheet.getRange(1, 1, 1, 2).setValues([
      ["Full Name", "eHub Class Name"]
    ]);

    // Populate Columns A and B
    if (outputValues.length > 0) {
      sheet.getRange(2, 1, outputValues.length, 2).setValues(outputValues);
    }

  });

  SpreadsheetApp.getUi().alert(
    "Successfully populated complete eHub class groups across all program tabs!"
  );

}
