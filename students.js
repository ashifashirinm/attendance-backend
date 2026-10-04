console.log("I AM USING THIS STUDENTS.JS");
const XLSX = require("xlsx");
const path = require("path");

// Excel file location
const filePath = path.join(
  __dirname,
  "../../analysis/data/Student_Attendance_Analysis.xlsx"
);

// Read Excel file
const workbook = XLSX.readFile(filePath);

// First sheet
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];

// Convert Excel to JSON
const data = XLSX.utils.sheet_to_json(sheet);

// Create student data
const students = data.map((row) => {
  const rollNo = String(row.Roll_No);

  const present = Number(row.P || 0);
  const absent = Number(row.A || 0);

  const total = present + absent;

  const attendance =
    total > 0
      ? Number(((present / total) * 100).toFixed(1))
      : 0;

  let status;

  if (attendance >= 90) {
    status = "Excellent";
  } else if (attendance >= 80) {
    status = "Good";
  } else if (attendance >= 75) {
    status = "Average";
  } else {
    status = "Low";
  }

  return {
    rollNo,
    present,
    absent,
    attendance,
    status,
    longestAbsence: Number(row.Longest_Absence || 0),
  };
});

console.log("NEW STUDENTS.JS LOADED");
console.log("FIRST STUDENT:", students[0]);

module.exports = students;