const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const csv = require("csv-parser");
const XLSX = require("xlsx");

const app = express();

app.use(cors());
app.use(express.json());

const PORT = 5000;

// ==========================================
// FILE PATHS
// ==========================================

// ==========================================
// FILE PATHS
// ==========================================

const analysisDataPath = path.join(
  __dirname,
  "data"
);

const studentAnalysisFile = path.join(
  analysisDataPath,
  "Student_Attendance_Analysis.xlsx"
);

const attendanceFile = path.join(
  analysisDataPath,
  "Attendance_July_September_2026_Final.xlsx"
);

const predictionFile = path.join(
  analysisDataPath,
  "Future_Attendance_Predictions.csv"
);

// ==========================================
// READ STUDENT ANALYSIS EXCEL
// ==========================================

function getStudents() {
  const workbook = XLSX.readFile(studentAnalysisFile);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(sheet);

  return data.map((student) => {
    const present = Number(student.P || 0);
    const absent = Number(student.A || 0);

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
  rollNo: String(student.Roll_No),
  present,
  absent,
  attendance,
  status,
  longestAbsence: Number(student.Longest_Absence || 0),
};
  });
}

// ==========================================
// GET LEAVE HISTORY
// ==========================================

function getLeaveDates(rollNo) {
  const workbook = XLSX.readFile(attendanceFile);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(sheet);

  return data
    .filter(
      (row) =>
        String(row.Roll_No).toUpperCase() === rollNo &&
        String(row.Status).toUpperCase() === "A"
    )
    .map((row) => {
  let date = row.Date;

  if (typeof date === "number") {
    const excelDate = XLSX.SSF.parse_date_code(date);

    date = `${String(excelDate.d).padStart(2, "0")}-${String(
      excelDate.m
    ).padStart(2, "0")}-${excelDate.y}`;
  }

  return {
    date,
    day: row.Day,
  };
});
}

// ==========================================
// GET ALL LEAVE HISTORY
// ==========================================

function getAllLeaveHistory() {
  const workbook = XLSX.readFile(attendanceFile);

  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];

  const data = XLSX.utils.sheet_to_json(sheet);

  return data
    .filter(
      (row) =>
        String(row.Status).toUpperCase() === "A"
    )
    .map((row) => {
      let date = row.Date;

      // Convert Excel date number to DD-MM-YYYY
      if (typeof date === "number") {
        const excelDate = XLSX.SSF.parse_date_code(date);

        date = `${String(excelDate.d).padStart(2, "0")}-${String(
          excelDate.m
        ).padStart(2, "0")}-${excelDate.y}`;
      }

      return {
        rollNo: String(row.Roll_No),
        date,
        day: row.Day,
      };
    });
}

// ==========================================
// HOME API
// ==========================================

app.get("/", (req, res) => {
  res.json({
    message: "Attendance Backend API is running successfully!",
  });
});

// ==========================================
// ALL LEAVE HISTORY API
// ==========================================

app.get("/api/leave-history", (req, res) => {
  try {
    const leaveHistory = getAllLeaveHistory();

    res.json(leaveHistory);
  } catch (error) {
    console.error("Error reading leave history:", error);

    res.status(500).json({
      message: "Unable to read leave history",
    });
  }
});

// ==========================================
// ALL STUDENTS API
// ==========================================

app.get("/api/students", (req, res) => {
  try {
    const students = getStudents();

    res.json(students);
  } catch (error) {
    console.error("Error reading student data:", error);

    res.status(500).json({
      message: "Unable to read student data",
    });
  }
});

// ==========================================
// INDIVIDUAL STUDENT API
// ==========================================

app.get("/api/students/:rollNo", (req, res) => {
  try {
    const rollNo = req.params.rollNo.toUpperCase();

    const students = getStudents();

    const student = students.find(
      (student) => student.rollNo === rollNo
    );

    if (!student) {
      return res.status(404).json({
        message: "Student not found",
      });
    }

    const leaveDates = getLeaveDates(rollNo);

    const studentDetails = {
      ...student,
      leaveDates,
    };

    res.json(studentDetails);
  } catch (error) {
    console.error("Error reading student details:", error);

    res.status(500).json({
      message: "Unable to read student details",
    });
  }
});

// ==========================================
// ALL PREDICTIONS API
// ==========================================

app.get("/api/predictions", (req, res) => {
  const predictions = [];

  if (!fs.existsSync(predictionFile)) {
    return res.status(404).json({
      message: "Prediction file not found",
    });
  }

  fs.createReadStream(predictionFile)
    .pipe(csv())
    .on("data", (row) => {
      predictions.push({
        date: String(row.Date || ""),
        rollNo: String(row.Roll_No || ""),
        day: String(row.Day || ""),
        dayNumber: Number(row.Day_Number || 0),

        previousAbsences: Number(
          row.Previous_Absences || 0
        ),

        recent7Absences: Number(
          row.Recent_7_Absences || 0
        ),

        recent7AbsenceRate: Number(
          row.Recent_7_Absence_Rate || 0
        ),

        recent3Absences: Number(
          row.Recent_3_Absences || 0
        ),

        longestAbsence: Number(
          row.Longest_Absence || 0
        ),

        prediction: String(
          row.Predicted_Status || ""
        ).toUpperCase(),
      });
    })
    .on("end", () => {
      res.json(predictions);
    })
    .on("error", (error) => {
      console.error(error);

      res.status(500).json({
        message: "Unable to read prediction data",
      });
    });
});

// ==========================================
// INDIVIDUAL PREDICTION API
// ==========================================
app.get("/api/predictions/:rollNo", (req, res) => {
  const rollNo = req.params.rollNo.toUpperCase();
  const predictions = [];

  if (!fs.existsSync(predictionFile)) {
    return res.status(404).json({
      message: "Prediction file not found",
    });
  }

  fs.createReadStream(predictionFile)
    .pipe(csv())
    .on("data", (row) => {
      if (
        String(row.Roll_No).toUpperCase() === rollNo
      ) {
        predictions.push({
          date: String(row.Date || ""),
          rollNo: String(row.Roll_No || ""),
          day: String(row.Day || ""),
          dayNumber: Number(row.Day_Number || 0),
          previousAbsences: Number(
            row.Previous_Absences || 0
          ),
          recent7Absences: Number(
            row.Recent_7_Absences || 0
          ),
          recent7AbsenceRate: Number(
            row.Recent_7_Absence_Rate || 0
          ),
          recent3Absences: Number(
            row.Recent_3_Absences || 0
          ),
          longestAbsence: Number(
            row.Longest_Absence || 0
          ),
          prediction: String(
            row.Predicted_Status || ""
          ).toUpperCase(),
        });
      }
    })
    .on("end", () => {
      if (predictions.length === 0) {
        return res.status(404).json({
          message: "Prediction not found",
        });
      }

      res.json(predictions);
    })
    .on("error", (error) => {
      console.error(error);

      res.status(500).json({
        message: "Unable to read prediction data",
      });
    });
});



// ==========================================
// START SERVER
// ==========================================

app.listen(PORT, () => {
  console.log(
    `Backend server running on http://localhost:${PORT}`
  );
});