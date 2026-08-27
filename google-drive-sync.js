// RNTU - Google Drive & Google Sheets Real-Time Synchronization Module
// Rabindranath Tagore University, Bhopal (A part of AISECT India)

export const GDriveSync = {
    STORAGE_KEY: 'rntu_gdrive_sync_url',
    STATUS_KEY: 'rntu_gdrive_last_sync',

    DEFAULT_URL: 'https://script.google.com/macros/s/AKfycbzbsB8xk0szjUei6QwEu4c_ZSQSgML1iUT1mpNmIa4QDaQlaoqRBRZGShwkz37Tjgdf/exec',

    getSyncUrl() {
        return localStorage.getItem(this.STORAGE_KEY) || this.DEFAULT_URL;
    },

    setSyncUrl(url) {
        if (url && url.trim()) {
            localStorage.setItem(this.STORAGE_KEY, url.trim());
            return true;
        }
        localStorage.setItem(this.STORAGE_KEY, this.DEFAULT_URL);
        return true;
    },

    getLastSyncTime() {
        return localStorage.getItem(this.STATUS_KEY) || 'Never synced';
    },

    setLastSyncTime() {
        const now = new Date().toLocaleString('en-IN', {
            dateStyle: 'medium',
            timeStyle: 'short'
        });
        localStorage.setItem(this.STATUS_KEY, now);
        return now;
    },

    /**
     * Helper to compute standard RNTU Grade from C, B, B+, A, A+
     */
    calculateGrade(cgpa, mockOrPercent) {
        const score = parseFloat(cgpa) ? (parseFloat(cgpa) * 9.5) : (parseFloat(mockOrPercent) || 70);
        if (score >= 85) return 'A+';
        if (score >= 75) return 'A';
        if (score >= 65) return 'B+';
        if (score >= 55) return 'B';
        return 'C';
    },

    /**
     * Send single action (insert, update, delete) or full batch to Google Apps Script
     */
    async pushToGoogleDrive(action, data) {
        const scriptUrl = this.getSyncUrl();
        if (!scriptUrl) {
            console.warn('[Google Drive Sync] No Web App URL configured. Changes saved locally/Firestore only.');
            return { success: false, reason: 'NO_URL' };
        }

        try {
            const payload = {
                timestamp: new Date().toISOString(),
                source: 'RNTU Students Progress Dashboard',
                action: action, // 'ADD', 'UPDATE', 'DELETE', 'BULK_SYNC'
                data: data
            };

            await fetch(scriptUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'text/plain;charset=utf-8'
                },
                body: JSON.stringify(payload)
            });

            this.setLastSyncTime();
            console.log(`[Google Drive Sync] ${action} event dispatched to Google Apps Script.`);
            return { success: true, timestamp: this.getLastSyncTime() };
        } catch (error) {
            console.error('[Google Drive Sync Error]:', error);
            return { success: false, error: error.message };
        }
    },

    /**
     * Synchronize all current students in one go to Google Drive / Sheets
     */
    async syncAllStudents(studentsArray) {
        return await this.pushToGoogleDrive('BULK_SYNC', studentsArray);
    },

    /**
     * Complete copy-paste Google Apps Script code for the user to deploy
     */
    getGoogleAppsScriptCode() {
        return `/**
 * ====================================================================
 * RNTU STUDENT LIFECYCLE & PROGRESS DASHBOARD - GOOGLE DRIVE SYNC
 * Rabindranath Tagore University, Bhopal (A part of AISECT India)
 * ====================================================================
 * 
 * INSTRUCTIONS TO DEPLOY:
 * 1. Open Google Drive (https://drive.google.com)
 * 2. Create a new Google Sheet named "RNTU_Student_Progress_Master_Data"
 * 3. In the Sheet, click "Extensions" -> "Apps Script"
 * 4. Replace all code in the script editor with THIS ENTIRE SCRIPT
 * 5. Click "Deploy" (top right) -> "New deployment"
 * 6. Select type: "Web app"
 * 7. Description: "RNTU Student Sync API"
 * 8. Execute as: "Me"
 * 9. Who has access: "Anyone" (allows dashboard to send data)
 * 10. Click "Deploy", Authorize access, and copy the "Web app URL"
 * 11. Paste that Web App URL in your RNTU Dashboard -> "Google Drive Sync" settings!
 */

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "active",
    message: "RNTU Student Progress Sync API is active & online!",
    time: new Date().toISOString()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var payload = JSON.parse(rawData);
    var action = payload.action;
    var student = payload.data;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RNTU_Master_Records");
    if (!sheet) {
      sheet = ss.insertSheet("RNTU_Master_Records");
      setupHeaders(sheet);
    }
    
    if (action === "ADD" || action === "UPDATE") {
      upsertStudent(sheet, student);
    } else if (action === "DELETE") {
      deleteStudent(sheet, student.id || student.enrollmentNo || student.name);
    } else if (action === "BULK_SYNC") {
      sheet.clearContents();
      setupHeaders(sheet);
      if (Array.isArray(student)) {
        student.forEach(function(st) {
          appendStudentRow(sheet, st);
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      action: action,
      syncedAt: new Date().toISOString()
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function setupHeaders(sheet) {
  var headers = [
    "ID / Roll No", "Full Name", "Student Overall Grade (C, B, B+, A, A+)", "Google Drive Portfolio Folder Link", "Email", "Mobile", "Gender",
    "Sector No", "Sector Name", "Course / Branch", "Current Year", "Semester", "Section", "Mentor Guide",
    "10th %", "12th / Diploma %", "Entrance Score", "Current CGPA",
    "Sem 1 Mock Aptitude", "Sem 1 Mock Tech", "Sem 1 Mock Comm", "Sem 1 Mock Overall", "Mock Date", "Mock Remarks",
    "Personality Communication", "Personality Presentation", "Personality Leadership", "Personality ProblemSolving", "Personality Etiquette", "Overall Personality Score",
    "Certifications", "Career Aspiration Goal", "Target Details",
    "Placement Status", "Company Name", "Designation", "Package (CTC LPA)", "Job Location",
    "Alumni Current Org", "Alumni Designation", "Alumni Experience", "Alumni LinkedIn",
    "Last Updated"
  ];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setBackground("#0d2c6c").setFontColor("#ffffff").setFontWeight("bold");
  sheet.setFrozenRows(1);
}

function mapStudentToRow(st) {
  var certs = "";
  if (Array.isArray(st.certifications)) {
    certs = st.certifications.map(function(c) { return (c.title || '') + " (" + (c.issuer || '') + ")"; }).join("; ");
  } else if (typeof st.certifications === "string") {
    certs = st.certifications;
  }
  
  return [
    st.enrollmentNo || st.id || "",
    st.name || st.fullName || "",
    st.grade || "A",
    st.drivePortfolioUrl || "",
    st.email || "",
    st.mobile || "",
    st.gender || "",
    st.sectorId || "",
    st.sectorName || "",
    st.course || st.branch || "",
    st.currentYear || "",
    st.currentSemester || "",
    st.section || "",
    st.mentorName || "",
    st.tenthMarks || "",
    st.twelfthMarks || "",
    st.entranceScore || "",
    st.cgpa || "",
    st.mockAptitude || "",
    st.mockTech || "",
    st.mockComm || "",
    st.mockOverall || "",
    st.mockDate || "",
    st.mockRemarks || "",
    st.scoreComm || "",
    st.scorePresentation || "",
    st.scoreLeadership || "",
    st.scoreProblemSolving || "",
    st.scoreEtiquette || "",
    st.personalityScore || "",
    certs,
    st.careerAspiration || "",
    st.targetDetails || "",
    st.placementStatus || "",
    st.placedCompany || "",
    st.placedRole || "",
    st.placedPackage || "",
    st.placedLocation || "",
    st.alumniOrg || "",
    st.alumniRole || "",
    st.alumniExp || "",
    st.alumniLinkedIn || "",
    new Date().toLocaleString()
  ];
}

function appendStudentRow(sheet, st) {
  sheet.appendRow(mapStudentToRow(st));
}

function upsertStudent(sheet, st) {
  var data = sheet.getDataRange().getValues();
  var searchKey = String(st.enrollmentNo || st.id || st.name).toLowerCase();
  var foundRow = -1;
  
  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0]).toLowerCase();
    var rowName = String(data[i][1]).toLowerCase();
    if (rowId === searchKey || (searchKey && rowName === searchKey)) {
      foundRow = i + 1;
      break;
    }
  }
  
  var newRow = mapStudentToRow(st);
  if (foundRow > 1) {
    sheet.getRange(foundRow, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }
}

function deleteStudent(sheet, idOrKey) {
  var data = sheet.getDataRange().getValues();
  var searchKey = String(idOrKey).toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var rowId = String(data[i][0]).toLowerCase();
    var rowName = String(data[i][1]).toLowerCase();
    if (rowId === searchKey || rowName === searchKey) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
`;
    },

    /**
     * Download comprehensive CSV file directly
     */
    downloadExcelCSV(students, filename = 'RNTU_Students_Master_Records.csv') {
        if (!students || students.length === 0) {
            alert('No student records available to export.');
            return;
        }

        const headers = [
            "Enrollment / Roll No", "Full Name", "Overall Grade (C, B, B+, A, A+)", "Google Drive Portfolio Folder Link", "Email", "Mobile", "Gender",
            "Sector ID", "Sector Name", "Course / Branch", "Current Year", "Semester", "Section", "Mentor Guide",
            "10th %", "12th / Diploma %", "Entrance Score", "Current CGPA",
            "Sem 1 Mock Aptitude (100)", "Sem 1 Mock Tech (100)", "Sem 1 Mock Comm (100)", "Sem 1 Mock Overall (100)", "Mock Date", "Mock Remarks", "Key Strengths", "Improvement Areas",
            "Comm Score", "Presentation Score", "Leadership Score", "Problem Solving Score", "Etiquette Score", "Overall Personality Score (100)", "Personality Readiness",
            "Certifications",
            "Career Aspiration", "Target Role / Details",
            "Placement Status", "Company Name", "Designation", "Package (CTC LPA)", "Job Location", "Offer Letter Received",
            "Alumni Organization", "Alumni Designation", "Alumni Total Experience (Yrs)", "Alumni LinkedIn Profile", "Alumni Willing To Mentor"
        ];

        const escapeCSV = (val) => {
            if (val === null || val === undefined) return '""';
            const str = String(val).replace(/"/g, '""');
            return `"${str}"`;
        };

        const rows = students.map(st => {
            let certs = '';
            if (Array.isArray(st.certifications)) {
                certs = st.certifications.map(c => `${c.title || ''} (${c.issuer || ''})`).join('; ');
            } else {
                certs = st.certifications || '';
            }

            return [
                escapeCSV(st.enrollmentNo || st.id || ''),
                escapeCSV(st.name || st.fullName || ''),
                escapeCSV(st.grade || 'A'),
                escapeCSV(st.drivePortfolioUrl || ''),
                escapeCSV(st.email || ''),
                escapeCSV(st.mobile || ''),
                escapeCSV(st.gender || ''),
                escapeCSV(st.sectorId || ''),
                escapeCSV(st.sectorName || ''),
                escapeCSV(st.course || st.branch || ''),
                escapeCSV(st.currentYear || ''),
                escapeCSV(st.currentSemester || ''),
                escapeCSV(st.section || ''),
                escapeCSV(st.mentorName || ''),
                escapeCSV(st.tenthMarks || ''),
                escapeCSV(st.twelfthMarks || ''),
                escapeCSV(st.entranceScore || ''),
                escapeCSV(st.cgpa || ''),
                escapeCSV(st.mockAptitude || ''),
                escapeCSV(st.mockTech || ''),
                escapeCSV(st.mockComm || ''),
                escapeCSV(st.mockOverall || ''),
                escapeCSV(st.mockDate || ''),
                escapeCSV(st.mockRemarks || ''),
                escapeCSV(st.mockStrengths || ''),
                escapeCSV(st.mockImprovements || ''),
                escapeCSV(st.scoreComm || ''),
                escapeCSV(st.scorePresentation || ''),
                escapeCSV(st.scoreLeadership || ''),
                escapeCSV(st.scoreProblemSolving || ''),
                escapeCSV(st.scoreEtiquette || ''),
                escapeCSV(st.personalityScore || ''),
                escapeCSV(st.personalityReadiness || ''),
                escapeCSV(certs),
                escapeCSV(st.careerAspiration || ''),
                escapeCSV(st.targetDetails || ''),
                escapeCSV(st.placementStatus || ''),
                escapeCSV(st.placedCompany || ''),
                escapeCSV(st.placedRole || ''),
                escapeCSV(st.placedPackage || ''),
                escapeCSV(st.placedLocation || ''),
                escapeCSV(st.offerLetterReceived || ''),
                escapeCSV(st.alumniOrg || ''),
                escapeCSV(st.alumniRole || ''),
                escapeCSV(st.alumniExp || ''),
                escapeCSV(st.alumniLinkedIn || ''),
                escapeCSV(st.alumniMentor || '')
            ].join(',');
        });

        const csvContent = "\uFEFF" + [headers.map(h => `"${h}"`).join(','), ...rows].join('\r\n');
        const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }
};
