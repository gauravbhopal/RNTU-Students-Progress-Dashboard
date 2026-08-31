// RNTU - Google Drive & Google Sheets Real-Time Synchronization Module
// Rabindranath Tagore University, Bhopal (A part of AISECT India)

export const GDriveSync = {
    STORAGE_KEY: 'rntu_gdrive_sync_url',
    STATUS_KEY: 'rntu_gdrive_last_sync',

    DEFAULT_URL: 'https://script.google.com/macros/s/AKfycbxgrFZEbE0_JTPE99sgstm10CXDWzFThyR3Q7YHALstX0m4uq0GfDizZptni2FeE72LgA/exec',

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
     * Helper to compute standard RNTU Grade strictly based on 1st Sem Mock Diagnostic Score (%)
     * Grade A+ : >= 90% (Top Tier)
     * Grade A  : >= 80% and < 90% (Excellent)
     * Grade B+ : >= 70% and < 80% (Good Baseline)
     * Grade B  : >= 60% and < 70% (Average Baseline)
     * Grade C  : < 60% (Needs Focused Training)
     */
    calculateGrade(mockScore) {
        const score = parseFloat(mockScore) || 0;
        if (score >= 90) return 'A+';
        if (score >= 80) return 'A';
        if (score >= 70) return 'B+';
        if (score >= 60) return 'B';
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
     * Fetch all students from Google Sheets via Google Apps Script Web App
     */
    async fetchStudentsFromGoogleSheet() {
        const scriptUrl = this.getSyncUrl();
        if (!scriptUrl) return { success: false, error: 'NO_URL' };

        try {
            const url = new URL(scriptUrl);
            url.searchParams.set('action', 'GET_ALL');
            url.searchParams.set('_t', Date.now());

            const response = await fetch(url.toString(), {
                method: 'GET',
                redirect: 'follow'
            });

            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }

            const res = await response.json();
            if (res && res.status === 'success' && Array.isArray(res.students)) {
                this.setLastSyncTime();
                return { success: true, students: res.students, count: res.students.length };
            } else {
                return { success: false, error: res.message || 'Invalid response from Google Apps Script' };
            }
        } catch (err) {
            console.error('[Google Drive Fetch Error]:', err);
            return { success: false, error: err.message };
        }
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
 * INSTRUCTIONS TO UPDATE / DEPLOY:
 * 1. Open your Google Sheet in Google Drive
 * 2. Click "Extensions" -> "Apps Script"
 * 3. Replace all code in the script editor with THIS ENTIRE SCRIPT
 * 4. Click "Deploy" -> "Manage deployments" -> Click Edit (pencil icon) -> Version: "New version" -> Click "Deploy"
 * 5. Your Web App URL is ready to sync live with your RNTU Dashboard!
 */

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("RNTU_Master_Records") || ss.getSheets()[0];
    var data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        students: [],
        count: 0
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    var students = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] && !row[1]) continue;
      
      var student = {
        id: String(row[0] || ('rntu_' + (Date.now() + i))),
        enrollmentNo: String(row[0] || ''),
        name: String(row[1] || ''),
        fullName: String(row[1] || ''),
        grade: String(row[2] || 'A'),
        drivePortfolioUrl: String(row[3] || ''),
        email: String(row[4] || ''),
        mobile: String(row[5] || ''),
        gender: String(row[6] || 'Male'),
        sectorId: String(row[7] || '1'),
        sectorName: String(row[8] || ''),
        course: String(row[9] || ''),
        currentYear: String(row[10] || '1st Year'),
        currentSemester: String(row[11] || '1'),
        section: String(row[12] || ''),
        tenthMarks: parseFloat(row[13]) || '',
        twelfthMarks: parseFloat(row[14]) || '',
        entranceScore: String(row[15] || ''),
        cgpa: parseFloat(row[16]) || '',
        mockAptitude: parseFloat(row[17]) || 70,
        mockTech: parseFloat(row[18]) || 68,
        mockComm: parseFloat(row[19]) || 72,
        mockConfidence: parseFloat(row[20]) || 75,
        mockOverall: parseFloat(row[21]) || 71.25,
        mockDate: row[22] ? String(row[22]) : '',
        mockStrengths: String(row[23] || ''),
        mockImprovements: String(row[24] || ''),
        mockRemarks: String(row[25] || ''),
        scoreComm: parseFloat(row[26]) || 75,
        scorePresentation: parseFloat(row[27]) || 70,
        scoreLeadership: parseFloat(row[28]) || 80,
        scoreProblemSolving: parseFloat(row[29]) || 78,
        scoreEtiquette: parseFloat(row[30]) || 82,
        scoreTechnical: parseFloat(row[31]) || 75,
        personalityScore: parseFloat(row[32]) || 76.7,
        personalityReadiness: String(row[33] || 'Placement Ready'),
        certifications: String(row[34] || ''),
        careerAspiration: String(row[35] || 'Job / Corporate Placement'),
        targetDetails: String(row[36] || ''),
        placementStatus: String(row[37] || 'Not Yet Placed'),
        placedCompany: String(row[38] || ''),
        placedRole: String(row[39] || ''),
        placedPackage: parseFloat(row[40]) || 0,
        placedLocation: String(row[41] || ''),
        offerLetterReceived: String(row[42] || 'No'),
        alumniOrg: String(row[43] || ''),
        alumniRole: String(row[44] || ''),
        alumniExp: parseFloat(row[45]) || 0,
        alumniLinkedIn: String(row[46] || ''),
        alumniMentor: String(row[47] || 'No'),
        lastUpdated: String(row[48] || '')
      };
      students.push(student);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      students: students,
      count: students.length
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
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
    "Enrollment / Roll No",
    "Full Name",
    "Student Overall Grade (C, B, B+, A, A+)",
    "Google Drive Portfolio Folder Link",
    "Email",
    "Mobile",
    "Gender",
    "Sector No",
    "Sector Name",
    "Course / Branch",
    "Current Year",
    "Semester",
    "Section / Batch",
    "10th %",
    "12th / Diploma %",
    "Entrance Score",
    "Current CGPA",
    "Sem 1 Mock Aptitude (100)",
    "Sem 1 Mock Tech (100)",
    "Sem 1 Mock Comm (100)",
    "Sem 1 Mock Confidence (100)",
    "Sem 1 Mock Overall Score (%)",
    "Mock Date",
    "Mock Key Strengths",
    "Mock Areas for Improvement",
    "Mock Evaluator Remarks",
    "Personality Communication",
    "Personality Presentation",
    "Personality Leadership",
    "Personality Problem Solving",
    "Personality Etiquette",
    "Technical Mastery",
    "Composite Personality Score (100)",
    "Personality Readiness Level",
    "Certifications & Portfolio",
    "Career Aspiration Goal",
    "Target Details",
    "Placement Status",
    "Recruiting Company",
    "Designation / Role",
    "Package (CTC LPA)",
    "Job Location",
    "Offer Letter Received",
    "Alumni Current Organization",
    "Alumni Designation",
    "Alumni Experience (Years)",
    "Alumni LinkedIn Profile",
    "Willing to Mentor",
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
    st.section || st.batchCode || "",
    st.tenthMarks || "",
    st.twelfthMarks || "",
    st.entranceScore || "",
    st.cgpa || "",
    st.mockAptitude || "",
    st.mockTech || "",
    st.mockComm || "",
    st.mockConfidence || "",
    st.mockOverall || "",
    st.mockDate || "",
    st.mockStrengths || "",
    st.mockImprovements || "",
    st.mockRemarks || "",
    st.scoreComm || "",
    st.scorePresentation || "",
    st.scoreLeadership || "",
    st.scoreProblemSolving || "",
    st.scoreEtiquette || "",
    st.scoreTechnical || "",
    st.personalityScore || "",
    st.personalityReadiness || "",
    certs,
    st.careerAspiration || "",
    st.targetDetails || "",
    st.placementStatus || "",
    st.placedCompany || "",
    st.placedRole || "",
    st.placedPackage || "",
    st.placedLocation || "",
    st.offerLetterReceived || "",
    st.alumniOrg || "",
    st.alumniRole || "",
    st.alumniExp || "",
    st.alumniLinkedIn || "",
    st.alumniMentor || "",
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
            "Sector ID", "Sector Name", "Course / Branch", "Current Year", "Semester", "Section",
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
    },

    /**
     * Download a clean pre-formatted Sample CSV Template for Bulk Upload
     */
    downloadSampleTemplateCSV(filename = 'RNTU_Student_Bulk_Upload_Sample_Template.csv') {
        const headers = [
            "Enrollment / Roll No",
            "Full Name",
            "Sector ID (1-6)",
            "Course / Branch",
            "Current Year",
            "Current Semester",
            "Section",
            "Email",
            "Mobile",
            "Gender",
            "10th %",
            "12th %",
            "Entrance Score",
            "Current CGPA",
            "Sem 1 Mock Aptitude (100)",
            "Sem 1 Mock Tech (100)",
            "Sem 1 Mock Comm (100)",
            "Sem 1 Mock Confidence (100)",
            "Key Strengths",
            "Improvement Areas",
            "Career Aspiration",
            "Target Details",
            "Placement Status",
            "Company Name",
            "Designation",
            "Package (CTC LPA)",
            "Job Location",
            "Google Drive Portfolio Folder Link"
        ];

        const sampleRows = [
            [
                "0101RNTU23001",
                "Aman Sharma",
                "1",
                "BE Computer Science",
                "3rd Year",
                "5",
                "Batch 2023-27 / Sec-A",
                "aman.sharma@rntu.ac.in",
                "9826012345",
                "Male",
                "86.5",
                "82.0",
                "JEE 88.5%ile",
                "8.65",
                "82",
                "85",
                "78",
                "80",
                "Strong analytical logic, proficient in C++ & Data Structures",
                "Advanced system architecture, public speaking",
                "Job / Corporate Placement",
                "Software Development Engineer in MNCs",
                "Placed",
                "TCS Digital",
                "Systems Engineer",
                "7.5",
                "Pune",
                "https://drive.google.com/drive/folders/sample-portfolio-aman"
            ],
            [
                "0102RNTU23002",
                "Priya Patel",
                "2",
                "BCA (Data Science)",
                "3rd Year",
                "5",
                "Batch 2023-26",
                "priya.patel@rntu.ac.in",
                "9425098765",
                "Female",
                "89.2",
                "85.6",
                "Merit Rank 14",
                "8.90",
                "88",
                "82",
                "85",
                "86",
                "High mathematical clarity, fluent communicator, Python enthusiast",
                "Big Data distributed tools",
                "Higher Studies",
                "MS in Data Science / GATE 2026",
                "Higher Studies",
                "",
                "",
                "0",
                "",
                "https://drive.google.com/drive/folders/sample-portfolio-priya"
            ],
            [
                "0106RNTU23003",
                "Rahul Chouhan",
                "6",
                "B.Sc. Agriculture (Sec-A)",
                "4th Year",
                "7",
                "Batch 2022-26 / Sec-A",
                "rahul.chouhan@rntu.ac.in",
                "9754011223",
                "Male",
                "78.4",
                "76.5",
                "PAT Rank 120",
                "8.20",
                "74",
                "80",
                "70",
                "75",
                "Deep practical understanding of agronomy and soil health",
                "Agri-marketing skills",
                "Start Own / Startup",
                "Hydroponics & Organic Farm Export Enterprise",
                "Entrepreneur",
                "AgriGreen Organics",
                "Co-Founder & Director",
                "8.0",
                "Bhopal",
                "https://drive.google.com/drive/folders/sample-portfolio-rahul"
            ]
        ];

        const escapeCSV = (val) => `"${String(val || '').replace(/"/g, '""')}"`;
        const content = "\uFEFF" + [
            headers.map(escapeCSV).join(','),
            ...sampleRows.map(row => row.map(escapeCSV).join(','))
        ].join('\r\n');

        const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        link.setAttribute('download', filename);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    },

    /**
     * Parse raw CSV text into array of structured Student objects
     */
    parseCSVToStudents(csvText) {
        if (!csvText || !csvText.trim()) return [];

        const SECTOR_NAME_MAP = {
            '1': 'Sector 1: Engineering & Technology',
            '2': 'Sector 2: Computer Applications & IT',
            '3': 'Sector 3: Commerce & Management',
            '4': 'Sector 4: Life & Physical Sciences',
            '5': 'Sector 5: Pharmacy',
            '6': 'Sector 6: Agriculture'
        };

        // Standard CSV Splitter handling quotes and multi-line cells
        const rows = [];
        let currentRow = [];
        let currentCell = '';
        let insideQuotes = false;

        const cleanText = csvText.replace(/\r\n/g, '\n').replace(/\r/g, '\n');

        for (let i = 0; i < cleanText.length; i++) {
            const char = cleanText[i];
            const nextChar = cleanText[i + 1];

            if (char === '"') {
                if (insideQuotes && nextChar === '"') {
                    currentCell += '"';
                    i++; // skip escaped quote
                } else {
                    insideQuotes = !insideQuotes;
                }
            } else if (char === ',' && !insideQuotes) {
                currentRow.push(currentCell.trim());
                currentCell = '';
            } else if (char === '\n' && !insideQuotes) {
                currentRow.push(currentCell.trim());
                if (currentRow.some(c => c !== '')) {
                    rows.push(currentRow);
                }
                currentRow = [];
                currentCell = '';
            } else {
                currentCell += char;
            }
        }
        if (currentCell || currentRow.length > 0) {
            currentRow.push(currentCell.trim());
            if (currentRow.some(c => c !== '')) {
                rows.push(currentRow);
            }
        }

        if (rows.length < 2) return [];

        const rawHeaders = rows[0].map(h => h.toLowerCase().replace(/[^a-z0-9]/g, ''));
        const students = [];

        for (let r = 1; r < rows.length; r++) {
            const row = rows[r];
            if (row.length === 0 || row.every(c => c === '')) continue;

            const getCol = (possibleNames) => {
                for (const name of possibleNames) {
                    const idx = rawHeaders.indexOf(name.toLowerCase().replace(/[^a-z0-9]/g, ''));
                    if (idx !== -1 && row[idx] !== undefined && row[idx] !== '') {
                        return row[idx];
                    }
                }
                return '';
            };

            const enrollmentNo = getCol(['enrollment', 'enrollmentno', 'rollno', 'roll', 'id', 'studentroll']);
            const name = getCol(['fullname', 'name', 'studentname', 'student']);
            if (!enrollmentNo && !name) continue;

            let sectorId = String(getCol(['sectorid', 'sector', 'sectorno']) || '1');
            if (sectorId.includes('1') || sectorId.toLowerCase().includes('eng')) sectorId = '1';
            else if (sectorId.includes('2') || sectorId.toLowerCase().includes('it') || sectorId.toLowerCase().includes('bca')) sectorId = '2';
            else if (sectorId.includes('3') || sectorId.toLowerCase().includes('comm') || sectorId.toLowerCase().includes('mgt')) sectorId = '3';
            else if (sectorId.includes('4') || sectorId.toLowerCase().includes('sci')) sectorId = '4';
            else if (sectorId.includes('5') || sectorId.toLowerCase().includes('pharm')) sectorId = '5';
            else if (sectorId.includes('6') || sectorId.toLowerCase().includes('agri')) sectorId = '6';
            else sectorId = '1';

            const cgpa = parseFloat(getCol(['cgpa', 'currentcgpa', 'sgpa'])) || '';
            const mockAptitude = parseFloat(getCol(['sem1mockaptitude', 'mockaptitude', 'aptitude'])) || 70;
            const mockTech = parseFloat(getCol(['sem1mocktech', 'mocktech', 'technical'])) || 68;
            const mockComm = parseFloat(getCol(['sem1mockcomm', 'mockcomm', 'communication'])) || 72;
            const mockConfidence = parseFloat(getCol(['sem1mockconfidence', 'mockconfidence', 'confidence'])) || 75;
            const mockOverall = ((mockAptitude + mockTech + mockComm + mockConfidence) / 4);

            let grade = getCol(['grade', 'overallgrade']);
            if (!grade) {
                grade = GDriveSync.calculateGrade(mockOverall);
            }

            const student = {
                id: enrollmentNo || ('rntu_' + (Date.now() + r)),
                enrollmentNo: enrollmentNo,
                name: name,
                fullName: name,
                grade: grade,
                drivePortfolioUrl: getCol(['driveportfoliofolderlink', 'driveportfoliourl', 'drivelink', 'portfolio', 'drivefolder']),
                email: getCol(['email', 'institutionalemail', 'mail']),
                mobile: getCol(['mobile', 'phone', 'whatsapp', 'contact']),
                gender: getCol(['gender', 'sex']) || 'Male',
                sectorId: sectorId,
                sectorName: SECTOR_NAME_MAP[sectorId] || 'Sector ' + sectorId,
                course: getCol(['coursebranch', 'course', 'branch', 'program']),
                currentYear: getCol(['currentyear', 'year']) || '1st Year',
                currentSemester: String(getCol(['currentsemester', 'semester', 'sem']) || '1'),
                section: getCol(['section', 'batch', 'sectionbatch']),
                tenthMarks: parseFloat(getCol(['10th', 'tenthmarks', '10thpercent'])) || '',
                twelfthMarks: parseFloat(getCol(['12th', 'twelfthmarks', '12thpercent'])) || '',
                entranceScore: getCol(['entrancescore', 'jee', 'cuet', 'pat']),
                cgpa: cgpa,
                mockAptitude: mockAptitude,
                mockTech: mockTech,
                mockComm: mockComm,
                mockConfidence: mockConfidence,
                mockOverall: parseFloat(mockOverall.toFixed(2)),
                mockDate: getCol(['mockdate', 'date']),
                mockStrengths: getCol(['keystrengths', 'mockstrengths', 'strengths']),
                mockImprovements: getCol(['improvementareas', 'mockimprovements', 'improvements']),
                mockRemarks: getCol(['evaluatorremarks', 'mockremarks', 'remarks']),
                scoreComm: parseFloat(getCol(['commscore', 'scorecomm'])) || 75,
                scorePresentation: parseFloat(getCol(['presentationscore', 'scorepresentation'])) || 70,
                scoreLeadership: parseFloat(getCol(['leadershipscore', 'scoreleadership'])) || 80,
                scoreProblemSolving: parseFloat(getCol(['problemsolvingscore', 'scoreproblemsolving'])) || 78,
                scoreEtiquette: parseFloat(getCol(['etiquettescore', 'scoreetiquette'])) || 82,
                scoreTechnical: parseFloat(getCol(['technicaldomainscore', 'scoretechnical'])) || 75,
                personalityScore: 76.7,
                personalityReadiness: 'Placement Ready',
                certifications: getCol(['certifications', 'certs']),
                careerAspiration: getCol(['careeraspiration', 'careeraspirationgoal', 'careergoal']) || 'Job / Corporate Placement',
                targetDetails: getCol(['targetdetails', 'targetrole', 'target']),
                placementStatus: getCol(['placementstatus', 'status']) || 'Not Yet Placed',
                placedCompany: getCol(['companyname', 'recruitingcompany', 'company']),
                placedRole: getCol(['designation', 'role', 'jobrole']),
                placedPackage: parseFloat(getCol(['packagectclpa', 'package', 'ctc', 'lpa'])) || 0,
                placedLocation: getCol(['joblocation', 'location', 'city']),
                offerLetterReceived: getCol(['offerletterreceived', 'offerletter']) || 'No',
                alumniOrg: getCol(['alumniorganization', 'alumniorg']),
                alumniRole: getCol(['alumnidesignation', 'alumnirole']),
                alumniExp: parseFloat(getCol(['alumniexperience', 'alumniexp'])) || 0,
                alumniLinkedIn: getCol(['alumnilinkedinprofile', 'alumnilinkedin', 'linkedin']),
                alumniMentor: getCol(['willingtomentor', 'alumnimentor']) || 'No',
                lastUpdated: new Date().toLocaleString()
            };

            // Compute composite personality
            const pAvg = (student.scoreComm + student.scorePresentation + student.scoreLeadership + student.scoreProblemSolving + student.scoreEtiquette + student.scoreTechnical) / 6;
            student.personalityScore = parseFloat(pAvg.toFixed(1));
            if (student.personalityScore >= 85) student.personalityReadiness = 'Exceptional / Top Talent';
            else if (student.personalityScore >= 70) student.personalityReadiness = 'Placement Ready';
            else if (student.personalityScore >= 50) student.personalityReadiness = 'Developing Aptitude';
            else student.personalityReadiness = 'Needs Focused Training';

            students.push(student);
        }

        return students;
    }
};

