import React, { useState } from 'react';
import { FileUp, Download, CheckCircle, AlertCircle, FileSpreadsheet, Clock, Users, Info } from 'lucide-react';
import XLSX from 'xlsx-js-style';

function App() {
  const [file, setFile] = useState(null);
  const [error, setError] = useState('');
  const [processedData, setProcessedData] = useState(null);
  const [stats, setStats] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  // --- HELPERS ---

  const formatOutputDate = (val) => {
    if (!val) return "";
    let dateObj = null;
    if (typeof val === 'number') {
      dateObj = new Date(Math.round((val - 25569) * 86400 * 1000));
      const offset = dateObj.getTimezoneOffset() * 60 * 1000;
      dateObj = new Date(dateObj.getTime() + offset);
    } else {
      dateObj = new Date(val);
    }

    if (!isNaN(dateObj.getTime())) {
      const day = String(dateObj.getDate()).padStart(2, '0');
      const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
      const month = monthNames[dateObj.getMonth()];
      const year = dateObj.getFullYear();
      return `${day}-${month}-${year}`;
    }
    return String(val);
  };

  const getDecimalHours = (value) => {
    if (value == null || value === '') return null;
    if (typeof value === 'number') return value * 24;

    const timeStr = String(value).trim();
    const match = timeStr.match(/^(\d{1,2})[:.](\d{2})(?:\s*:(\d{2}))?(?:\s*([AaPp][Mm]))?/);
    if (match) {
      let h = parseInt(match[1], 10);
      let m = parseInt(match[2], 10);
      const meridiem = match[4] ? match[4].toLowerCase() : null;
      if (meridiem === 'pm' && h < 12) h += 12;
      if (meridiem === 'am' && h === 12) h = 0;
      return h + (m / 60);
    }
    return null;
  };

  const formatTimeOutput = (val) => {
    if (val == null || val === '') return "";
    if (typeof val === 'number') {
      const totalMinutes = Math.round(val * 24 * 60);
      const h = Math.floor(totalMinutes / 60);
      const m = totalMinutes % 60;
      return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }
    return String(val);
  };

  const findValueToRight = (row, startIdx) => {
    for (let c = startIdx + 1; c < row.length; c++) {
      let cell = String(row[c] || "").trim();
      if (!cell || cell === ':' || cell === '-') continue;
      cell = cell.replace(/^[:\-\s]+/, '').trim();
      if (cell) return cell;
    }
    return null;
  };

  const handleFileUpload = (e) => {
    const uploadedFile = e.target.files[0];
    if (!uploadedFile) return;
    setFile(uploadedFile);
    setError('');
    setProcessedData(null);
    setStats(null);
  };

  const processFile = async () => {
    if (!file) return;
    setIsProcessing(true);
    setError('');

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);

      const employeeGroups = new Map();

      // Patterns
      // Patterns
      const patterns = {
        empCode: /Emp(?:loyee)?[\s\.]*Code/i,
        empName: /Emp(?:loyee)?[\s\.]*Name/i,
        date: /^(?:ATT\.?\s*DATE|DATE)$/i,
        inTime: /^(?:In\s*Time|InTime)$/i,
        outTime: /^(?:Out\s*Time|OutTime)$/i,
        status: /^Status$/i,
        punch: /Punch/i,
        shift: /^S\.?|Shift/i,
        department: /Department|Dept\.?/i
      };

      for (const sheetName of workbook.SheetNames) {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

        if (!rows || rows.length === 0) continue;

        let currentEmp = { code: null, name: null, department: null };
        let headerIndices = null;

        for (let r = 0; r < rows.length; r++) {
          const row = rows[r];
          if (!row || row.length === 0) continue;

          // 1. METADATA SCANNING
          let foundMetadata = false;
          for (let c = 0; c < row.length; c++) {
            const cellText = String(row[c] || "").trim();

            // Check Emp Code
            if (patterns.empCode.test(cellText)) {
              let val = null;
              if (cellText.match(/[:\-]\s*\S/)) {
                const parts = cellText.split(/[:\-]/);
                if (parts[1]) val = parts[1].trim();
              } else {
                val = findValueToRight(row, c);
              }
              if (val) {
                // If new code is found and different, reset context
                if (currentEmp.code && val !== currentEmp.code) {
                  headerIndices = null;
                  currentEmp = { code: null, name: null, department: null };
                }
                currentEmp.code = val;
                foundMetadata = true;
              }
            }

            // Check Emp Name
            if (patterns.empName.test(cellText)) {
              let val = null;
              if (cellText.match(/[:\-]\s*\S/)) {
                const parts = cellText.split(/[:\-]/);
                if (parts[1]) val = parts[1].trim();
              } else {
                val = findValueToRight(row, c);
              }
              if (val) {
                currentEmp.name = val.replace(/\s+/g, ' ').toUpperCase();
                foundMetadata = true;
              }
            }
          }

          if (foundMetadata) continue;

          // 2. HEADER DETECTION
          if (currentEmp.code) {
            let tempIndices = {};

            row.forEach((cell, idx) => {
              const str = String(cell).trim();
              if (!str) return;

              if (patterns.date.test(str)) tempIndices.date = idx;
              else if (patterns.inTime.test(str) && !patterns.shift.test(str)) tempIndices.inTime = idx;
              else if (patterns.outTime.test(str) && !patterns.shift.test(str)) tempIndices.outTime = idx;
              else if (patterns.status.test(str)) tempIndices.status = idx;
              else if (patterns.punch.test(str)) tempIndices.punch = idx;
            });

            if (tempIndices.date !== undefined &&
              tempIndices.inTime !== undefined &&
              tempIndices.status !== undefined) {
              // Relaxed check: OutTime might be missing from header set but usually strict format has it
              if (tempIndices.outTime !== undefined) {
                headerIndices = tempIndices;
                continue;
              }
            }
          }

          // 3. DATA PROCESSING
          if (currentEmp.code && headerIndices) {
            const dateVal = row[headerIndices.date];
            if (!dateVal || patterns.date.test(String(dateVal))) continue;

            const inTimeVal = row[headerIndices.inTime];
            let hasInTime = inTimeVal && !patterns.inTime.test(String(inTimeVal));

            const statusVal = row[headerIndices.status] || "";
            const normalizeStatus = String(statusVal).trim().toUpperCase();

            const isWeekOff = /WEEK\s*OFF|WO|OFF/i.test(normalizeStatus);
            const isHoliday = /HOLIDAY|PH/i.test(normalizeStatus);
            const isAbsent = normalizeStatus.includes("ABSENT") || normalizeStatus === 'A';

            // Rule: Include if InTime present OR Status contains Absent/Holiday/WeekOff
            if (!hasInTime && !isAbsent && !isWeekOff && !isHoliday) {
              continue;
            }

            const outTimeVal = row[headerIndices.outTime];

            let totalHours = "";
            let outTimeDisplay = "";

            if (outTimeVal && !patterns.outTime.test(String(outTimeVal))) {
              outTimeDisplay = formatTimeOutput(outTimeVal);
              const inHrsDecimal = getDecimalHours(inTimeVal);
              const outHrsDecimal = getDecimalHours(outTimeVal);

              if (inHrsDecimal !== null && outHrsDecimal !== null) {
                let diff = outHrsDecimal - inHrsDecimal;
                if (diff < 0) diff += 24;
                totalHours = parseFloat(diff.toFixed(2));
              }
            }

            const finalName = currentEmp.name || "UNKNOWN";
            const punchStr = headerIndices.punch !== undefined ? row[headerIndices.punch] : "";

            if (!employeeGroups.has(currentEmp.code)) {
              employeeGroups.set(currentEmp.code, {
                code: currentEmp.code,
                name: finalName,
                department: currentEmp.department || "",
                records: []
              });
            }

            employeeGroups.get(currentEmp.code).records.push({
              empCode: currentEmp.code,
              empName: finalName,
              date: formatOutputDate(dateVal),
              inTime: hasInTime ? formatTimeOutput(inTimeVal) : "",
              outTime: outTimeDisplay,
              status: statusVal,
              punchRecords: punchStr,
              totalWorkingHours: totalHours
            });
          }
        }
      }

      if (employeeGroups.size === 0) {
        throw new Error("No valid records found.");
      }

      const groupsArray = Array.from(employeeGroups.values());

      let totalRecs = 0;
      let totalHrsSum = 0;

      groupsArray.forEach(g => {
        totalRecs += g.records.length;

        // Calculate Days from Rows
        let absentCount = 0;
        let presentCount = 0;
        let holidayCount = 0;
        let weekOffCount = 0;
        let empTotalHours = 0;

        g.records.forEach(r => {
          if (typeof r.totalWorkingHours === 'number') {
            totalHrsSum += r.totalWorkingHours;
            empTotalHours += r.totalWorkingHours;
          }

          if (r.status) {
            const s = String(r.status).toUpperCase();
            // Absent Logic
            if (s.includes("ABSENT") || s === 'A') {
              absentCount++;
            }
            // Present Logic
            if (s.includes("PRESENT") || s === 'P' || (typeof r.totalWorkingHours === 'number' && r.totalWorkingHours > 0)) {
              presentCount++;
            }
            // Week Off Logic
            if (/WEEK\s*OFF|WO|OFF/i.test(s)) {
              weekOffCount++;
            }
            // Holiday Logic
            if (/HOLIDAY|PH/i.test(s)) {
              holidayCount++;
            }
          } else if (typeof r.totalWorkingHours === 'number' && r.totalWorkingHours > 0) {
            // implicit present if hours exist but status blank
            presentCount++;
          }
        });
        g.absentDays = absentCount;
        g.presentDays = presentCount;
        g.holidayDays = holidayCount;
        g.weekOffDays = weekOffCount;
        g.totalWorkingHours = empTotalHours;
      });

      setProcessedData(groupsArray);

      setStats({
        totalEmployees: groupsArray.length,
        totalHours: totalHrsSum.toFixed(2)
      });

    } catch (err) {
      console.error(err);
      setError(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!processedData) return;

    // Create Workbook
    const wb = XLSX.utils.book_new();

    // Headers
    const headers = ["Emp Code", "Emp Name", "Att. Date", "InTime", "OutTime", "Status", "Punch Records", "Total Working Hours", "AbsentDays", "PresentDays"];

    // Build rows with Styles
    const wsData = [];

    // Helper to create styled cell
    const createCell = (val, isHeader = false, isWarning = false, isBold = false) => {
      const cell = { v: val, t: 's', s: { font: { name: "Calibri", sz: 11 } } };
      if (typeof val === 'number') cell.t = 'n';

      if (isHeader) {
        cell.s.font.bold = true;
        cell.s.fill = { fgColor: { rgb: "E0E0E0" } };
        cell.s.alignment = { horizontal: "center" };
      }

      if (isBold) {
        cell.s.font.bold = true;
      }

      if (isWarning) {
        cell.s.fill = { fgColor: { rgb: "FFFF00" } }; // Yellow background
      }

      return cell;
    };

    // Add Header Row
    wsData.push(headers.map(h => createCell(h, true)));

    processedData.forEach((group) => {
      let employeeTotal = 0;
      let hasData = false;

      group.records.forEach(r => {
        // Rule: If InTime exists AND OutTime is empty -> Yellow
        // Important: Do not highlight if InTime itself is empty (like Absent rows)
        const isMissingOut = (r.inTime && r.inTime !== "") && (!r.outTime || r.outTime === "");

        const rowCells = [
          createCell(r.empCode, false, isMissingOut),
          createCell(r.empName, false, isMissingOut),
          createCell(r.date, false, isMissingOut),
          createCell(r.inTime, false, isMissingOut),
          createCell(r.outTime, false, isMissingOut),
          createCell(r.status, false, isMissingOut),
          createCell(r.punchRecords, false, isMissingOut),
          createCell(r.totalWorkingHours, false, isMissingOut),
          createCell("", false, isMissingOut), // AbsentDays blank for detail rows
          createCell("", false, isMissingOut)  // PresentDays blank for detail rows
        ];

        wsData.push(rowCells);

        if (typeof r.totalWorkingHours === 'number') {
          employeeTotal += r.totalWorkingHours;
        }
        hasData = true;
      });

      // Summary Row
      if (hasData) {
        const summaryRow = [
          createCell(group.code),
          createCell(group.name),
          createCell(""),
          createCell(""),
          createCell(""),
          createCell(""),
          createCell("Total Hours", false, false, true),
          createCell(parseFloat(employeeTotal.toFixed(2)), false, false, true),
          createCell(group.absentDays, false, false, true), // ABSENT DAYS
          createCell(group.presentDays, false, false, true) // PRESENT DAYS
        ];
        wsData.push(summaryRow);

        // 4 Blank Rows
        for (let i = 0; i < 4; i++) wsData.push([]);
      }
    });

    const ws = XLSX.utils.aoa_to_sheet(wsData);

    // Set Column Widths
    ws['!cols'] = [
      { wch: 10 },
      { wch: 25 },
      { wch: 15 },
      { wch: 10 },
      { wch: 10 },
      { wch: 20 },
      { wch: 30 },
      { wch: 15 },
      { wch: 10 },
      { wch: 10 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Attendance_Output");
    XLSX.writeFile(wb, "Final_Attendance_Report.xlsx");
  };

  const handleSummaryDownload = () => {
    if (!processedData) return;

    // Sort by Emp Code Ascending
    const sortedData = [...processedData].sort((a, b) => {
      return String(a.code || "").localeCompare(String(b.code || ""), undefined, { numeric: true });
    });

    // Create Workbook
    const wb = XLSX.utils.book_new();

    // Headers: Sl, Emp Code, Name, P, A, WO, Total Hr
    const headers = ["Sl", "Emp Code", "Name", "P", "A", "WO", "Total Hr"];

    // Header Style
    const headerStyle = {
      font: { name: "Calibri", sz: 11, bold: true },
      alignment: { horizontal: "center", vertical: "center" },
      fill: { fgColor: { rgb: "E0E0E0" } },
      border: {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      }
    };

    // Data Style
    const dataStyle = {
      font: { name: "Calibri", sz: 11 },
      alignment: { horizontal: "center", vertical: "center" },
      border: {
        top: { style: "thin" },
        bottom: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" }
      }
    };

    const leftAlignStyle = {
      ...dataStyle,
      alignment: { ...dataStyle.alignment, horizontal: "left" }
    };

    const wsData = [];

    // Add Header
    wsData.push(headers.map(h => ({ v: h, t: 's', s: headerStyle })));

    // Add Rows
    sortedData.forEach((emp, index) => {
      wsData.push([
        { v: index + 1, t: 'n', s: dataStyle },
        { v: emp.code, t: 's', s: dataStyle },
        { v: emp.name, t: 's', s: leftAlignStyle },
        { v: emp.presentDays, t: 'n', s: dataStyle },
        { v: emp.absentDays, t: 'n', s: dataStyle },
        { v: emp.weekOffDays, t: 'n', s: dataStyle },
        { v: Number(emp.totalWorkingHours.toFixed(2)), t: 'n', s: dataStyle }
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet([]);

    // Add data to sheet manually to preserve styles
    XLSX.utils.sheet_add_aoa(ws, [], { origin: "A1" }); // Initialize

    // Populating grid
    ws['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: headers.length - 1, r: wsData.length - 1 } });

    for (let R = 0; R < wsData.length; ++R) {
      for (let C = 0; C < headers.length; ++C) {
        const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
        ws[cellRef] = wsData[R][C];
      }
    }

    // Column widths
    ws['!cols'] = [
      { wch: 5 },  // Sl
      { wch: 15 }, // Emp Code
      { wch: 30 }, // Name
      { wch: 5 },  // P
      { wch: 5 },  // A
      { wch: 5 },  // WO
      { wch: 12 }  // Total Hr
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Department_Report");
    XLSX.writeFile(wb, "Department_Wise_Attendance_Report.xlsx");
  };

  return (
    <div className="container animate-fade-in">
      <div className="app-header">
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem', marginBottom: '0.5rem' }}>
          <Clock size={40} color="var(--primary)" />
          <h1 className="app-title">Attendance Pro</h1>
        </div>
        <p className="app-subtitle">
          Professional Attendance Parsing & Reporting Tool.
          <br />
          Supports multi-sheet processing with strict validation.
        </p>
      </div>

      <div className="card">
        <div className="upload-area"
          onClick={() => document.getElementById('fileInput').click()}
          onDragOver={(e) => e.preventDefault()}
          onDrop={(e) => {
            e.preventDefault();
            if (e.dataTransfer.files[0]) {
              setFile(e.dataTransfer.files[0]);
              setError('');
              setProcessedData(null);
            }
          }}
        >
          <FileUp size={64} className="upload-icon" />
          <h3 className="upload-text">{file ? file.name : "Upload Attendance File"}</h3>
          <p className="upload-subtext">
            {file ? "Click to change file" : "Drag & drop or click to browse (.xlsx, .xls)"}
          </p>
          <input
            id="fileInput"
            type="file"
            accept=".xlsx, .xls"
            style={{ display: 'none' }}
            onChange={handleFileUpload}
          />
        </div>

        <div style={{ marginTop: '2rem', display: 'flex', justifyContent: 'center' }}>
          <button
            className="btn btn-primary"
            onClick={processFile}
            disabled={!file || isProcessing}
            style={{ minWidth: '240px', fontSize: '1.1rem' }}
          >
            {isProcessing ? (
              "Processing..."
            ) : (
              <>
                <FileSpreadsheet size={22} /> Process & Generate
              </>
            )}
          </button>
        </div>
      </div>

      {error && (
        <div className="error-msg animate-fade-in">
          <AlertCircle size={24} />
          <span style={{ fontWeight: 500 }}>{error}</span>
        </div>
      )}



      {processedData && (
        <div className="animate-fade-in">
          <div className="card">
            <div className="card-header">
              <div style={{ display: 'flex', alignItems: 'center', gap: '0.75rem' }}>
                <Users size={24} color="var(--primary)" />
                <h2 className="card-title">Data Preview</h2>
              </div>

              <div className="btn-actions">
                <button className="btn btn-primary" onClick={handleDownload} style={{ marginRight: '0.5rem' }}>
                  <Download size={18} /> Detailed Report
                </button>
                <button className="btn btn-success" onClick={handleSummaryDownload}>
                  <FileSpreadsheet size={18} /> Summary Excel
                </button>
              </div>
            </div>

            <div style={{ marginBottom: '1.5rem', display: 'flex', alignItems: 'center', gap: '0.5rem', background: '#eff6ff', padding: '1rem', borderRadius: '8px', color: '#1e40af' }}>
              <Info size={18} />
              <span className="text-sm font-medium">
                Previewing first 5 rows (Yellow background indicates missing OutTime)
              </span>
            </div>

            <div className="table-container">
              <table>
                <thead>
                  <tr>
                    <th>Emp Name</th>
                    <th>Date</th>
                    <th>In Time</th>
                    <th>Out Time</th>
                    <th>Status</th>
                    <th>Total Hours</th>
                  </tr>
                </thead>
                <tbody>
                  {processedData.length > 0 && processedData[0].records.length > 0 ? (
                    processedData[0].records.slice(0, 5).map((r, idx) => (
                      <tr key={idx} style={{ backgroundColor: (r.inTime && !r.outTime) ? '#fef3c7' : undefined }}>
                        <td className="font-medium">{r.empName}</td>
                        <td>{r.date}</td>
                        <td>{r.inTime || "-"}</td>
                        <td style={{ color: (r.inTime && !r.outTime) ? '#b45309' : 'inherit', fontWeight: (r.inTime && !r.outTime) ? 600 : 400 }}>
                          {r.outTime || (r.inTime ? "MISSING" : "-")}
                        </td>
                        <td>
                          <span style={{
                            padding: '0.25rem 0.75rem',
                            borderRadius: '999px',
                            fontSize: '0.75rem',
                            fontWeight: 600,
                            backgroundColor: r.status.includes('PRESENT') ? '#dcfce7' : '#fee2e2',
                            color: r.status.includes('PRESENT') ? '#166534' : '#991b1b'
                          }}>
                            {r.status}
                          </span>
                        </td>
                        <td>{r.totalWorkingHours ? `${r.totalWorkingHours} hrs` : "-"}</td>
                      </tr>
                    ))
                  ) : (
                    <tr><td colSpan="6" style={{ textAlign: 'center', padding: '2rem', color: 'var(--text-secondary)' }}>No records found.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      <footer style={{ textAlign: 'center', marginTop: '4rem', color: '#94a3b8', fontSize: '0.875rem', opacity: 0.8 }}>
        <p>© {new Date().getFullYear()} Attendance Pro System</p>
      </footer>
    </div>
  );
}

export default App;
