// ============================================================
//  StudentService.gs
//  CRUD operations for student records
// ============================================================

const StudentService = (() => {
  function _sheet() { return getSheet('Students'); }

  // Helper to find column index dynamically based on header name in row 1
  function getColIndex(headerName) {
    const sheet = _sheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    return headers.indexOf(headerName);
  }

  function getAllStudents() {
    const sheet = _sheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const headers = data[0];
    const seatRowIdx = headers.indexOf('Seat Row');
    const seatNumIdx = headers.indexOf('Seat Number');

    return data.slice(1).map(row => {
      let obj = { 
        studentId: String(row[0]), 
        lastName:  row[1], 
        firstName: row[2],
        // Explicitly map these so the Seating Chart can find them
        seatRow: seatRowIdx >= 0 ? String(row[seatRowIdx]) : '',
        seatNumber: seatNumIdx >= 0 ? String(row[seatNumIdx]) : ''
      };
      
      headers.forEach((h, i) => {
        if(i > 2) {
          // Do NOT convert text columns into booleans
          if (['Seat Row', 'Seat Number', 'Not Walking Reason', 'Notes', 'Added By', 'Timestamp'].includes(h)) {
             // Convert Dates to strings so Apps Script doesn't silently return null payloads
             let val = row[i];
             if (val instanceof Date) { val = val.toLocaleString(); }
             obj[h] = val;
          } else {
             // It's a category cord/program, so safely convert to boolean
             const val = row[i];
             obj[h] = (val === true || String(val).toUpperCase() === 'TRUE' || val === 1);
          }
        }
      });
      return obj;
    });
  }

  function bulkAssignCategory(studentIds, categoryId) {
    const user = getCurrentUser();
    const sheet = _sheet();
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the column index for the provided category ID dynamically
    const colIndex = headers.indexOf(categoryId);
    
    if (colIndex === -1) {
      return { 
        success: false, 
        error: `Category '${categoryId}' not found as a column header in the Students sheet.` 
      };
    }

    let updatedCount = 0;
    const idsToFind = [...new Set(studentIds.map(id => String(id).trim()).filter(id => id))];

    for (let i = 1; i < data.length; i++) {
       const sid = String(data[i][0]).trim();
       if (idsToFind.includes(sid)) {
           data[i][colIndex] = true;
           updatedCount++;
       }
    }

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return { success: true, updated: updatedCount };
  }

  function getStats() {
    let students = getAllStudents();

    // --- NEW GATEKEEPER FILTER ---
    // Remove Staff, Aides, and JH from the math so graduation counts are 100% accurate
    students = students.filter(student => {
        const isStaff = String(student.studentId).startsWith('STAFF-');
        const note = student.Notes ? String(student.Notes).trim().toUpperCase() : '';
        return !isStaff && note !== 'JH' && note !== 'A';
    });
    // -----------------------------

    const sheet = _sheet();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 1. Calculate base status stats
    // We look for the exact header names as they appear in your sheet
    let stats = { 
      total: students.length,
      notWalking: students.filter(s => s['notwalking'] === true).length,
      nonGrad:    students.filter(s => s['nongrad'] === true).length,
      notParticipating: students.filter(s => s['notwalking'] === true || s['nongrad'] === true).length,
      // Walking is Total minus those who are not walking or non-grad
      walking:    students.filter(s => s['notwalking'] !== true && s['nongrad'] !== true).length
    };
    
    // 2. Dynamically calculate stats for all other category columns
    headers.slice(3).forEach(h => {
      // Skip the ones we handled manually above
      if (h !== 'notwalking' && h !== 'nongrad' && h !== 'Student ID' && h !== 'Last Name' && h !== 'First Name') {
        stats[h] = students.filter(s => s[h] === true).length;
      }
    });
    
    return stats;
  }

  function addStudent(payload) { /* ... implement as needed ... */ }
  function updateStudent(studentId, payload) { /* ... implement as needed ... */ }
  function deleteStudent(studentId) { /* ... implement as needed ... */ }

  // EXPOSED PUBLIC API
  return { 
    getAllStudents, 
    bulkAssignCategory, 
    addStudent, 
    updateStudent, 
    deleteStudent, 
    getStats 
  };
})();
