// ============================================================
//  SeatingService.gs
//  Seating chart configuration and auto-assignment
// ============================================================

const SeatingService = (() => {

  const DEFAULT_CONFIG = {
    sections: [
      { id: 'L', label: 'Left Side (Odds)',  rows: [18, 19, 20, 20, 20, 20, 20, 20, 20, 21, 15] },
      { id: 'R', label: 'Right Side (Evens)', rows: [18, 19, 20, 20, 20, 20, 20, 20, 20, 21, 15] }
    ],
    honorFront: true,
    militaryFront: false,
    sortOrder: 'alpha' 
  };

  function _sheet() { return getSheet('SeatingConfig'); }

  function getSeatingConfig() {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length < 2 || !data[1][0]) return DEFAULT_CONFIG;
    try { return JSON.parse(data[1][0]); } catch(e) { return DEFAULT_CONFIG; }
  }

  function saveSeatingConfig(config) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const sheet = _sheet();
    if (sheet.getLastRow() < 2) sheet.appendRow(['config']);
    sheet.getRange(2, 1).setValue(JSON.stringify(config));
    return { success: true };
  }

  function generateSeatingAssignments(overwriteAll = false) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const config = getSeatingConfig();
    const allStudents = StudentService.getAllStudents().filter(s => !s.notWalking);

    let unassigned = [];
    const occupiedSeats = new Set();

    // 1. Separate manual assignments from unassigned
    allStudents.forEach(st => {
      if (!overwriteAll && st.seatRow && st.seatNumber) {
        occupiedSeats.add(`${st.seatRow}-${st.seatNumber}`);
      } else {
        unassigned.push(st);
      }
    });

    // 2. Sort unassigned
    if (config.sortOrder === 'alpha') {
      unassigned.sort((a, b) => (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName));
    } else if (config.sortOrder === 'id') {
      unassigned.sort((a, b) => String(a.studentId).localeCompare(String(b.studentId)));
    }
    if (config.honorFront) {
      unassigned = [...unassigned.filter(st => st.honor), ...unassigned.filter(st => !st.honor)];
    }

    // 3. Assign Outside-In with Continuous Odd/Even Numbering
    let pointer = 0;
    const assignments = [];
    
    const leftSec = config.sections.find(sec => sec.id === 'L');
    const rightSec = config.sections.find(sec => sec.id === 'R');
    if(!leftSec || !rightSec) return { success: false, error: 'Missing L and R config.' };

    const maxRows = Math.max(leftSec.rows.length, rightSec.rows.length);
    let currentOffset = 0;

    for (let r = 0; r < maxRows; r++) {
      const leftCount = leftSec.rows[r] || 0;
      const rightCount = rightSec.rows[r] || 0;
      const maxSeats = Math.max(leftCount, rightCount);
      const rowOffset = currentOffset; 
      
      // Calculate the specific label for this row (JH if it's the last row)
      const isLastLeftRow = (r === leftSec.rows.length - 1);
      const leftDisplayRow = isLastLeftRow ? 'LJH' : `L${r+1}`;
      
      const isLastRightRow = (r === rightSec.rows.length - 1);
      const rightDisplayRow = isLastRightRow ? 'RJH' : `R${r+1}`;

      for (let s = 0; s < maxSeats; s++) {
        // Left side (Odds)
        if (s < leftCount) {
          const physicalSeat = rowOffset + ((leftCount - 1 - s) * 2) + 1;
          const seatId = `${leftDisplayRow}-${physicalSeat}`;
          if (!occupiedSeats.has(seatId) && pointer < unassigned.length) {
            assignments.push({ studentId: unassigned[pointer++].studentId, rowLabel: leftDisplayRow, seatNumber: physicalSeat });
            occupiedSeats.add(seatId);
          }
        }
        // Right side (Evens)
        if (s < rightCount) {
          const physicalSeat = rowOffset + (rightCount - s) * 2;
          const seatId = `${rightDisplayRow}-${physicalSeat}`;
          if (!occupiedSeats.has(seatId) && pointer < unassigned.length) {
            assignments.push({ studentId: unassigned[pointer++].studentId, rowLabel: rightDisplayRow, seatNumber: physicalSeat });
            occupiedSeats.add(seatId);
          }
        }
      }
      currentOffset += (leftCount + rightCount); // Update cumulative offset
    }

    // 4. Bulk update the Students sheet for performance
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Students');
    const data = sheet.getDataRange().getValues();
    const assignMap = {};
    assignments.forEach(a => assignMap[String(a.studentId)] = a);

    data.forEach((row, index) => {
       if(index === 0) return; // Skip headers
       const sid = String(row[0]).trim();
       if (!sid) return;

       const a = assignMap[sid];
       if (a) {
          row[11] = a.rowLabel;
          row[12] = a.seatNumber;
       } else if (overwriteAll && row[11]) {
          // Clear seat if they were unassigned and we are overwriting
          row[11] = '';
          row[12] = '';
       }
    });

    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return { success: true, assigned: assignments.length, total: allStudents.length };
  }

  function getSeatingChart() {
    const config   = getSeatingConfig();
    const students = StudentService.getAllStudents();
    const seatMap = {};
    students.forEach(s => {
      if (s.seatRow && s.seatNumber) seatMap[`${s.seatRow}-${s.seatNumber}`] = s;
    });
    return { config, seatMap, students };
  }

  return { getSeatingConfig, saveSeatingConfig, generateSeatingAssignments, getSeatingChart };
})();
