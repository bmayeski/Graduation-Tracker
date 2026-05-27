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
      
      // Use alphabet letters for rows (A = 65)
      const letter = String.fromCharCode(65 + r);
      const leftDisplayRow = `${letter} - L`;
      const rightDisplayRow = `${letter} - R`;

      for (let s = 0; s < maxSeats; s++) {
        // Left side (Odds starting from 1)
        if (s < leftCount) {
          // Changed to assign odd numbers sequentially (1, 3, 5...)
          const physicalSeat = (s * 2) + 1; 
          const seatId = `${leftDisplayRow}-${physicalSeat}`;
          if (!occupiedSeats.has(seatId) && pointer < unassigned.length) {
            assignments.push({ studentId: unassigned[pointer++].studentId, rowLabel: leftDisplayRow, seatNumber: physicalSeat });
            occupiedSeats.add(seatId);
          }
        }
        // Right side (Evens starting from 2)
        if (s < rightCount) {
          // Changed to assign even numbers sequentially (2, 4, 6...)
          const physicalSeat = (s * 2) + 2; 
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
    const sheet = getSheet('Students');
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

  function assignSeatManually(rowLabel, seatNumber, inputValue, previousId) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    
    const sheet = getSheet('Students');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idCol = 0; // Student ID
    const lastCol = 1; // Last Name
    const firstCol = 2; // First Name
    const seatRowCol = headers.indexOf('Seat Row');
    const seatNumCol = headers.indexOf('Seat Number');
    const notesCol = headers.indexOf('Notes');
    
    if (seatRowCol === -1 || seatNumCol === -1) return { success: false, error: 'Seat columns not found' };

    // 1. If there was a previous student/staff in this seat, clear them out first
    if (previousId) {
       for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]).trim() === String(previousId).trim()) {
             sheet.getRange(i + 1, seatRowCol + 1).setValue('');
             sheet.getRange(i + 1, seatNumCol + 1).setValue('');
             break;
          }
       }
    }

    // 2. If the input is empty, they just wanted to clear the seat. Return success.
    if (!inputValue) return { success: true };

    // 3. Check if they typed an acronym, or a name/ID with an acronym at the end
    let searchVal = inputValue.trim();
    let acronymToSave = '';
    const validAcronyms = ['V', 'S', 'JH', 'A', 'W'];
    
    // If they JUST typed "A" or "JH" to reserve an empty seat
    if (validAcronyms.includes(searchVal.toUpperCase())) {
       let newRow = new Array(headers.length).fill('');
       newRow[idCol] = 'STAFF-' + Math.floor(Math.random() * 10000); 
       newRow[lastCol] = searchVal.toUpperCase(); 
       newRow[firstCol] = 'Staff'; 
       newRow[seatRowCol] = rowLabel;
       newRow[seatNumCol] = seatNumber;
       sheet.appendRow(newRow);
       return { success: true };
    }
    
    // If they typed multiple words, check if the LAST (or first) word is an acronym (e.g. "Mr Smith A" or "12345 V")
    const parts = searchVal.split(' ');
    if (parts.length > 1) {
       const lastWord = parts[parts.length - 1].toUpperCase();
       const firstWord = parts[0].toUpperCase();
       
       if (validAcronyms.includes(lastWord)) {
          acronymToSave = lastWord;
          parts.pop(); // Remove the acronym from the search string
          searchVal = parts.join(' ').trim();
       } else if (validAcronyms.includes(firstWord)) {
          acronymToSave = firstWord;
          parts.shift(); // Remove the acronym from the search string
          searchVal = parts.join(' ').trim();
       }
    }

    // 4. Search for the ID or Name using the cleaned searchVal
    const isId = /^\d+$/.test(searchVal);
    
    if (isId) {
       // Search for the Student ID
       let found = false;
       for (let i = 1; i < data.length; i++) {
          if (String(data[i][idCol]).trim() === searchVal) {
             sheet.getRange(i + 1, seatRowCol + 1).setValue(rowLabel);
             sheet.getRange(i + 1, seatNumCol + 1).setValue(seatNumber);
             if (acronymToSave && notesCol !== -1) sheet.getRange(i + 1, notesCol + 1).setValue(acronymToSave);
             found = true;
             break;
          }
       }
       if (!found) return { success: false, error: 'Student ID not found in database.' };
    } else {
       // It's a Staff Name
       let found = false;
       for (let i = 1; i < data.length; i++) {
          const fullName = (String(data[i][firstCol]) + " " + String(data[i][lastCol])).toLowerCase();
          if (fullName.includes(searchVal.toLowerCase()) || String(data[i][lastCol]).toLowerCase() === searchVal.toLowerCase()) {
             sheet.getRange(i + 1, seatRowCol + 1).setValue(rowLabel);
             sheet.getRange(i + 1, seatNumCol + 1).setValue(seatNumber);
             if (acronymToSave && notesCol !== -1) sheet.getRange(i + 1, notesCol + 1).setValue(acronymToSave);
             found = true;
             break;
          }
       }
       
       // If no match is found, append them as a new "Staff" row
       if (!found) {
         let newRow = new Array(headers.length).fill('');
         newRow[idCol] = 'STAFF-' + Math.floor(Math.random() * 10000); 
         newRow[lastCol] = searchVal; // The clean name, e.g. "Mr Smith"
         newRow[firstCol] = 'Staff';
         newRow[seatRowCol] = rowLabel;
         newRow[seatNumCol] = seatNumber;
         if (acronymToSave && notesCol !== -1) newRow[notesCol] = acronymToSave;
         sheet.appendRow(newRow);
       }
    }

    return { success: true };
  }

  function assignJuniorHonorsToRowK() {
    if (!isAdmin()) return { success: false, error: 'Admin only' };

    const sheet = getSheet('Students');
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    const idCol = 0;
    const lastCol = 1;
    const firstCol = 2;
    const notesCol = headers.indexOf('Notes');
    const seatRowCol = headers.indexOf('Seat Row');
    const seatNumCol = headers.indexOf('Seat Number');

    if (seatRowCol === -1 || seatNumCol === -1 || notesCol === -1) {
      return { success: false, error: 'Missing necessary columns.' };
    }

    // 1. Gather all unassigned JH students
    const jhStudents = [];
    for (let i = 1; i < data.length; i++) {
      const note = String(data[i][notesCol]).trim().toUpperCase();
      const hasSeat = data[i][seatRowCol] && data[i][seatNumCol];
      
      if (note === 'JH' && !hasSeat) {
        jhStudents.push({
          rowIdx: i + 1,
          last: String(data[i][lastCol]),
          first: String(data[i][firstCol])
        });
      }
    }

    if (jhStudents.length === 0) return { success: false, error: 'No unassigned Junior Honors found.' };

    // 2. Get Row K capacities (Index 10 because A=0, B=1 ... K=10)
    const config = getSeatingConfig();
    const leftSec = config.sections.find(s => s.id === 'L');
    const rightSec = config.sections.find(s => s.id === 'R');

    if (!leftSec || !rightSec || leftSec.rows.length <= 10 || rightSec.rows.length <= 10) {
      return { success: false, error: 'Row K does not exist! Please add at least 11 rows in Layout Config.' };
    }

    const leftCount = leftSec.rows[10];
    const rightCount = rightSec.rows[10];
    const maxSeats = Math.max(leftCount, rightCount);

    // 3. Assign them alternating outside-in specifically for Row K
    const updates = [];
    let pointer = 0;

    for (let s = 0; s < maxSeats; s++) {
      // Left Side (Odds) - Outside in
      if (s < leftCount && pointer < jhStudents.length) {
        const physicalSeat = (s * 2) + 1;
        updates.push({ rowIdx: jhStudents[pointer].rowIdx, rowLabel: 'K - L', seatNum: physicalSeat });
        pointer++;
      }
      // Right Side (Evens) - Outside in
      if (s < rightCount && pointer < jhStudents.length) {
        const physicalSeat = ((rightCount - 1 - s) * 2) + 2;
        updates.push({ rowIdx: jhStudents[pointer].rowIdx, rowLabel: 'K - R', seatNum: physicalSeat });
        pointer++;
      }
    }

    // 4. Save the assignments directly back to the Students sheet
    updates.forEach(u => {
      sheet.getRange(u.rowIdx, seatRowCol + 1).setValue(u.rowLabel);
      sheet.getRange(u.rowIdx, seatNumCol + 1).setValue(u.seatNum);
    });

    return { success: true, count: updates.length, totalJH: jhStudents.length };
  }

  return { 
    getSeatingConfig, saveSeatingConfig, generateSeatingAssignments, 
    getSeatingChart, assignSeatManually, assignJuniorHonorsToRowK // <--- Add it here
  };
})();
