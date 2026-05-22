// ============================================================
//  SeatingService.gs
//  Seating chart configuration and auto-assignment
// ============================================================

const SeatingService = (() => {

  // Default row config: [ { rowLabel, seats, section } ]
  const DEFAULT_CONFIG = {
    sections: [
      { id: 'A', label: 'Section A (Left)',   rows: 10, seatsPerRow: 20 },
      { id: 'B', label: 'Section B (Center)', rows: 12, seatsPerRow: 24 },
      { id: 'C', label: 'Section C (Right)',  rows: 10, seatsPerRow: 20 }
    ],
    honorFront:    true,  // Place honor grads in front rows
    militaryFront: false,
    sortOrder:     'alpha' // 'alpha' | 'id' | 'category'
  };

  function _sheet() { return getSheet(SHEET_SEATING); }

  function getSeatingConfig() {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();

    if (data.length < 2 || !data[1][0]) return DEFAULT_CONFIG;

    try {
      return JSON.parse(data[1][0]);
    } catch(e) {
      return DEFAULT_CONFIG;
    }
  }

  function saveSeatingConfig(config) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const sheet = _sheet();
    // Store config as JSON in cell A2
    if (sheet.getLastRow() < 2) sheet.appendRow(['config']);
    sheet.getRange(2, 1).setValue(JSON.stringify(config));
    return { success: true };
  }

  /**
   * Auto-assign all walking students to seats.
   * Honor grads → front rows (if honorFront).
   * Otherwise alphabetical by last name within each section.
   */
  function generateSeatingAssignments() {
    if (!isAdmin()) return { success: false, error: 'Admin only' };

    const config   = getSeatingConfig();
    const students = StudentService.getAllStudents().filter(s => !s.notWalking);

    // Build sort order
    let sorted = [...students];
    if (config.sortOrder === 'alpha') {
      sorted.sort((a, b) => (a.lastName + a.firstName).localeCompare(b.lastName + b.firstName));
    } else if (config.sortOrder === 'id') {
      sorted.sort((a, b) => String(a.studentId).localeCompare(String(b.studentId)));
    }

    // Optionally bubble honor grads to front
    if (config.honorFront) {
      const honor    = sorted.filter(s => s.honor);
      const nonHonor = sorted.filter(s => !s.honor);
      sorted = [...honor, ...nonHonor];
    }
    if (config.militaryFront) {
      const mil    = sorted.filter(s => s.military);
      const nonMil = sorted.filter(s => !s.military);
      sorted = [...mil, ...nonMil];
    }

    // Assign seats sequentially across sections/rows
    let pointer = 0;
    const assignments = [];

    config.sections.forEach(section => {
      for (let row = 1; row <= section.rows; row++) {
        for (let seat = 1; seat <= section.seatsPerRow; seat++) {
          if (pointer >= sorted.length) break;
          const student = sorted[pointer++];
          assignments.push({
            studentId:  student.studentId,
            section:    section.id,
            rowLabel:   `${section.id}${row}`,
            seatNumber: seat
          });
        }
      }
    });

    // Write assignments back to Students sheet
    assignments.forEach(a => {
      StudentService.updateStudent(a.studentId, { seatRow: a.rowLabel, seatNumber: a.seatNumber });
    });

    return { success: true, assigned: assignments.length, total: students.length };
  }

  function getSeatingChart() {
    const config   = getSeatingConfig();
    const students = StudentService.getAllStudents();

    // Build a lookup map by seat
    const seatMap = {};
    students.forEach(s => {
      if (s.seatRow && s.seatNumber) {
        seatMap[`${s.seatRow}-${s.seatNumber}`] = s;
      }
    });

    return { config, seatMap, students };
  }

  return { getSeatingConfig, saveSeatingConfig, generateSeatingAssignments, getSeatingChart };
})();
