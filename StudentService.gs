// ============================================================
//  StudentService.gs
//  CRUD operations for student records
// ============================================================

const StudentService = (() => {

  // Column indices (0-based) in the Students sheet
  const COL = {
    ID:           0,  // Student ID
    LAST:         1,  // Last Name
    FIRST:        2,  // First Name
    HONOR:        3,  // Honor Grad
    CTE:          4,  // CTE
    AVID:         5,  // AVID
    ASB:          6,  // ASB
    MILITARY:     7,  // Military
    SPECIAL:      8,  // Special Recognition
    NOT_WALKING:  9,  // Not Walking
    NW_REASON:    10, // Not Walking Reason
    SEAT_ROW:     11, // Assigned Row
    SEAT_NUM:     12, // Assigned Seat Number
    NOTES:        13, // Admin Notes
    ADDED_BY:     14, // Email of person who added
    TIMESTAMP:    15  // Last modified
  };

  const HEADERS = [
    'Student ID', 'Last Name', 'First Name',
    'Honor Grad', 'CTE', 'AVID', 'ASB', 'Military', 'Special Recognition',
    'Not Walking', 'Not Walking Reason',
    'Seat Row', 'Seat Number', 'Notes', 'Added By', 'Timestamp'
  ];

  function _sheet() { return getSheet(SHEET_STUDENTS); }

  function _rowToObj(row) {
    return {
      studentId:    row[COL.ID],
      lastName:     row[COL.LAST],
      firstName:    row[COL.FIRST],
      honor:        row[COL.HONOR]     === true || row[COL.HONOR]     === 'TRUE',
      cte:          row[COL.CTE]       === true || row[COL.CTE]       === 'TRUE',
      avid:         row[COL.AVID]      === true || row[COL.AVID]      === 'TRUE',
      asb:          row[COL.ASB]       === true || row[COL.ASB]       === 'TRUE',
      military:     row[COL.MILITARY]  === true || row[COL.MILITARY]  === 'TRUE',
      special:      row[COL.SPECIAL]   === true || row[COL.SPECIAL]   === 'TRUE',
      notWalking:   row[COL.NOT_WALKING] === true || row[COL.NOT_WALKING] === 'TRUE',
      nwReason:     row[COL.NW_REASON],
      seatRow:      row[COL.SEAT_ROW],
      seatNumber:   row[COL.SEAT_NUM],
      notes:        row[COL.NOTES],
      addedBy:      row[COL.ADDED_BY],
      timestamp:    row[COL.TIMESTAMP] ? row[COL.TIMESTAMP].toString() : ''
    };
  }

  function getAllStudents() {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).filter(r => r[COL.ID] !== '').map(_rowToObj);
  }

  function getStudentById(studentId) {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    const sid   = String(studentId).trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][COL.ID]).trim() === sid) {
        return { rowIndex: i + 1, student: _rowToObj(data[i]) };
      }
    }
    return null;
  }

  function addStudent(payload) {
    const user = getCurrentUser();

    // Permission check: admins can add any field;
    // category managers can only mark their category
    if (user.role !== 'admin') {
      const categoryFields = ['honor','cte','avid','asb','military','special'];
      const allowed = user.categories.map(c => c.toLowerCase());
      categoryFields.forEach(f => {
        if (!allowed.includes(f)) payload[f] = false;
      });
    }

    // Check for duplicate
    const existing = getStudentById(payload.studentId);
    if (existing) {
      // Update the existing record instead
      return updateStudent(payload.studentId, payload);
    }

    const sheet = _sheet();
    const now   = new Date();
    const row   = [
      payload.studentId || '',
      payload.lastName  || '',
      payload.firstName || '',
      payload.honor     || false,
      payload.cte       || false,
      payload.avid      || false,
      payload.asb       || false,
      payload.military  || false,
      payload.special   || false,
      payload.notWalking || false,
      payload.nwReason  || '',
      '',   // seatRow - assigned later
      '',   // seatNumber - assigned later
      payload.notes     || '',
      user.email,
      now
    ];
    sheet.appendRow(row);
    return { success: true, action: 'added', studentId: payload.studentId };
  }

  function updateStudent(studentId, payload) {
    const user   = getCurrentUser();
    const result = getStudentById(studentId);
    if (!result) return { success: false, error: 'Student not found' };

    const { rowIndex, student } = result;
    const sheet = _sheet();
    const now   = new Date();

    // Non-admins can only update their assigned category flags
    const merged = { ...student, ...payload };
    if (user.role !== 'admin') {
      const allowed = user.categories.map(c => c.toLowerCase());
      ['honor','cte','avid','asb','military','special'].forEach(f => {
        if (!allowed.includes(f)) merged[f] = student[f]; // restore original
      });
    }

    sheet.getRange(rowIndex, 1, 1, 16).setValues([[
      merged.studentId,
      merged.lastName,
      merged.firstName,
      merged.honor,
      merged.cte,
      merged.avid,
      merged.asb,
      merged.military,
      merged.special,
      merged.notWalking,
      merged.nwReason || '',
      merged.seatRow  || '',
      merged.seatNumber || '',
      merged.notes    || '',
      user.email,
      now
    ]]);

    return { success: true, action: 'updated', studentId };
  }

  function deleteStudent(studentId) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const result = getStudentById(studentId);
    if (!result) return { success: false, error: 'Student not found' };
    _sheet().deleteRow(result.rowIndex);
    return { success: true };
  }

  function getStats() {
    const students = getAllStudents();
    const walking  = students.filter(s => !s.notWalking);
    return {
      total:       students.length,
      walking:     walking.length,
      notWalking:  students.filter(s => s.notWalking).length,
      honor:       students.filter(s => s.honor).length,
      cte:         students.filter(s => s.cte).length,
      avid:        students.filter(s => s.avid).length,
      asb:         students.filter(s => s.asb).length,
      military:    students.filter(s => s.military).length,
      special:     students.filter(s => s.special).length
    };
  }

  function bulkImport(rows) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const results = rows.map(r => addStudent(r));
    return { success: true, count: results.length };
  }

  return { getAllStudents, getStudentById, addStudent, updateStudent,
           deleteStudent, getStats, bulkImport };
})();
