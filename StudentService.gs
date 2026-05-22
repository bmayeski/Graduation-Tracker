// ============================================================
//  StudentService.gs
//  CRUD operations for student records
// ============================================================

const StudentService = (() => {

  // Column indices (0-based) matching your exact Sheet headers
  const COL = {
    ID:              0,  LAST:            1,  FIRST:           2,
    NON_GRAD:        3,  NOT_WALKING:     4,  HONOR:           5,
    CTE:             6,  AVID:            7,  ASB:             8,
    MILITARY:        9,  CSF:             10, HPC:             11,
    VAPA:            12, PA:              13, SSA:             14,
    ART:             15, FNL:             16, YEARBOOK:        17,
    NORSECREW:       18, SHARP_MEM:       19, KEY:             20,
    BI_LITERACY:     21, GOLD_STATE_SEAL: 22, HOSA:            23,
    ACADEMIC_LEAGUE: 24, NW_REASON:       25, SEAT_ROW:        26,
    SEAT_NUM:        27, NOTES:           28, ADDED_BY:        29,
    TIMESTAMP:       30
  };

  const HEADERS = [
    'Student ID', 'Last Name', 'First Name', 'Not Walking',
    'Honor Grad', 'CTE', 'AVID', 'ASB', 'Military', 'CSF', 'HPC', 'VAPA', 
    'PA', 'SSA', 'Art', 'FNL', 'Yearbook', 'Norsecrew', 'Sharp-Mem', 
    'Key Club', 'Bi-Literacy', 'Gold State Seal', 'HOSA', 'Academic League', 
    'Not Walking Reason', 'Seat Row', 'Seat Number', 'Notes', 'Added By', 'Timestamp'
  ];

  // This array is used for permission checks to see if a teacher can edit a field
  const CAT_FIELDS = [
    'honor', 'cte', 'avid', 'asb', 'military', 'csf', 'hpc', 'vapa', 'pa', 
    'ssa', 'art', 'fnl', 'yearbook', 'norsecrew', 'sharp-mem', 'key club', 
    'bi-literacy', 'gold state seal', 'hosa', 'academic league'
  ];

  function _sheet() { return getSheet('Students'); }

  function _rowToObj(row) {
    return {
      studentId:       row[COL.ID],
      lastName:        row[COL.LAST],
      firstName:       row[COL.FIRST],
      notWalking:      row[COL.NOT_WALKING] === true || row[COL.NOT_WALKING] === 'TRUE',
      honor:           row[COL.HONOR]       === true || row[COL.HONOR]       === 'TRUE',
      cte:             row[COL.CTE]         === true || row[COL.CTE]         === 'TRUE',
      avid:            row[COL.AVID]        === true || row[COL.AVID]        === 'TRUE',
      asb:             row[COL.ASB]         === true || row[COL.ASB]         === 'TRUE',
      military:        row[COL.MILITARY]    === true || row[COL.MILITARY]    === 'TRUE',
      csf:             row[COL.CSF]         === true || row[COL.CSF]         === 'TRUE',
      hpc:             row[COL.HPC]         === true || row[COL.HPC]         === 'TRUE',
      vapa:            row[COL.VAPA]        === true || row[COL.VAPA]        === 'TRUE',
      pa:              row[COL.PA]          === true || row[COL.PA]          === 'TRUE',
      ssa:             row[COL.SSA]         === true || row[COL.SSA]         === 'TRUE',
      art:             row[COL.ART]         === true || row[COL.ART]         === 'TRUE',
      fnl:             row[COL.FNL]         === true || row[COL.FNL]         === 'TRUE',
      yearbook:        row[COL.YEARBOOK]    === true || row[COL.YEARBOOK]    === 'TRUE',
      norsecrew:       row[COL.NORSECREW]   === true || row[COL.NORSECREW]   === 'TRUE',
      'sharp-mem':     row[COL.SHARP_MEM]   === true || row[COL.SHARP_MEM]   === 'TRUE',
      'key club':      row[COL.KEY_CLUB]    === true || row[COL.KEY_CLUB]    === 'TRUE',
      'bi-literacy':   row[COL.BI_LITERACY] === true || row[COL.BI_LITERACY] === 'TRUE',
      'gold state seal': row[COL.GOLD_STATE_SEAL] === true || row[COL.GOLD_STATE_SEAL] === 'TRUE',
      hosa:            row[COL.HOSA]        === true || row[COL.HOSA]        === 'TRUE',
      'academic league': row[COL.ACADEMIC_LEAGUE] === true || row[COL.ACADEMIC_LEAGUE] === 'TRUE',
      nwReason:        row[COL.NW_REASON],
      seatRow:         row[COL.SEAT_ROW],
      seatNumber:      row[COL.SEAT_NUM],
      notes:           row[COL.NOTES],
      addedBy:         row[COL.ADDED_BY],
      timestamp:       row[COL.TIMESTAMP] ? row[COL.TIMESTAMP].toString() : ''
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

    // Permission check: admins can add any field; category managers can only mark their category
    if (user.role !== 'admin') {
      const allowed = user.categories.map(c => c.toLowerCase());
      CAT_FIELDS.forEach(f => {
        if (!allowed.includes(f)) payload[f] = false;
      });
    }

    // Check for duplicate
    const existing = getStudentById(payload.studentId);
    if (existing) {
      return updateStudent(payload.studentId, payload);
    }

    const sheet = _sheet();
    const now   = new Date();
    
    // Build the row exactly matching the 30 columns
    const row = [
      payload.studentId || '',
      payload.lastName  || '',
      payload.firstName || '',
      payload.notWalking || false,
      payload.honor     || false,
      payload.cte       || false,
      payload.avid      || false,
      payload.asb       || false,
      payload.military  || false,
      payload.csf       || false,
      payload.hpc       || false,
      payload.vapa      || false,
      payload.pa        || false,
      payload.ssa       || false,
      payload.art       || false,
      payload.fnl       || false,
      payload.yearbook  || false,
      payload.norsecrew || false,
      payload['sharp-mem'] || false,
      payload['key club'] || false,
      payload['bi-literacy'] || false,
      payload['gold state seal'] || false,
      payload.hosa      || false,
      payload['academic league'] || false,
      payload.nwReason  || '',
      payload.seatRow   || '',
      payload.seatNumber|| '',
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
    
    // Merge new payload with existing student data
    const merged = { ...student, ...payload };
    
    // Non-admins can only update their assigned category flags
    if (user.role !== 'admin') {
      const allowed = user.categories.map(c => c.toLowerCase());
      CAT_FIELDS.forEach(f => {
        if (!allowed.includes(f)) merged[f] = student[f]; // restore original
      });
    }

    sheet.getRange(rowIndex, 1, 1, 30).setValues([[
      merged.studentId,
      merged.lastName,
      merged.firstName,
      merged.notWalking,
      merged.honor,
      merged.cte,
      merged.avid,
      merged.asb,
      merged.military,
      merged.csf,
      merged.hpc,
      merged.vapa,
      merged.pa,
      merged.ssa,
      merged.art,
      merged.fnl,
      merged.yearbook,
      merged.norsecrew,
      merged['sharp-mem'],
      merged['key club'],
      merged['bi-literacy'],
      merged['gold state seal'],
      merged.hosa,
      merged['academic league'],
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
      // You can add more stats here if you want them on the dashboard!
    };
  }

  function bulkImport(rows) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const results = rows.map(r => addStudent(r));
    return { success: true, count: results.length };
  }

  function bulkAssignCategory(studentIds, categoryId) {
    const user = getCurrentUser();
    
    // Permission check
    if (user.role !== 'admin') {
       const allowed = user.categories.map(c => c.toLowerCase());
       if (!allowed.includes(categoryId.toLowerCase())) {
           return { success: false, error: 'Permission denied for this category.' };
       }
    }

    // Use the ID exactly as provided from your spreadsheet data
    const formattedCatId = categoryId.toUpperCase(); 
    
    // Ensure your COL constant in StudentService.gs uses these exact keys:
    // KEY: 20 (instead of KEY_CLUB)
    const colIndex = COL[formattedCatId];
    
    if (colIndex === undefined) {
       return { success: false, error: `Category column '${formattedCatId}' not found. Available keys: ${Object.keys(COL).join(', ')}` };
    }

    let updatedCount = 0;
    const idsToFind = [...new Set(studentIds.map(id => String(id).trim()).filter(id => id))];
    const notFound = [...idsToFind];

    for (let i = 1; i < data.length; i++) {
       const sid = String(data[i][COL.ID]).trim();
       const idx = notFound.indexOf(sid);
       if (idx !== -1) {
           data[i][colIndex] = true; 
           data[i][COL.TIMESTAMP] = new Date();
           data[i][COL.ADDED_BY] = user.email;
           updatedCount++;
           notFound.splice(idx, 1);
       }
    }

    if (updatedCount > 0) {
       sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    return { success: true, updated: updatedCount, notFound: notFound };
  }

  return { getAllStudents, getStudentById, addStudent, updateStudent,
           deleteStudent, getStats, bulkImport, bulkAssignCategory };
})();
