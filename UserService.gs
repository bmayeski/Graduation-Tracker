// ============================================================
//  UserService.gs
//  User management and role/permission assignment
// ============================================================

const UserService = (() => {

  // Roles: 'admin' | 'category_manager' | 'viewer'
  const HEADERS = ['Email', 'Role', 'Categories', 'Display Name', 'Added Date'];

  function _sheet() { return getSheet(SHEET_USERS); }

  function getAllUsers() {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return { success: true, users: [] };

    return {
      success: true,
      users: data.slice(1).filter(r => r[0]).map(r => ({
        email:       r[0],
        role:        r[1],
        categories:  r[2] ? r[2].split(',') : [],
        displayName: r[3],
        addedDate:   r[4] ? r[4].toString() : ''
      }))
    };
  }

  function addUser(email, role, categories, displayName) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };

    const sheet    = _sheet();
    const data     = sheet.getDataRange().getValues();
    const existing = data.slice(1).findIndex(r => r[0] === email);

    const row = [email, role, Array.isArray(categories) ? categories.join(',') : '', displayName || '', new Date()];

    if (existing >= 0) {
      sheet.getRange(existing + 2, 1, 1, 5).setValues([row]);
    } else {
      sheet.appendRow(row);
    }

    return { success: true };
  }

  function removeUser(email) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === email) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, error: 'User not found' };
  }

  return { getAllUsers, addUser, removeUser };
})();
