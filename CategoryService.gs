// ============================================================
//  CategoryService.gs
//  Source of Truth for Category Metadata & Admin Management
// ============================================================

const CategoryService = (() => {
  function _sheet() { return getSheet('Categories'); }

  /**
   * Reads all categories from the spreadsheet.
   * A:ID, B:Label, C:Color, D:Icon, E:Sort, F:Group, G:ShowNav
   */
  function getAllCategories() {
    const data = _sheet().getDataRange().getValues();
    return data.slice(1).filter(r => r[0]).map(r => ({
      id:            String(r[0]).trim().toLowerCase(), 
      label:         r[1],
      color:         r[2],
      icon:          r[3],
      sortPriority:  Number(r[4]) || 99,
      group:         String(r[5] || 'Cord').trim(),
      showInSidebar: r[6] === true || String(r[6]).toUpperCase() === 'TRUE'
    }));
  }

  /**
   * Adds a new category row to the spreadsheet. [cite: 830]
   * Payload: {id, label, color, icon, sortPriority, group}
   */
  function addCategory(payload) {
    const sheet = _sheet();
    // We add the two new columns (Group and ShowInNav) to the appendRow logic
    sheet.appendRow([
      payload.id.toLowerCase().trim(),
      payload.label,
      payload.color,
      payload.icon,
      payload.sortPriority || 99,
      payload.group || 'Cord', // Defaults to Cord group if not specified [cite: 1023]
      false                    // Defaults to NOT showing in sidebar initially [cite: 1024]
    ]);
    return { success: true };
  }

  /**
   * Deletes a category row based on its ID. [cite: 831]
   */
  function deleteCategory(categoryId) {
    const sheet = _sheet();
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === categoryId.toLowerCase()) + 1;
    
    if (rowIndex > 1) { 
      sheet.deleteRow(rowIndex);
      return { success: true };
    }
    return { success: false, error: 'Category not found' };
  }

  /**
   * Toggles whether a category appears in the left-hand sidebar. [cite: 1025]
   */
  function toggleSidebar(categoryId, showStatus) {
    const sheet = _sheet();
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => String(row[0]).trim() === categoryId) + 1;
    
    if (rowIndex > 1) {
      sheet.getRange(rowIndex, 7).setValue(showStatus); // Column G [cite: 1024]
      return { success: true };
    }
    return { success: false, error: 'Category not found' };
  }

  return { getAllCategories, addCategory, deleteCategory, toggleSidebar }; 
})();
