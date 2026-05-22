// ============================================================
//  CategoryService.gs
//  Manages the list of graduation categories & their colors
// ============================================================

const CategoryService = (() => {

  const DEFAULT_CATEGORIES = [
    { id: 'honor',    label: 'Honor Grad',         color: '#F0C040', icon: 'star',         sortPriority: 1 },
    { id: 'cte',      label: 'CTE',                color: '#4A90D9', icon: 'tool',         sortPriority: 2 },
    { id: 'avid',     label: 'AVID',               color: '#7B68EE', icon: 'books',        sortPriority: 3 },
    { id: 'asb',      label: 'ASB',                color: '#50C878', icon: 'flag',         sortPriority: 4 },
    { id: 'military', label: 'Military',           color: '#8B6914', icon: 'shield',       sortPriority: 5 },
    { id: 'special',  label: 'Special Recognition',color: '#E8735A', icon: 'award',        sortPriority: 6 }
  ];

  function _sheet() { return getSheet(SHEET_CATEGORIES); }

  function getAllCategories() {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    if (data.length <= 1) return DEFAULT_CATEGORIES;

    return data.slice(1).filter(r => r[0]).map(r => ({
      id:           r[0],
      label:        r[1],
      color:        r[2],
      icon:         r[3],
      sortPriority: Number(r[4]) || 99
    }));
  }

  function updateCategory(id, updates) {
    if (!isAdmin()) return { success: false, error: 'Admin only' };
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        if (updates.label)        sheet.getRange(i+1,2).setValue(updates.label);
        if (updates.color)        sheet.getRange(i+1,3).setValue(updates.color);
        if (updates.icon)         sheet.getRange(i+1,4).setValue(updates.icon);
        if (updates.sortPriority) sheet.getRange(i+1,5).setValue(updates.sortPriority);
        return { success: true };
      }
    }
    return { success: false, error: 'Category not found' };
  }

  return { getAllCategories, updateCategory };
})();
