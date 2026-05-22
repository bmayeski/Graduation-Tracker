// ============================================================
//  CategoryService.gs
//  Manages the list of graduation categories & their colors
// ============================================================

const CategoryService = (() => {
  function _sheet() { return getSheet('Categories'); }

  function getAllCategories() {
    const sheet = _sheet();
    const data  = sheet.getDataRange().getValues();
    
    // Skip the header row (slice(1)) and filter out empty IDs
    return data.slice(1).filter(r => r[0]).map(r => ({
      id:           String(r[0]).trim().toLowerCase(), // The ID in your sheet (e.g., 'key')
      label:        r[1],
      color:        r[2],
      icon:         r[3],
      sortPriority: Number(r[4]) || 99
    }));
  }

  return { getAllCategories };
})();
