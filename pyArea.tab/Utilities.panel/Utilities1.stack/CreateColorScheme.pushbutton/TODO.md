# TODO - CreateColorScheme Script

## High Priority Issues

### 1. Fix False Positive Modified Entries
**Status:** ✅ RESOLVED  
**Priority:** HIGH  
**Description:**  
When updating an existing color scheme, some entries were showing as "modified" even though they had the same RGB color as the CSV data.

**Root Cause:**
- Entries were being added to the "Modified" list whenever ANY change occurred (color, caption, or pattern)
- Caption and pattern changes were causing entries to appear as "modified" even when RGB values were identical

**Solution Implemented:**
- Modified line 607-608 to only add entries to the `modified` list when RGB color actually changes
- Caption and pattern updates now happen silently without appearing in the Modified report
- Only true color changes are now reported in the "Modified" section

**Date Resolved:** Oct 31, 2025

---

### 2. Auto-Select Active Area Scheme
**Status:** OPEN  
**Priority:** MEDIUM  
**Description:**  
Set the initial value in the area scheme dropdown to match the area scheme of the currently active area plan view (if one is open).

**Current Behavior:**
- Area scheme dropdown defaults to index 0 (first scheme alphabetically)
- User must manually select the correct area scheme even if they're already in an area plan view

**Desired Behavior:**
- If active view is an area plan, detect its area scheme
- Set the dropdown to that area scheme automatically
- Fall back to index 0 if no area plan is active

**Implementation Location:**
- Modify `setup_area_schemes()` method (around line 34-48)
- After populating dropdown, check active view type
- If it's an area plan view, get its GenLevel.AreaScheme or similar property
- Set SelectedIndex to match that scheme

**Code Hints:**
```python
# In setup_area_schemes():
active_view = revit.doc.ActiveView
if hasattr(active_view, 'AreaScheme'):
    active_scheme_id = active_view.AreaScheme.Id
    for i, scheme in enumerate(self._area_schemes):
        if scheme.Id == active_scheme_id:
            self.area_scheme_cb.SelectedIndex = i
            break
```

---

## Additional Notes
- Script location: `CreateColorScheme.pushbutton\CreateColorScheme_script.py`
- Debug logging enabled - check pyRevit output window
- Related files: `CreateColorSchemeWindow.xaml`, CSV files in `lib\` directory
