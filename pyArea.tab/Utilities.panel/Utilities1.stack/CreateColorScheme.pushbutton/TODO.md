# TODO - CreateColorScheme Script

## High Priority Issues

### 1. Fix False Positive Modified Entries
**Status:** OPEN  
**Priority:** HIGH  
**Description:**  
When updating an existing color scheme, some entries are showing as "modified" even though they have the same color and pattern as the CSV data.

**Current Behavior:**
- Entries with identical RGB values and patterns appear in the "Modified" section
- The old and new colors shown are the same
- Debug logging added but issue persists

**Investigation Notes:**
- Added debug logging on lines 598-603 to track color_changed, caption_changed, pattern_changed flags
- Pattern change detection added on line 596
- Need to check debug output to identify which flag is incorrectly triggering

**Possible Causes:**
1. Caption comparison issue (empty string vs None)
2. Pattern ID comparison issue
3. Color tuple comparison issue (data type mismatch?)
4. CSV data encoding/whitespace issues

**Next Steps:**
- Review debug logs from pyRevit output window
- Check if caption_changed logic needs refinement (line 595)
- Verify color tuple comparison is working correctly
- Consider adding print statements directly to console for easier debugging

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
