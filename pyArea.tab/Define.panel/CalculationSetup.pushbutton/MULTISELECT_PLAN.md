# Multi-Selection Refactoring Plan
**CalculationSetup Script Enhancement**

**Version:** 2.0 | **Date:** November 12, 2025 | **Status:** Planning / Review (Updated with GPT improvements)

---

## 1. Executive Summary

### Objective
Add checkbox-based multi-selection with bulk editing and operations.

### Key Features
- ✅ Multi-select with checkboxes (explicit, visible)
- ✅ Bulk field editing with mixed-value detection and "No change" default
- ✅ Explicit Apply/Cancel in bulk mode (replaces auto-save)
- ✅ Bulk change preview before apply
- ✅ Bulk remove operations
- ✅ Bulk track AreaPlans on multiple Sheets (non-placement)
- ✅ Bulk set representing view (multiple children → single parent)
- ✅ State preservation across tree rebuilds
- ✅ Session-scoped undo for last bulk change

### Constraints
- **Same ElementType rule:** All checked nodes must be same type
- **Same AreaScheme rule:** All checked nodes must be in same AreaScheme
- **Validation enforced:** Auto-revert invalid selections

### Estimated Effort
**32-40 hours** across 6 phases (includes Apply/Cancel, preview, undo)

---

## 2. Architecture: Dual-Mode System

### Mode 1: Single Selection (Default)
- Click tree item → select it
- Properties panel shows all fields (editable)
- **Auto-save:** Changes save immediately on field blur (current behavior)
- Add/Remove work on single node
- Current behavior preserved

### Mode 2: Multi-Selection (Bulk Mode)
- Check multiple nodes → enter "Bulk Mode"
- Properties panel shows:
  - **Same values:** Displayed, editable
  - **Mixed values:** "[Mixed]" badge, default "No change" state
  - **Dirty tracking:** Only modified fields are applied
  - **Warning banner:** "⚠️ Editing N <type> elements"
  - **Apply/Cancel buttons:** Explicit confirmation required (replaces auto-save)
  - **Preview button:** Shows what will change before apply
- Add/Remove work on all checked nodes
- **Field restrictions:** Municipality/Variant disabled for AreaScheme in bulk mode
- Strict validation: same type + same scheme

---

## 3. Implementation Phases

### Phase 1: Core Infrastructure (8-10 hours)

#### TreeNode Changes
```python
class TreeNode:
    def __init__(self, ...):
        self.IsChecked = False  # NEW
        self._check_changed_callback = None
```

#### XAML Changes
```xml
<CheckBox IsChecked="{Binding IsChecked, Mode=TwoWay}" 
          Margin="0,0,5,0"
          ToolTip="Select for bulk operations"/>
```

#### New State Management
```python
class CalculationSetupWindow:
    _checked_nodes = []           # List[TreeNode]
    _selection_mode = "single"    # "single" | "multi"
    _check_constraints = None     # (ElementType, AreaSchemeId)
    _dirty_fields = set()         # NEW: Track modified fields in bulk mode
    _bulk_undo_snapshot = {}      # NEW: {ElementId: data_dict} for undo
    
    def on_node_checked(self, node, is_checked):
        # Validate against constraints
        # Update _checked_nodes
        # Switch mode
        # Update UI (show/hide Apply/Cancel buttons)
        # Clear dirty tracking on mode switch
    
    def _validate_check(self, node):
        # Enforce same ElementType + same AreaScheme
    
    def _get_areascheme_for_node(self, node):
        # Return AreaScheme ElementId for any node type
```

**Files Modified:**
- `CalculationSetup_script.py` (lines 31-100, add new methods)
- `CalculationSetupWindow.xaml` (lines 55-65)

---

### Phase 2: Bulk Properties Panel with Apply/Cancel (12-14 hours)

#### Mixed Value Detection with "No Change" Default
```python
def _update_bulk_properties_panel(self):
    # Collect values from all checked nodes
    # Detect mixed vs same
    # Create fields with [Mixed] badges, default to "No change" state
    # Show Apply/Cancel/Preview buttons
    # Hide auto-save handlers (only trigger on Apply)
    # Disable Municipality/Variant fields if ElementType == "AreaScheme"

def _create_bulk_field_control(self, field_name, field_props, value, is_mixed):
    # Add [Mixed] badge if is_mixed=True
    # Set Tag="no_change" for mixed fields (not included in patch)
    # Add tooltip: "Values differ. Leave unchanged or edit to apply same value to all"
    # Wire TextChanged/SelectionChanged to mark field as dirty
```

#### Dirty Field Tracking
```python
def on_bulk_field_changed(self, sender, args):
    """Mark field as dirty when user edits it"""
    field_name = self._get_field_name_from_control(sender)
    
    # Clear "no_change" tag
    if hasattr(sender, 'Tag') and sender.Tag == "no_change":
        sender.Tag = "dirty"
    
    # Add to dirty set
    self._dirty_fields.add(field_name)
    
    # Update Preview button state (enable if any dirty fields)
    self._update_preview_button()
```

#### Bulk Change Preview
```python
def on_preview_clicked(self, sender, args):
    """Show preview of what will change"""
    if not self._dirty_fields:
        forms.alert("No changes to preview.")
        return
    
    # Build preview text
    preview = "The following changes will be applied to {} elements:\n\n".format(
        len(self._checked_nodes)
    )
    
    for field_name in sorted(self._dirty_fields):
        control = self._field_controls[field_name]
        value = self._get_control_value(control)
        preview += "  • {}: {}\n".format(field_name, value if value else "(empty)")
    
    # Show preview dialog
    forms.alert(preview, title="Preview Bulk Changes")
```

#### Bulk Apply with TransactionGroup
```python
def on_bulk_apply_clicked(self, sender, args):
    """Apply changes to all checked nodes"""
    if not self._dirty_fields:
        forms.alert("No changes to apply.")
        return
    
    # Collect only dirty field values
    data_patch = {}
    for field_name in self._dirty_fields:
        control = self._field_controls[field_name]
        value = self._get_control_value(control)
        if value is not None:
            data_patch[field_name] = value
    
    # Confirm if mixed fields were changed
    mixed_fields_changed = [f for f in self._dirty_fields 
                           if self._field_controls[f].Tag == "dirty"]
    
    if mixed_fields_changed:
        preview = "Applying changes to {} elements:\n\n".format(len(self._checked_nodes))
        preview += "\n".join("  • " + f for f in sorted(self._dirty_fields))
        preview += "\n\nContinue?"
        
        if not forms.alert(preview, yes=True, no=True):
            return
    
    # Save undo snapshot before applying
    self._save_undo_snapshot()
    
    # Apply in TransactionGroup for atomic rollback on error
    try:
        with revit.TransactionGroup("Bulk Update pyArea Data"):
            success_count = 0
            with revit.Transaction("Apply Bulk Changes"):
                for node in self._checked_nodes:
                    # Merge patch with existing data
                    existing_data = data_manager.get_data(node.Element) or {}
                    existing_data.update(data_patch)
                    
                    if data_manager.set_data(node.Element, existing_data):
                        success_count += 1
            
            if success_count == len(self._checked_nodes):
                # Success - commit group
                self._dirty_fields.clear()
                self.update_properties_panel()  # Refresh to show saved values
                forms.toast("✓ Updated {} elements".format(success_count))
            else:
                # Partial failure - rollback
                raise Exception("Only {} of {} elements updated".format(
                    success_count, len(self._checked_nodes)
                ))
    
    except Exception as e:
        forms.alert("Error: {}".format(e))

def on_bulk_cancel_clicked(self, sender, args):
    """Cancel changes and refresh properties panel"""
    if self._dirty_fields:
        if forms.alert("Discard unsaved changes?", yes=True, no=True):
            self._dirty_fields.clear()
            self.update_properties_panel()
    else:
        self._dirty_fields.clear()
        self.update_properties_panel()
```

**Files Modified:**
- `CalculationSetup_script.py` (lines 752-1094, 1198-1242)

---

### Phase 3: Bulk Operations (4-6 hours)

#### Button Logic Updates
```python
def _update_add_button_text(self):
    if self._selection_mode == "multi":
        # Show: "➕ Bulk Add..." or "(Not supported)"
        # Show: "🗑 Remove (N)"
```

#### Bulk Remove
```python
def _bulk_remove(self):
    # Confirm with count
    # Handle each ElementType:
    #   - RepresentedAreaPlan: remove from parent's list
    #   - AreaScheme: cascade to children
    #   - Others: delete data
    # Clear checked state
    # Rebuild tree
```

#### Bulk Track AreaPlans on Sheets
```python
def _bulk_track_areaplans_on_sheets(self):
    """
    Add AreaPlan views to tracking for multiple sheets.
    NOTE: Does NOT place views on sheets (Revit limitation).
    Only ensures views have JSON data so they appear in tree.
    """
    # Show view picker (multi-select)
    # For each selected view:
    #   - Ensure view has data (even if empty) via data_manager.set_data()
    #   - This makes view visible in tree under its sheet
    # One TransactionGroup for atomic rollback
    # Toast: "Tracking N view(s) for M sheet(s)"
```

**Files Modified:**
- `CalculationSetup_script.py` (lines 718-750, 1255-1291, 1919-2014)

---

### Phase 3b: Bulk Set Representing View (3-4 hours)

#### Move Multiple Children to Single Parent
```python
def _bulk_set_representing_view(self):
    """
    Move multiple AreaPlan_NotOnSheet or RepresentedAreaPlan nodes to a single parent.
    Validates: all children not on sheets, parent on sheet + same scheme.
    Flattens nested represented views like single-move logic.
    """
    # Verify all checked nodes are AreaPlan_NotOnSheet or RepresentedAreaPlan
    if not all(n.ElementType in ["AreaPlan_NotOnSheet", "RepresentedAreaPlan"] 
               for n in self._checked_nodes):
        forms.alert("Bulk set representing view only supports views not on sheets.")
        return
    
    # Verify none are on sheets (edge case validation)
    views_on_sheets = self._get_views_on_sheets()
    for node in self._checked_nodes:
        if node.Element.Id in views_on_sheets:
            forms.alert("Cannot set representing view for '{}' - it's on a sheet.".format(
                node.DisplayName
            ))
            return
    
    # Get area scheme
    area_scheme = self._checked_nodes[0].Element.AreaScheme
    
    # Get available parents (on sheets, same scheme, not already represented)
    available_parents = self._get_available_parent_views(area_scheme)
    
    if not available_parents:
        forms.alert("No available parent views.\n\n"
                   "Parent must be:\n"
                   "- Same AreaScheme\n"
                   "- Placed on a sheet\n"
                   "- Not already represented by another view")
        return
    
    # Show parent picker (single-select)
    class ParentOption(forms.TemplateListItem):
        def __init__(self, view):
            super(ParentOption, self).__init__(view, checked=False)
        
        @property
        def name(self):
            return "■ {}".format(self.item.Name if hasattr(self.item, 'Name') else "?")
    
    options = [ParentOption(v) for v in available_parents]
    
    selected_parent = forms.SelectFromList.show(
        options,
        title="Select Parent View for {} Children".format(len(self._checked_nodes)),
        button_name="Set Parent"
    )
    
    if not selected_parent:
        return
    
    parent_view = selected_parent.item if isinstance(selected_parent, ParentOption) else selected_parent
    
    # Apply in TransactionGroup
    try:
        with revit.TransactionGroup("Bulk Set Representing View"):
            with revit.Transaction("Move Children to Parent"):
                # Remove from current parents (if any)
                for node in self._checked_nodes:
                    if node.Parent and node.Parent.ElementType in ["AreaPlan", "AreaPlan_NotOnSheet"]:
                        parent_data = data_manager.get_data(node.Parent.Element) or {}
                        rep_ids = parent_data.get("RepresentedViews", [])
                        view_id_str = str(node.Element.Id.Value)
                        
                        if view_id_str in rep_ids:
                            rep_ids.remove(view_id_str)
                        
                        if rep_ids:
                            parent_data["RepresentedViews"] = rep_ids
                        else:
                            parent_data.pop("RepresentedViews", None)
                        
                        data_manager.set_data(node.Parent.Element, parent_data)
                
                # Get new parent's data
                new_parent_data = data_manager.get_data(parent_view) or {}
                rep_ids = new_parent_data.get("RepresentedViews", [])
                
                if not isinstance(rep_ids, list):
                    rep_ids = []
                
                # Add all checked children (and flatten any nested represented views)
                for node in self._checked_nodes:
                    view_id_str = str(node.Element.Id.Value)
                    
                    # Add to parent
                    if view_id_str not in rep_ids:
                        rep_ids.append(view_id_str)
                    
                    # Flatten: if child has its own represented views, add them to parent
                    child_data = data_manager.get_data(node.Element)
                    if child_data and "RepresentedViews" in child_data:
                        nested_ids = child_data.get("RepresentedViews", [])
                        for nested_id in nested_ids:
                            if nested_id not in rep_ids:
                                rep_ids.append(nested_id)
                        
                        # Remove RepresentedViews from child (flatten hierarchy)
                        child_data.pop("RepresentedViews", None)
                        data_manager.set_data(node.Element, child_data)
                
                # Save parent's updated list
                new_parent_data["RepresentedViews"] = rep_ids
                data_manager.set_data(parent_view, new_parent_data)
        
        # Clear checked state and rebuild
        self._checked_nodes.clear()
        self._check_constraints = None
        self._selection_mode = "single"
        self.rebuild_tree()
        
        forms.toast("✓ Moved {} view(s) to '{}'".format(
            len(self._checked_nodes), parent_view.Name
        ))
    
    except Exception as e:
        forms.alert("Error: {}".format(e))
```

**Files Modified:**
- `CalculationSetup_script.py` (new method, ~150 lines)

---

### Phase 4: State Preservation (2-3 hours)

#### Preserve Checked State Across Rebuilds
```python
def rebuild_tree(self):
    # Save checked_ids before rebuild
    # Build tree
    # Restore expansion (existing)
    # Restore checked state (NEW)

def _restore_checked_state(self, checked_ids):
    # Traverse tree
    # Re-check nodes by ElementId
    # Restore _checked_nodes and _check_constraints
    # Update UI
```

**Files Modified:**
- `CalculationSetup_script.py` (lines 437-441, add new method)

---

### Phase 5: Polish & Safety (3-5 hours)

#### Visual Enhancements
- Highlight checked nodes (light blue background with border)
- Bulk mode banner above tree: "⚠️ Bulk Mode: N <type> selected | Apply/Cancel/Preview/Clear All"
- "Clear All" button to uncheck all nodes
- "Select All Visible" button (same type + scheme constraint)
- Apply/Cancel/Preview/Undo buttons in properties panel (bulk mode only)
- Selection count badge in title

#### Safety Features  
- **Max 50 checked nodes** (hard limit, warn at 40)
- **Performance optimization:** Cache sheets/views collectors at session start
- **TransactionGroup** for atomic rollback on any error
- Toast notifications for all bulk operations
- Comprehensive error handling with partial success reporting
- **Keyboard shortcuts (optional):**
  - Ctrl+Click: Toggle checkbox
  - Shift+Click: Range select (if constraints pass)

#### Button State Management
```python
def _update_bulk_mode_buttons(self):
    """Update visibility/state of bulk mode buttons"""
    if self._selection_mode == "multi":
        # Show bulk buttons
        self.btn_bulk_apply.Visibility = System.Windows.Visibility.Visible
        self.btn_bulk_cancel.Visibility = System.Windows.Visibility.Visible
        self.btn_bulk_preview.Visibility = System.Windows.Visibility.Visible
        self.btn_bulk_undo.Visibility = System.Windows.Visibility.Visible
        
        # Enable/disable based on state
        self.btn_bulk_apply.IsEnabled = len(self._dirty_fields) > 0
        self.btn_bulk_preview.IsEnabled = len(self._dirty_fields) > 0
        self.btn_bulk_undo.IsEnabled = bool(self._bulk_undo_snapshot)
    else:
        # Hide bulk buttons
        self.btn_bulk_apply.Visibility = System.Windows.Visibility.Collapsed
        self.btn_bulk_cancel.Visibility = System.Windows.Visibility.Collapsed
        self.btn_bulk_preview.Visibility = System.Windows.Visibility.Collapsed
        self.btn_bulk_undo.Visibility = System.Windows.Visibility.Collapsed
```

**Files Modified:**
- `CalculationSetupWindow.xaml` (add styles, banner, bulk buttons)
- `CalculationSetup_script.py` (add limits, confirmations, keyboard handlers)

---

### Phase 6: Session-Scoped Undo (2-3 hours)

#### Undo Last Bulk Change
```python
def _save_undo_snapshot(self):
    """Save current state of checked nodes before bulk apply"""
    self._bulk_undo_snapshot = {}
    
    for node in self._checked_nodes:
        elem_id = node.Element.Id
        data = data_manager.get_data(node.Element)
        # Deep copy to avoid reference issues
        self._bulk_undo_snapshot[elem_id] = dict(data) if data else {}

def on_bulk_undo_clicked(self, sender, args):
    """Undo last bulk change (session-scoped only)"""
    if not self._bulk_undo_snapshot:
        forms.alert("No bulk change to undo.")
        return
    
    # Confirm
    if not forms.alert(
        "Undo last bulk change for {} elements?".format(len(self._bulk_undo_snapshot)),
        yes=True, no=True
    ):
        return
    
    # Restore from snapshot
    try:
        with revit.TransactionGroup("Undo Bulk Change"):
            success_count = 0
            with revit.Transaction("Restore Previous State"):
                for elem_id, data in self._bulk_undo_snapshot.items():
                    element = self._doc.GetElement(elem_id)
                    if element:
                        if data_manager.set_data(element, data):
                            success_count += 1
            
            if success_count == len(self._bulk_undo_snapshot):
                # Clear snapshot after successful undo (can't undo twice)
                self._bulk_undo_snapshot.clear()
                self._dirty_fields.clear()
                self.update_properties_panel()
                forms.toast("✓ Undone changes for {} elements".format(success_count))
            else:
                raise Exception("Only {} of {} elements restored".format(
                    success_count, len(self._bulk_undo_snapshot)
                ))
    
    except Exception as e:
        forms.alert("Error during undo: {}".format(e))
```

#### Undo Limitations
- **Session-scoped only:** Undo clears on dialog close or new bulk operation
- **Single-level undo:** No redo, no undo stack
- **Bulk operations only:** Single-node changes use Revit's native undo

**Files Modified:**
- `CalculationSetup_script.py` (add undo methods, snapshot management)
- `CalculationSetupWindow.xaml` (add Undo button)

---

## 4. Technical Specifications

### Validation Rules

#### Same ElementType Constraint
```python
allowed_combinations = [
    ["AreaScheme"],
    ["Sheet"],
    ["AreaPlan"],
    ["AreaPlan_NotOnSheet"],
    ["RepresentedAreaPlan"]
]
# Cannot mix: ["Sheet", "AreaPlan"]
# Cannot mix: ["AreaPlan", "RepresentedAreaPlan"]
```

#### Same AreaScheme Constraint
```python
# For AreaScheme: use node.Element.Id
# For Sheet: use data["AreaSchemeId"] from JSON
# For AreaPlan/RepresentedAreaPlan: use node.Element.AreaScheme.Id
```

### Bulk Operations Support Matrix

| Element Type | Bulk Remove | Bulk Track Views | Bulk Set Parent | Bulk Edit Fields | Field Restrictions |
|--------------|-------------|------------------|-----------------|------------------|--------------------|
| AreaScheme | ✅ (cascade) | ❌ | ❌ | ✅ | Municipality/Variant disabled |
| Sheet | ✅ | ✅ | ❌ | ✅ | None |
| AreaPlan | ✅ | ❌ | ❌ | ✅ | None |
| AreaPlan_NotOnSheet | ✅ | ❌ | ✅ | ✅ | None |
| RepresentedAreaPlan | ✅ (from parent) | ❌ | ✅ | ✅ | None |

**Notes:**
- **Bulk Track Views:** Adds AreaPlan views to tracking (creates JSON data, does NOT place on sheets)
- **Bulk Set Parent:** Moves multiple children to single parent view (on sheet, same scheme)
- **Field Restrictions:** Municipality/Variant editing disabled for AreaScheme in bulk mode to prevent structural changes

---

## 5. Risk Assessment

| Risk | Impact | Mitigation |
|------|--------|------------|
| **Accidental bulk overwrites** | High | Confirmation dialogs for mixed values, preview changes |
| **Performance (100+ nodes)** | Medium | Limit to 50 checked nodes, optimize tree rebuild |
| **State loss on rebuild** | Medium | Preserve checked_ids + expansion paths |
| **Cross-scheme selection** | High | Strict validation, clear error messages |
| **UI complexity** | Low | Clear mode indicators, help tooltips |

---

## 6. Testing Strategy

### Unit Tests (Manual)
- [ ] Check single node → enter bulk mode
- [ ] Check different types → validation rejects
- [ ] Check different schemes → validation rejects
- [ ] Uncheck all → return to single mode
- [ ] Edit mixed field → show confirmation
- [ ] Edit same field → apply to all
- [ ] Bulk remove Sheets (5) → all removed
- [ ] Bulk add AreaPlan to Sheets (3) → all updated
- [ ] Rebuild tree → checked state preserved
- [ ] Expansion state preserved with checked nodes

### Edge Cases
- [ ] Check node, delete it externally, rebuild → graceful handling
- [ ] Check represented view, place on sheet → validation handles
- [ ] Bulk edit Municipality field → updates correctly
- [ ] Mixed boolean field (three-state checkbox) → works
- [ ] Check 50 nodes → performance acceptable
- [ ] Check 51 nodes → warning shown

---

## 7. Implementation Timeline

| Phase | Duration | Dependencies |
|-------|----------|--------------|
| Phase 1: Core Infrastructure | 8-10 hours | None |
| Phase 2: Bulk Properties + Apply/Cancel | 12-14 hours | Phase 1 complete |
| Phase 3: Bulk Operations | 4-6 hours | Phase 1 complete |
| Phase 3b: Bulk Set Representing View | 3-4 hours | Phase 1 complete |
| Phase 4: State Preservation | 2-3 hours | Phase 1 complete |
| Phase 5: Polish & Safety | 3-5 hours | Phases 1-4 complete |
| Phase 6: Session-Scoped Undo | 2-3 hours | Phase 2 complete |
| **Testing & Refinement** | 6-8 hours | All phases complete |
| **Total** | **40-53 hours** | - |

---

## 8. Alternative Approaches (Rejected)

### Keyboard-Driven Multi-Select (Ctrl/Shift)
**Pros:** Native feel, no UI clutter  
**Cons:** WPF TreeView doesn't support it natively, fragile to implement, harder to maintain  
**Decision:** Rejected - too complex for benefit

### Two Separate Trees (Single + Multi)
**Pros:** Complete separation of concerns  
**Cons:** Confusing UX, data duplication  
**Decision:** Rejected - poor UX

### Read-Only Bulk Properties (No Editing)
**Pros:** Simplest, safest  
**Cons:** Limited utility, user requested editing  
**Decision:** Rejected - doesn't meet requirements

---

## 9. Migration Strategy

### Backwards Compatibility
- ✅ No data schema changes
- ✅ Single-selection mode unchanged
- ✅ Existing workflows unaffected
- ✅ Checkboxes don't interfere with tree navigation

### Rollout Plan
1. **Branch code** before starting
2. **Implement Phase 1** → test thoroughly
3. **Implement Phases 2-3** → test bulk operations
4. **Implement Phases 4-5** → test state preservation
5. **User testing** with sample project
6. **Merge to main** after approval

---

## 10. Success Criteria

- [ ] Can select 10 sheets, edit FLOOR_NAME, apply to all in one action
- [ ] Can select 5 AreaPlans, bulk remove, confirm all gone
- [ ] Mixed values show "[Mixed]" badge and require confirmation
- [ ] Checked state survives tree rebuild
- [ ] Cannot check nodes from different AreaSchemes (validation works)
- [ ] Performance acceptable with 50 checked nodes
- [ ] No regressions in single-selection mode
- [ ] User can complete workflows 50% faster

---

## 11. Design Decisions (Approved by GPT Review)

1. **Checkbox visibility:** ✅ Always visible (clearer UX)

2. **Max checked nodes:** ✅ 50 hard limit with warning at 40

3. **Bulk edit mode:** ✅ Explicit Apply/Cancel (replaces auto-save in bulk mode)

4. **Mixed value handling:** ✅ Default "No change", only apply dirty fields

5. **Bulk change preview:** ✅ Show preview dialog before apply

6. **"Clear All" button:** ✅ Add button (faster workflow)

7. **Auto-check children:** ❌ No (too implicit, validation complexity)

8. **Bulk add RepresentedViews to multiple parents:** ❌ No (complex, rare use case)

9. **Bulk Set Representing View:** ✅ Yes, move multiple children to single parent

10. **Field restrictions in bulk mode:** ✅ Disable Municipality/Variant for AreaScheme

11. **Session-scoped undo:** ✅ Single-level undo for last bulk operation

12. **Keyboard shortcuts:** ✅ Optional (Ctrl+Click toggle, Shift+Click range)

---

## 12. Next Steps

1. ✅ **Incorporate GPT improvements** - DONE
2. **Final user approval** of design decisions (section 11)
3. **Create feature branch:** `feature/multiselect-bulk-ops`
4. **Begin Phase 1 implementation** (Core Infrastructure)
5. **Incremental testing** after each phase
6. **User acceptance testing** with sample project before merge

---

## 13. Summary of GPT Improvements

### Key Enhancements Added
1. **Explicit Apply/Cancel in bulk mode** (replaces auto-save for safety)
2. **Dirty field tracking** (only modified fields applied)
3. **"No change" default for mixed values** (safer, more explicit)
4. **Preview dialog** before applying bulk changes
5. **TransactionGroup** for atomic rollback on errors
6. **Field restrictions** (Municipality/Variant disabled for AreaScheme in bulk)
7. **Bulk Set Representing View** (Phase 3b) - move multiple children to single parent
8. **Session-scoped undo** (Phase 6) - single-level undo for last bulk operation
9. **Keyboard shortcuts** (optional Ctrl/Shift+Click)
10. **Performance optimizations** (cache collectors, 50-node limit)

### Timeline Impact
- Original estimate: 28-34 hours
- Updated estimate: **40-53 hours** (includes safety features and undo)
- Additional time justified by significantly improved safety and UX

---

**Document Status:** ✅ Updated with GPT review (v2.0)  
**Approval Required:** Yes (final design decisions review)  
**Estimated Start Date:** TBD after final approval  
**Implementation Confidence:** High (detailed specifications, proven patterns)
