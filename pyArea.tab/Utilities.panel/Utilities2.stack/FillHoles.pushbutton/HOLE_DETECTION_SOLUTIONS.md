# Hole Detection Solutions - Brainstorming

## The Problem
When areas fail to convert to solids (due to self-intersecting curves, thin slivers, etc.), the boolean union approach cannot detect holes/gaps between them.

---

## Solution Ideas

### 1. **Boundary Segment Probing** ❌ TESTED - TOO SLOW
Place test areas at segment ends, slightly offset to both sides.

**How it works:**
- For each area boundary segment end, go back 0.5ft and offset perpendicular
- Try `doc.Create.NewArea(view, uv)` at each offset point
- If Revit creates an area → it's a gap (keep it)
- If fails → area already exists there (rollback)

**Test Results:**
- 31 areas → 3152 probe points → 1967 candidates after pre-filter
- Even with 2D containment pre-filter, too many SubTransactions
- **Very slow on real projects** (minutes for large models)
- Pre-filter using `_point_in_curveloop_2d` helps but not enough

**Why it's slow:**
- Each segment has 2 ends × 2 sides = 4 probe points
- Complex boundaries have many segments
- SubTransaction per probe point is expensive

**Verdict:** Works but not viable for production. Need transaction-free approach.

---

### 2. **Loop Expansion/Inflation** ⭐ (Your Idea)
Slightly expand area boundaries before union so they overlap.

**How it works:**
- For each curve loop, offset all curves outward by ~0.01ft
- Overlapping boundaries will union successfully
- Holes remain as true interior voids

**Pros:**
- Prevents gaps from failed unions
- Simple geometric transformation

**Cons:**
- Complex to implement curve offsetting correctly
- May create new self-intersections
- Arcs/splines harder to offset

```python
def _inflate_curveloop(curve_loop, offset=0.01):
    # Use CurveLoop.CreateViaOffset if available (Revit 2022+)
    try:
        return DB.CurveLoop.CreateViaOffset(curve_loop, offset, DB.XYZ.BasisZ)
    except:
        # Manual offset for older versions
        pass
```

---

### 3. **Area Scheme Boundary Subtraction**
Use the area scheme's overall boundary as the "universe" and subtract all areas.

**How it works:**
- Get the area scheme boundary (outer limit)
- Create solid from scheme boundary
- Subtract each area's solid individually
- Remainder = all gaps

**Pros:**
- Doesn't require union of all areas
- Each subtraction is independent (one failure doesn't cascade)

**Cons:**
- Need to identify scheme boundary
- Many boolean operations

---

### 4. **Revit's Built-in Area Detection**
Let Revit tell us where areas can be created.

**How it works:**
- Use a coarse grid (50ft spacing) just to find candidate regions
- For each candidate, use binary search to find actual gap boundaries
- Much faster than fine grid

**Pros:**
- Leverages Revit's spatial engine
- Accurate

**Cons:**
- Still O(n) on grid size

---

### 5. **2D Polygon Boolean Operations (WPF)** ✅ IMPLEMENTED - FINAL SOLUTION
Use 2D computational geometry instead of 3D booleans.

**FINAL ALGORITHM:**

```
┌─────────────────────────────────────────────────────────────────┐
│  1. EXTRACT BOUNDARIES                                           │
│     For each area:                                               │
│     ├── Get boundary segments from Revit                        │
│     ├── Tessellate curves → (x,y) points                        │
│     └── Identify: Largest loop = exterior                       │
│                   Smaller loops = interior holes (gaps!)         │
│                                                                  │
│  2. CREATE WPF POLYGONS                                          │
│     └── Polygon2D(points=exterior_points)                       │
│                                                                  │
│  3. FIND GAPS BETWEEN AREAS                                      │
│     ├── Create bounding box around all polygons                 │
│     ├── Subtract each polygon from bbox (individually!)         │
│     ├── Extract contours from result                            │
│     └── Filter: Largest = outer margin (skip)                   │
│                 Others = interior gaps (fill!)                  │
│                                                                  │
│  4. CREATE AREAS AT GAP CENTROIDS                                │
│     ├── Find interior point for each gap (handles non-convex)  │
│     └── doc.Create.NewArea(view, point)                         │
└─────────────────────────────────────────────────────────────────┘
```

**Key Files:**
- `pyArea.tab/lib/polygon_2d.py` - WPF geometry wrapper
- `FillHoles_script.py` - Main script using `_find_holes_using_2d()`

**Key Functions:**
```python
# In FillHoles_script.py:
_boundary_loop_to_points(loop)      # Tessellate boundary → [(x,y), ...]
_find_holes_using_2d(areas)         # Main gap detection

# In polygon_2d.py:
Polygon2D(points=list)              # Create from points
Polygon2D.difference(other)         # Boolean subtraction
Polygon2D.get_contours()            # Extract boundary points
find_all_gap_regions_2d_from_polygons()  # Gap detection
_find_interior_point(contour)       # Find point inside non-convex polygon
```

**Why Individual Subtraction?**
Union of adjacent areas can "merge" and lose gaps. Subtracting each polygon 
individually from the bounding box preserves all non-covered regions.

**Pros:**
- No Revit geometry failures (WPF handles all cases)
- Well-tested algorithms (Microsoft's WPF geometry)
- Fast - no Revit transactions during analysis
- Works with complex geometry (self-intersections, thin slivers)
- IronPython 2.7 compatible (no external dependencies)

**Cons:**
- Tessellation approximates arcs/splines (slight precision loss)
- Large outer margin contour needs filtering

---

### 6. **Hole Loop Extraction from Individual Areas**
Look for holes IN each area, not between areas.

**How it works:**
- Many areas have interior loops (courtyards, shafts)
- Extract these directly without boolean operations
- `area.GetBoundarySegments()` returns multiple loops for donut shapes

**Pros:**
- No boolean operations needed
- Already implemented partially

**Cons:**
- Only finds interior holes, not gaps between areas

---

### 7. **Boundary Graph Analysis**
Build a graph of boundary connectivity.

**How it works:**
- Each boundary segment is an edge
- Shared segments connect two areas
- Unshared segments indicate exterior or hole boundaries
- Walk the graph to find hole perimeters

**Pros:**
- Topologically robust
- No geometry calculations

**Cons:**
- Complex implementation
- Need to match shared boundaries (tolerance issues)

---

### 8. **Progressive Union with Fallback**
Union areas one by one, track failures, handle them specially.

**How it works:**
- Union areas incrementally
- When union fails, try:
  1. Skip the problem area, continue union
  2. Mark failed area's boundary for probing (Solution 1)
  3. Try inflating the failed area (Solution 2)

**Pros:**
- Combines multiple strategies
- Graceful degradation

---

### 9. **Centroid-Based Gap Detection**
For each "potential gap region", test its centroid.

**How it works:**
- Compute Voronoi diagram of area centroids
- Test points in Voronoi cells that don't belong to any area
- These are gap candidates

**Pros:**
- Mathematically elegant
- Adaptive to area distribution

**Cons:**
- Complex to implement
- May miss small gaps

---

### 10. **Ray Casting from Area Boundaries**
Cast rays from boundary segments to find nearby gaps.

**How it works:**
- From each boundary segment midpoint, cast ray perpendicular
- Find intersection with other boundaries
- If large gap between boundary and intersection → potential void

**Pros:**
- Fast
- Finds gaps adjacent to known areas

**Cons:**
- May miss gaps not adjacent to scanned areas

---

## Recommended Approach

**Current Implementation: Solution 5 (Primary) + 3D Fallback**

1. **PRIMARY**: 2D Polygon Boolean (Solution 5) using WPF geometry
   - Convert all area boundaries to WPF PathGeometry
   - Union all polygons using CombinedGeometry
   - Find gaps by subtracting union from bounding box
   - Place areas at gap centroids

2. **FALLBACK**: 3D Boolean Union (original approach)
   - Only used if 2D module unavailable or fails
   - May have issues with complex geometry

This gives us:
- ✅ Interior holes (donuts) - from 2D polygon difference
- ✅ Gaps between areas - from 2D polygon difference
- ✅ Robust handling of complex geometry - no Revit solid failures

---

## Implementation Priority (Updated)

| Solution | Effort | Reliability | Speed | Status |
|----------|--------|-------------|-------|--------|
| 5. 2D Polygon Boolean (WPF) | Medium | High | Fast | ✅ IMPLEMENTED |
| 6. Interior Loop Extraction | Low | High | Fast | ✅ Built-in |
| 1. Boundary Probing | Medium | High | Slow | ❌ Too slow |
| 8. Progressive Union | Medium | Medium | Medium | ⭐ Fallback |
| 2. Loop Inflation | High | Medium | Fast | Not needed |
