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

### 5. **Curve Loop Winding/Containment Analysis**
Use 2D computational geometry instead of 3D booleans.

**How it works:**
- Flatten all boundaries to 2D polygons
- Use polygon boolean library (Clipper/shapely via IronPython)
- Find difference between bounding rectangle and union of polygons

**Pros:**
- No Revit geometry failures
- Well-tested algorithms
- Fast

**Cons:**
- Requires external library or custom implementation
- Loses arc precision (must tessellate)

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

**Hybrid: Solutions 1 + 6 + 8**

1. **First**: Extract interior holes from individual areas (Solution 6) - guaranteed to work
2. **Then**: Progressive union with tracking (Solution 8) - get what we can from 3D
3. **Finally**: For failed areas, probe their boundaries (Solution 1) - catches gaps near failures

This gives us:
- ✅ Interior holes (donuts) - from individual area analysis
- ✅ Gaps where union succeeded - from 3D boolean
- ✅ Gaps near failed areas - from boundary probing

---

## Implementation Priority

| Solution | Effort | Reliability | Speed | Priority |
|----------|--------|-------------|-------|----------|
| 1. Boundary Probing | Medium | High | Fast | ⭐⭐⭐ |
| 6. Interior Loop Extraction | Low | High | Fast | ⭐⭐⭐ |
| 8. Progressive Union | Medium | Medium | Medium | ⭐⭐ |
| 2. Loop Inflation | High | Medium | Fast | ⭐ |
| 5. 2D Polygon Boolean | High | High | Fast | ⭐ |
