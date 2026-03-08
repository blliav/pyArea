# -*- coding: utf-8 -*-
"""DXF export helper functions.

Municipality-specific geometry routines that are too large or specialised
to live in the main ExportDXF_script.  All functions in this module
operate on plain (x, y) coordinate tuples — no Revit API dependency.
"""

import math
from collections import defaultdict


def get_cluster_frames_for_telaviv(area_polylines):
    """Compute Tel-Aviv cluster frames by merging area boundary segments.

    Instead of reconstructing arcs after WPF flattening, this approach reuses
    each area's already-processed boundary polyline (points + bulges), removes
    shared interior edges (which appear twice from adjacent areas), and chains
    the remaining exterior edges into closed frame polylines.

    Arc segments are preserved exactly as-is from the original area polylines.

    Args:
        area_polylines: List of (dxf_pts, bulges) tuples where:
            dxf_pts: closed list of (x, y) tuples
            bulges: list of bulge values aligned with dxf_pts

    Returns:
        List of (dxf_pts, bulges) tuples where:
            dxf_pts: list of (x, y) tuples (closed polyline)
            bulges: list of bulge values (or None if no arcs)
    """
    if not area_polylines:
        return []

    # Rounding precision for coordinate matching.
    # Use 1 decimal (0.1 cm = 1 mm) so shared edges from adjacent areas
    # cancel even if floating-point transforms produce slightly different
    # coordinates for the same wall.  Individual exterior edges are much
    # longer than 1 mm so this won't merge distinct vertices.
    key_decimals = 1

    def _pt_key(pt):
        return (round(pt[0], key_decimals), round(pt[1], key_decimals))

    def _edge_key(pt_a, pt_b):
        """Normalized (direction-independent) key for an edge."""
        ka = _pt_key(pt_a)
        kb = _pt_key(pt_b)
        return (ka, kb) if ka <= kb else (kb, ka)

    # Step 1: Collect boundary segments from each area polyline
    all_segments = []
    for dxf_pts, bulges in area_polylines:
        if not dxf_pts or len(dxf_pts) < 2:
            continue

        local_bulges = bulges if bulges else [0.0] * len(dxf_pts)
        # Polyline is closed: last point == first point.
        # Segments go from dxf_pts[i] to dxf_pts[i+1].
        seg_count = min(len(dxf_pts) - 1, len(local_bulges))
        for i in range(seg_count):
            s = dxf_pts[i]
            e = dxf_pts[i + 1]
            if _pt_key(s) == _pt_key(e):
                continue  # skip zero-length
            all_segments.append((s, e, local_bulges[i]))

    if not all_segments:
        return []

    # Step 2: Cancel shared interior edges
    # Shared edges appear twice (once per area, opposite directions).
    edge_buckets = {}
    for seg in all_segments:
        key = _edge_key(seg[0], seg[1])
        if key not in edge_buckets:
            edge_buckets[key] = []
        edge_buckets[key].append(seg)

    exterior_segments = []
    for seg_list in edge_buckets.values():
        if len(seg_list) == 1:
            # Unique edge — exterior
            exterior_segments.append(seg_list[0])
        elif len(seg_list) % 2 == 1:
            # Odd count — one leftover exterior segment
            exterior_segments.append(seg_list[0])
        # Even count — all shared, discard

    if not exterior_segments:
        return []

    # Step 3: Build directed adjacency — for each segment store both its
    # forward and reverse orientations so we can orient edges at chain time.
    adj = defaultdict(list)  # pt_key -> list of (seg_idx, is_forward)
    for idx, (s, e, b) in enumerate(exterior_segments):
        adj[_pt_key(s)].append((idx, True))   # forward: s -> e
        adj[_pt_key(e)].append((idx, False))  # reverse: e -> s

    def _seg_endpoints(idx, forward):
        """Get (start, end, bulge) for a segment in the given direction."""
        s, e, b = exterior_segments[idx]
        if forward:
            return s, e, b
        else:
            return e, s, -b

    def _angle_of(dx, dy):
        """Angle in radians from positive X axis, range [0, 2*pi)."""
        a = math.atan2(dy, dx)
        if a < 0:
            a += 2.0 * math.pi
        return a

    def _pick_leftmost(current_key, prev_pt, candidates, used):
        """At a junction, pick the unused edge that turns leftmost (CCW).
        
        This keeps the traversal on the outermost boundary and avoids
        diving into interior spikes.
        
        'prev_pt' is the point we came from; we compute the incoming
        direction and pick the candidate whose outgoing direction makes
        the largest CCW turn (smallest CW turn).
        """
        if not candidates:
            return None
        # Incoming direction vector (from prev_pt to current_key)
        cur = current_key  # already rounded tuple
        in_dx = cur[0] - prev_pt[0]
        in_dy = cur[1] - prev_pt[1]
        in_angle = _angle_of(in_dx, in_dy)

        best = None
        best_turn = -1.0
        for seg_idx, fwd in candidates:
            if used[seg_idx]:
                continue
            cs, ce, cb = _seg_endpoints(seg_idx, fwd)
            out_dx = ce[0] - cs[0]
            out_dy = ce[1] - cs[1]
            out_angle = _angle_of(out_dx, out_dy)
            # CW turn from incoming to outgoing (we want the one with
            # the largest CCW turn, i.e. smallest CW turn).
            # CCW turn = (out_angle - in_angle + pi) mod 2pi
            # We want the LARGEST CCW turn to stay on the exterior.
            turn = (out_angle - in_angle + math.pi) % (2.0 * math.pi)
            if turn > best_turn:
                best_turn = turn
                best = (seg_idx, fwd)
        return best

    used = [False] * len(exterior_segments)
    all_loops = []

    for seed_idx in range(len(exterior_segments)):
        if used[seed_idx]:
            continue
        used[seed_idx] = True

        s, e, b = exterior_segments[seed_idx]
        loop_start_key = _pt_key(s)
        current_key = _pt_key(e)
        prev_pt = s  # actual coordinates of previous point
        oriented = [(s, e, b)]

        guard = 0
        while current_key != loop_start_key and guard < 100000:
            guard += 1
            pick = _pick_leftmost(current_key, prev_pt, adj[current_key], used)
            if pick is None:
                break
            seg_idx, fwd = pick
            used[seg_idx] = True
            cs, ce, cb = _seg_endpoints(seg_idx, fwd)
            oriented.append((cs, ce, cb))
            prev_pt = cs
            current_key = _pt_key(ce)

        # Only keep closed loops
        if current_key != loop_start_key or not oriented:
            continue

        # Build the polyline points and bulges from the oriented chain
        frame_pts = [oriented[0][0]]
        frame_bulges = []
        for _s, _e, _b in oriented:
            frame_bulges.append(_b)
            frame_pts.append(_e)

        # Ensure closure
        if _pt_key(frame_pts[-1]) != _pt_key(frame_pts[0]):
            frame_pts.append(frame_pts[0])
            frame_bulges.append(0.0)

        # Step 3b: Remove self-intersection spikes from the loop.
        # Interior wall edges that survived cancellation cause the loop
        # to visit the same vertex twice, forming a "lasso" detour.
        # We detect repeated vertex keys and excise the smaller sub-loop
        # (the spike), keeping the larger one (the exterior).
        # Work on the N "real" vertices (exclude the closing duplicate).
        n_segs = len(frame_bulges)  # == len(frame_pts) - 1
        verts = [(frame_pts[i], frame_bulges[i]) for i in range(n_segs)]

        cleaned = True
        while cleaned:
            cleaned = False
            seen = {}  # pt_key -> index in verts
            for i in range(len(verts)):
                pk = _pt_key(verts[i][0])
                if pk in seen:
                    j = seen[pk]
                    # verts[j..i-1] is the detour sub-loop
                    detour_len = i - j
                    main_len = len(verts) - detour_len
                    if main_len >= 3:
                        # Excise the detour: keep verts[0..j-1] + verts[i..]
                        verts = verts[:j] + verts[i:]
                        cleaned = True
                        break
                    else:
                        # Detour is the main part — keep the detour instead
                        verts = verts[j:i]
                        cleaned = True
                        break
                seen[pk] = i

        if len(verts) < 3:
            continue

        # Rebuild frame_pts and frame_bulges
        frame_pts = [v[0] for v in verts]
        frame_bulges = [v[1] for v in verts]
        # Re-close
        frame_pts.append(frame_pts[0])

        has_arcs = any(abs(b) > 1e-12 for b in frame_bulges)
        all_loops.append((frame_pts, frame_bulges if has_arcs else None))

    # Step 4: Keep only exterior loops (discard islands / holes).
    # Exterior loops have the largest absolute area. For each group of
    # loops, keep only the outermost one. Since we only need the exterior
    # perimeter, keep only the loop with the largest absolute signed area.
    def _loop_signed_area(pts):
        """Shoelace signed area for a list of (x, y) tuples."""
        n = len(pts)
        if n < 3:
            return 0.0
        sa = 0.0
        for i in range(n):
            j = (i + 1) % n
            sa += pts[i][0] * pts[j][1]
            sa -= pts[j][0] * pts[i][1]
        return sa / 2.0

    if not all_loops:
        return []

    # Compute absolute area for each loop
    loop_areas = []
    for pts, bulges in all_loops:
        loop_areas.append(abs(_loop_signed_area(pts)))

    # Find the largest loop — that is the exterior perimeter.
    # If there are multiple disconnected clusters (separate groups of
    # touching areas), we need to keep one exterior loop per cluster.
    # Use point-in-polygon to detect which smaller loops are inside larger ones.
    def _point_in_polygon(px, py, polygon):
        """Ray-casting point-in-polygon test."""
        n = len(polygon)
        inside = False
        j = n - 1
        for i in range(n):
            xi, yi = polygon[i][0], polygon[i][1]
            xj, yj = polygon[j][0], polygon[j][1]
            if ((yi > py) != (yj > py)) and (px < (xj - xi) * (py - yi) / (yj - yi) + xi):
                inside = not inside
            j = i
        return inside

    # Sort loops by area descending
    indexed_loops = sorted(enumerate(all_loops), key=lambda x: loop_areas[x[0]], reverse=True)

    dxf_frames = []
    for idx, (pts, bulges) in indexed_loops:
        # Check if this loop's centroid is inside any already-accepted larger loop
        cx = sum(p[0] for p in pts) / len(pts)
        cy = sum(p[1] for p in pts) / len(pts)
        is_interior = False
        for accepted_pts, _ in dxf_frames:
            if _point_in_polygon(cx, cy, accepted_pts):
                is_interior = True
                break
        if not is_interior:
            dxf_frames.append((pts, bulges))

    return dxf_frames
