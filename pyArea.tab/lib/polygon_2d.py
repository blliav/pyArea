# -*- coding: utf-8 -*-
"""2D Polygon Boolean Operations using WPF Geometry.

This module provides 2D polygon boolean operations using .NET's System.Windows.Media
geometry classes, which are compatible with IronPython 2.7.

MAIN USE CASE: Finding gaps between Revit Area elements
==================================================

ALGORITHM (used by FillHoles):
1. Create Polygon2D from (x,y) points: Polygon2D(points=[(x1,y1), (x2,y2), ...])
2. Create bounding box: Polygon2D.create_rectangle(min_x, min_y, max_x, max_y)
3. Subtract each area polygon from bbox: result = bbox.difference(area_polygon)
4. Extract gap contours: result.get_contours()
5. Find interior point for each gap: _find_interior_point(contour)

KEY CLASSES/FUNCTIONS:
- Polygon2D: Main class wrapping WPF PathGeometry
- Polygon2D(points=list): Create polygon from (x,y) points
- Polygon2D.difference(other): Boolean subtraction
- Polygon2D.get_contours(): Extract polygon boundary as points
- find_all_gap_regions_2d_from_polygons(): Main gap detection function
- _find_interior_point(): Find point inside non-convex polygon
- visualize_2d_geometry(): Debug visualization window
"""

import clr
clr.AddReference('WindowsBase')
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')

import System
from System.Windows import Point as WpfPoint
from System.Windows.Media import (
    PathGeometry, PathFigure, PathFigureCollection,
    LineSegment, PolyLineSegment, ArcSegment,
    GeometryCombineMode, CombinedGeometry, 
    RectangleGeometry, Geometry,
    SweepDirection, FillRule
)
from System.Windows import Rect, Size

# Tolerance for geometry operations
TOLERANCE = 0.001  # feet


class Polygon2D(object):
    """A 2D polygon represented as a WPF PathGeometry.
    
    Supports boolean operations via WPF's CombinedGeometry.
    """
    
    def __init__(self, geometry=None, points=None):
        """Initialize a Polygon2D.
        
        Args:
            geometry: A WPF Geometry object (PathGeometry, CombinedGeometry, etc.)
            points: List of (x, y) tuples defining a closed polygon
        """
        if geometry is not None:
            self._geometry = geometry
        elif points is not None:
            self._geometry = self._create_path_from_points(points)
        else:
            self._geometry = PathGeometry()
    
    @property
    def geometry(self):
        """Get the underlying WPF Geometry object."""
        return self._geometry
    
    @property
    def bounds(self):
        """Get the bounding box as (min_x, min_y, max_x, max_y)."""
        if self._geometry is None:
            return None
        rect = self._geometry.Bounds
        if rect.IsEmpty:
            return None
        return (rect.X, rect.Y, rect.X + rect.Width, rect.Y + rect.Height)
    
    @property
    def is_empty(self):
        """Check if the geometry is empty."""
        if self._geometry is None:
            return True
        return self._geometry.IsEmpty()
    
    def get_area(self):
        """Calculate the area of the polygon."""
        if self._geometry is None:
            return 0.0
        return self._geometry.GetArea()
    
    @staticmethod
    def _create_path_from_points(points):
        """Create a PathGeometry from a list of (x, y) points."""
        if not points or len(points) < 3:
            return PathGeometry()
        
        path_geometry = PathGeometry()
        path_geometry.FillRule = FillRule.EvenOdd
        
        figure = PathFigure()
        figure.StartPoint = WpfPoint(points[0][0], points[0][1])
        figure.IsClosed = True
        figure.IsFilled = True
        
        # Create line segments for remaining points
        segment_points = System.Collections.Generic.List[WpfPoint]()
        for x, y in points[1:]:
            segment_points.Add(WpfPoint(x, y))
        
        poly_segment = PolyLineSegment(segment_points, True)
        figure.Segments.Add(poly_segment)
        
        path_geometry.Figures.Add(figure)
        return path_geometry
    
    @staticmethod
    def _create_path_with_holes(exterior_points, hole_points_list):
        """Create a PathGeometry from exterior points and interior holes.
        
        Uses EvenOdd fill rule - interior figures become holes.
        
        Args:
            exterior_points: List of (x, y) for the outer boundary
            hole_points_list: List of lists of (x, y) for each hole
        
        Returns:
            PathGeometry with holes
        """
        if not exterior_points or len(exterior_points) < 3:
            return PathGeometry()
        
        path_geometry = PathGeometry()
        path_geometry.FillRule = FillRule.EvenOdd  # EvenOdd makes interior figures holes
        
        # Add exterior boundary
        ext_figure = PathFigure()
        ext_figure.StartPoint = WpfPoint(exterior_points[0][0], exterior_points[0][1])
        ext_figure.IsClosed = True
        ext_figure.IsFilled = True
        
        ext_segment_points = System.Collections.Generic.List[WpfPoint]()
        for x, y in exterior_points[1:]:
            ext_segment_points.Add(WpfPoint(x, y))
        ext_figure.Segments.Add(PolyLineSegment(ext_segment_points, True))
        path_geometry.Figures.Add(ext_figure)
        
        # Add each hole as another figure (EvenOdd rule makes it subtract)
        if hole_points_list:
            for hole_points in hole_points_list:
                if hole_points and len(hole_points) >= 3:
                    hole_figure = PathFigure()
                    hole_figure.StartPoint = WpfPoint(hole_points[0][0], hole_points[0][1])
                    hole_figure.IsClosed = True
                    hole_figure.IsFilled = True
                    
                    hole_segment_points = System.Collections.Generic.List[WpfPoint]()
                    for x, y in hole_points[1:]:
                        hole_segment_points.Add(WpfPoint(x, y))
                    hole_figure.Segments.Add(PolyLineSegment(hole_segment_points, True))
                    path_geometry.Figures.Add(hole_figure)
        
        return path_geometry
    
    @classmethod
    def from_points_with_holes(cls, exterior_points, hole_points_list=None):
        """Create a Polygon2D with holes.
        
        Args:
            exterior_points: List of (x, y) for the outer boundary
            hole_points_list: List of lists of (x, y) for each hole (optional)
        
        Returns:
            Polygon2D instance
        """
        if not exterior_points:
            return cls()
        
        if hole_points_list:
            geometry = cls._create_path_with_holes(exterior_points, hole_points_list)
        else:
            geometry = cls._create_path_from_points(exterior_points)
        
        return cls(geometry=geometry)
    
    @classmethod
    def from_curveloop(cls, curve_loop, tessellation_tolerance=0.1):
        """Create a Polygon2D from a Revit CurveLoop.
        
        Args:
            curve_loop: Revit DB.CurveLoop object
            tessellation_tolerance: Max distance for arc tessellation (feet)
        
        Returns:
            Polygon2D instance
        """
        from Autodesk.Revit import DB
        
        points = []
        for curve in curve_loop:
            # Get curve endpoints and tessellation
            if isinstance(curve, DB.Line):
                # Lines: just use start point
                start = curve.GetEndPoint(0)
                points.append((start.X, start.Y))
            else:
                # Arcs, splines, etc: tessellate
                tessellated = curve.Tessellate()
                for i, pt in enumerate(tessellated):
                    if i == len(tessellated) - 1:
                        continue  # Skip last point (will be start of next curve)
                    points.append((pt.X, pt.Y))
        
        if not points:
            return cls()
        
        return cls(points=points)
    
    @classmethod
    def from_boundary_segments(cls, boundary_segments):
        """Create a Polygon2D from Revit boundary segments.
        
        Args:
            boundary_segments: List of BoundarySegment objects from Area.GetBoundarySegments()
        
        Returns:
            Polygon2D instance
        """
        from Autodesk.Revit import DB
        
        points = []
        for segment in boundary_segments:
            curve = segment.GetCurve()
            if curve is None:
                continue
            
            if isinstance(curve, DB.Line):
                start = curve.GetEndPoint(0)
                points.append((start.X, start.Y))
            else:
                tessellated = curve.Tessellate()
                for i, pt in enumerate(tessellated):
                    if i == len(tessellated) - 1:
                        continue
                    points.append((pt.X, pt.Y))
        
        if not points:
            return cls()
        
        return cls(points=points)
    
    def union(self, other):
        """Create a union of this polygon with another.
        
        Args:
            other: Another Polygon2D instance
        
        Returns:
            New Polygon2D representing the union
        """
        if self.is_empty:
            return other
        if other.is_empty:
            return self
        
        combined = CombinedGeometry(
            GeometryCombineMode.Union,
            self._geometry,
            other._geometry
        )
        
        # Flatten to PathGeometry for better performance in subsequent operations
        flattened = combined.GetFlattenedPathGeometry()
        return Polygon2D(geometry=flattened)
    
    def intersection(self, other):
        """Create an intersection of this polygon with another.
        
        Args:
            other: Another Polygon2D instance
        
        Returns:
            New Polygon2D representing the intersection
        """
        if self.is_empty or other.is_empty:
            return Polygon2D()
        
        combined = CombinedGeometry(
            GeometryCombineMode.Intersect,
            self._geometry,
            other._geometry
        )
        
        flattened = combined.GetFlattenedPathGeometry()
        return Polygon2D(geometry=flattened)
    
    def difference(self, other):
        """Subtract another polygon from this one.
        
        Args:
            other: Polygon2D to subtract
        
        Returns:
            New Polygon2D representing self - other
        """
        if self.is_empty:
            return Polygon2D()
        if other.is_empty:
            return self
        
        combined = CombinedGeometry(
            GeometryCombineMode.Exclude,
            self._geometry,
            other._geometry
        )
        
        flattened = combined.GetFlattenedPathGeometry()
        return Polygon2D(geometry=flattened)
    
    def xor(self, other):
        """Exclusive or of this polygon with another.
        
        Args:
            other: Another Polygon2D instance
        
        Returns:
            New Polygon2D representing symmetric difference
        """
        if self.is_empty:
            return other
        if other.is_empty:
            return self
        
        combined = CombinedGeometry(
            GeometryCombineMode.Xor,
            self._geometry,
            other._geometry
        )
        
        flattened = combined.GetFlattenedPathGeometry()
        return Polygon2D(geometry=flattened)
    
    def contains_point(self, x, y):
        """Check if a point is inside the polygon.
        
        Args:
            x, y: Point coordinates
        
        Returns:
            True if point is inside, False otherwise
        """
        if self.is_empty:
            return False
        
        return self._geometry.FillContains(WpfPoint(x, y))
    
    def get_contours(self):
        """Extract all contours from the geometry.
        
        Returns:
            List of contours, where each contour is a list of (x, y) points.
            The first contour is typically the exterior, subsequent ones are holes.
        """
        contours = []
        
        if self.is_empty:
            return contours
        
        # Get as PathGeometry to access figures
        path_geo = self._geometry
        if not isinstance(path_geo, PathGeometry):
            path_geo = self._geometry.GetFlattenedPathGeometry()
        
        for figure in path_geo.Figures:
            points = [(figure.StartPoint.X, figure.StartPoint.Y)]
            
            for segment in figure.Segments:
                if isinstance(segment, LineSegment):
                    points.append((segment.Point.X, segment.Point.Y))
                elif isinstance(segment, PolyLineSegment):
                    for pt in segment.Points:
                        points.append((pt.X, pt.Y))
                # Other segment types would need additional handling
            
            # Remove duplicate closing point if present
            if len(points) > 1 and points[0] == points[-1]:
                points = points[:-1]
            
            if len(points) >= 3:
                contours.append(points)
        
        return contours
    
    def get_centroids(self):
        """Get centroids of all contours.
        
        Returns:
            List of (x, y) centroids, one per contour
        """
        centroids = []
        for contour in self.get_contours():
            if not contour:
                continue
            cx = sum(p[0] for p in contour) / len(contour)
            cy = sum(p[1] for p in contour) / len(contour)
            centroids.append((cx, cy))
        return centroids
    
    @staticmethod
    def union_all(polygons):
        """Create a union of multiple polygons.
        
        Args:
            polygons: List of Polygon2D instances
        
        Returns:
            New Polygon2D representing the union of all
        """
        if not polygons:
            return Polygon2D()
        
        # Filter out empty polygons
        valid_polygons = [p for p in polygons if not p.is_empty]
        if not valid_polygons:
            return Polygon2D()
        
        # Iteratively union all polygons
        result = valid_polygons[0]
        for p in valid_polygons[1:]:
            result = result.union(p)
        
        return result
    
    @staticmethod
    def create_rectangle(min_x, min_y, max_x, max_y):
        """Create a rectangular polygon.
        
        Args:
            min_x, min_y: Bottom-left corner
            max_x, max_y: Top-right corner
        
        Returns:
            Polygon2D rectangle
        """
        rect_geo = RectangleGeometry(
            Rect(WpfPoint(min_x, min_y), WpfPoint(max_x, max_y))
        )
        return Polygon2D(geometry=rect_geo)
    
    @staticmethod
    def find_holes(polygons, margin=1.0):
        """Find holes/gaps between polygons.
        
        This creates a bounding rectangle around all polygons,
        then subtracts the union of all polygons from it.
        The remaining regions are the holes/gaps.
        
        Args:
            polygons: List of Polygon2D instances
            margin: Extra margin around bounding box (feet)
        
        Returns:
            Polygon2D representing the holes (may contain multiple contours)
        """
        if not polygons:
            return Polygon2D()
        
        # Get union of all polygons
        union = Polygon2D.union_all(polygons)
        if union.is_empty:
            return Polygon2D()
        
        # Get bounding box with margin
        bounds = union.bounds
        if bounds is None:
            return Polygon2D()
        
        min_x, min_y, max_x, max_y = bounds
        min_x -= margin
        min_y -= margin
        max_x += margin
        max_y += margin
        
        # Create bounding rectangle
        bounding_rect = Polygon2D.create_rectangle(min_x, min_y, max_x, max_y)
        
        # Subtract union from rectangle to get holes
        holes = bounding_rect.difference(union)
        
        return holes
    
    @staticmethod
    def find_gaps_between_polygons(polygons, outer_boundary=None, margin=1.0):
        """Find gaps between polygons, excluding the outer margin.
        
        This is similar to find_holes but filters out regions that are
        part of the outer boundary margin (not true interior gaps).
        
        Args:
            polygons: List of Polygon2D instances
            outer_boundary: Optional Polygon2D defining the valid region.
                           If None, uses convex hull bounding box with margin.
            margin: Margin for bounding box (if no outer_boundary provided)
        
        Returns:
            Polygon2D representing interior gaps only
        """
        if not polygons:
            return Polygon2D()
        
        # Get union of all polygons
        union = Polygon2D.union_all(polygons)
        if union.is_empty:
            return Polygon2D()
        
        if outer_boundary is not None and not outer_boundary.is_empty:
            # Use provided outer boundary
            gaps = outer_boundary.difference(union)
        else:
            # Use bounding box with margin
            bounds = union.bounds
            if bounds is None:
                return Polygon2D()
            
            min_x, min_y, max_x, max_y = bounds
            # Use a small margin to capture edge gaps without adding extra regions
            min_x -= margin
            min_y -= margin
            max_x += margin
            max_y += margin
            
            bounding_rect = Polygon2D.create_rectangle(min_x, min_y, max_x, max_y)
            gaps = bounding_rect.difference(union)
        
        return gaps
    
    def get_interior_contours(self, area_threshold=1.0):
        """Get interior contours (holes) that meet area threshold.
        
        Args:
            area_threshold: Minimum area in square feet for a hole to be returned
        
        Returns:
            List of contours (each contour is a list of (x, y) points)
        """
        contours = self.get_contours()
        
        if len(contours) <= 1:
            return []  # No interior holes
        
        # Calculate areas and find the largest (exterior)
        contour_areas = []
        for contour in contours:
            area = self._calculate_contour_area(contour)
            contour_areas.append((contour, abs(area)))
        
        # Sort by area descending - largest is exterior
        contour_areas.sort(key=lambda x: x[1], reverse=True)
        
        # Return all but the largest, filtered by threshold
        interior = [c for c, a in contour_areas[1:] if a >= area_threshold]
        return interior
    
    @staticmethod
    def _calculate_contour_area(contour):
        """Calculate signed area of a contour using shoelace formula."""
        if len(contour) < 3:
            return 0.0
        
        area = 0.0
        n = len(contour)
        for i in range(n):
            j = (i + 1) % n
            area += contour[i][0] * contour[j][1]
            area -= contour[j][0] * contour[i][1]
        
        return area / 2.0


def visualize_2d_geometry(polygons, gap_contours, gap_centroids, title="2D Geometry Debug"):
    """Show a WPF window visualizing the 2D geometry.
    
    Args:
        polygons: List of Polygon2D objects (the areas)
        gap_contours: List of contours (each contour is list of (x,y) points)
        gap_centroids: List of (x, y) centroids
        title: Window title
    """
    import System
    from System.Windows import Window, WindowStartupLocation, ResizeMode
    from System.Windows.Controls import Canvas, TextBlock
    from System.Windows.Shapes import Polygon as WpfPolygon, Ellipse, Polyline
    from System.Windows.Media import Brushes, SolidColorBrush, Color, PointCollection
    from System.Windows import Thickness
    
    # Calculate bounds
    all_points = []
    for poly in polygons:
        for contour in poly.get_contours():
            all_points.extend(contour)
    for contour in gap_contours:
        all_points.extend(contour)
    
    if not all_points:
        print("  [2D VIZ] No points to visualize")
        return
    
    min_x = min(p[0] for p in all_points)
    max_x = max(p[0] for p in all_points)
    min_y = min(p[1] for p in all_points)
    max_y = max(p[1] for p in all_points)
    
    # Window size
    width = 800
    height = 600
    margin = 50
    
    # Scale to fit
    data_width = max_x - min_x
    data_height = max_y - min_y
    if data_width == 0:
        data_width = 1
    if data_height == 0:
        data_height = 1
    
    scale_x = (width - 2 * margin) / data_width
    scale_y = (height - 2 * margin) / data_height
    scale = min(scale_x, scale_y)
    
    def transform(x, y):
        """Transform data coordinates to screen coordinates."""
        sx = margin + (x - min_x) * scale
        sy = height - margin - (y - min_y) * scale  # Flip Y
        return sx, sy
    
    # Create window
    window = Window()
    window.Title = title
    window.Width = width
    window.Height = height + 50  # Extra for title
    window.WindowStartupLocation = WindowStartupLocation.CenterScreen
    window.ResizeMode = ResizeMode.CanResize
    
    canvas = Canvas()
    canvas.Background = Brushes.White
    window.Content = canvas
    
    # Draw area polygons (blue, semi-transparent)
    area_brush = SolidColorBrush(Color.FromArgb(100, 0, 100, 255))
    area_stroke = SolidColorBrush(Color.FromArgb(255, 0, 50, 150))
    
    for poly in polygons:
        for contour in poly.get_contours():
            if len(contour) < 3:
                continue
            wpf_poly = WpfPolygon()
            wpf_poly.Fill = area_brush
            wpf_poly.Stroke = area_stroke
            wpf_poly.StrokeThickness = 1
            
            points = PointCollection()
            for x, y in contour:
                sx, sy = transform(x, y)
                points.Add(WpfPoint(sx, sy))
            wpf_poly.Points = points
            canvas.Children.Add(wpf_poly)
    
    # Draw gap contours (red)
    gap_brush = SolidColorBrush(Color.FromArgb(150, 255, 0, 0))
    gap_stroke = SolidColorBrush(Color.FromArgb(255, 200, 0, 0))
    
    for contour in gap_contours:
        if len(contour) < 3:
            continue
        wpf_poly = WpfPolygon()
        wpf_poly.Fill = gap_brush
        wpf_poly.Stroke = gap_stroke
        wpf_poly.StrokeThickness = 2
        
        points = PointCollection()
        for x, y in contour:
            sx, sy = transform(x, y)
            points.Add(WpfPoint(sx, sy))
        wpf_poly.Points = points
        canvas.Children.Add(wpf_poly)
    
    # Draw centroids (green circles)
    for cx, cy in gap_centroids:
        sx, sy = transform(cx, cy)
        ellipse = Ellipse()
        ellipse.Width = 10
        ellipse.Height = 10
        ellipse.Fill = Brushes.Green
        ellipse.Stroke = Brushes.DarkGreen
        ellipse.StrokeThickness = 2
        Canvas.SetLeft(ellipse, sx - 5)
        Canvas.SetTop(ellipse, sy - 5)
        canvas.Children.Add(ellipse)
        
        # Add coordinate label
        label = TextBlock()
        label.Text = "({:.1f}, {:.1f})".format(cx, cy)
        label.FontSize = 9
        label.Foreground = Brushes.DarkGreen
        Canvas.SetLeft(label, sx + 8)
        Canvas.SetTop(label, sy - 5)
        canvas.Children.Add(label)
    
    # Add legend
    legend = TextBlock()
    legend.Text = "Blue=Areas, Red=Gaps, Green=Centroids | Polygons:{} Gaps:{} Centroids:{}".format(
        len(polygons), len(gap_contours), len(gap_centroids))
    legend.FontSize = 12
    legend.Foreground = Brushes.Black
    Canvas.SetLeft(legend, 10)
    Canvas.SetTop(legend, 10)
    canvas.Children.Add(legend)
    
    # Show window
    window.ShowDialog()


def visualize_2d_geometry_zoomable(polygons, gap_contours, gap_centroids, title="2D Geometry Debug"):
    """Show a non-modal zoomable window visualizing 2D geometry.
    
    Features:
    - Mouse wheel to zoom in/out (no scrollbar interference)
    - Left-click and drag to pan
    - Fixed pixel-width strokes (LayoutTransform doesn't scale strokes)
    - Non-modal (stays open, doesn't block Revit)
    
    Args:
        polygons: List of Polygon2D objects (the areas)
        gap_contours: List of contours (each contour is list of (x,y) points)
        gap_centroids: List of (x, y) centroids
        title: Window title
    """
    import System
    from System.Windows import Window, WindowStartupLocation, ResizeMode
    from System.Windows.Controls import Canvas, TextBlock, Border
    from System.Windows.Shapes import Polygon as WpfPolygon, Ellipse
    from System.Windows.Media import Brushes, SolidColorBrush, Color, PointCollection, ScaleTransform, TranslateTransform, TransformGroup
    from System.Windows import Thickness
    from System.Windows.Input import MouseButtonState
    
    # Calculate bounds
    all_points = []
    for poly in polygons:
        for contour in poly.get_contours():
            all_points.extend(contour)
    for contour in gap_contours:
        all_points.extend(contour)
    
    if not all_points:
        print("  [2D VIZ] No points to visualize")
        return
    
    min_x = min(p[0] for p in all_points)
    max_x = max(p[0] for p in all_points)
    min_y = min(p[1] for p in all_points)
    max_y = max(p[1] for p in all_points)
    
    # Data dimensions
    data_width = max_x - min_x
    data_height = max_y - min_y
    if data_width == 0:
        data_width = 1
    if data_height == 0:
        data_height = 1
    
    # Window and canvas size - fixed, no resize
    window_width = 900
    window_height = 700
    margin = 50
    
    # Initial scale to fit data in window
    scale_x = (window_width - 2 * margin) / data_width
    scale_y = (window_height - 2 * margin) / data_height
    initial_scale = min(scale_x, scale_y)
    
    def transform(x, y, scale=initial_scale):
        """Transform data coordinates to canvas coordinates."""
        sx = margin + (x - min_x) * scale
        sy = window_height - margin - (y - min_y) * scale  # Flip Y
        return sx, sy
    
    # Create window - fixed size, no resize
    window = Window()
    window.Title = title
    window.Width = window_width
    window.Height = window_height
    window.WindowStartupLocation = WindowStartupLocation.CenterScreen
    window.ResizeMode = ResizeMode.NoResize  # Fixed size
    
    # Grid to layer canvas and legend
    from System.Windows.Controls import Grid
    grid = Grid()
    grid.Width = window_width
    grid.Height = window_height
    
    # Canvas - same size as window
    canvas = Canvas()
    canvas.Width = window_width
    canvas.Height = window_height
    canvas.Background = Brushes.White
    canvas.ClipToBounds = True
    
    # Transform group for zoom/pan
    transform_group = TransformGroup()
    scale_transform = ScaleTransform(1.0, 1.0)
    translate_transform = TranslateTransform(0, 0)
    transform_group.Children.Add(scale_transform)
    transform_group.Children.Add(translate_transform)
    canvas.RenderTransform = transform_group
    
    # Add canvas to grid
    grid.Children.Add(canvas)
    
    # Store shapes for inverse stroke scaling - use dict since WPF objects don't allow custom attributes
    shape_thickness_map = {}
    centroid_elements = []  # Store (element, base_size, cx, cy) for centroids
    
    window.Content = grid
    
    # Mouse interaction state
    class MouseState:
        def __init__(self):
            self.is_dragging = False
            self.last_pos = None
    
    mouse_state = MouseState()
    
    # Helper to update stroke thickness (inverse scale)
    def update_stroke_thickness():
        """Apply inverse scale to strokes and centroids to keep them constant size."""
        try:
            current_scale = scale_transform.ScaleX
            inverse = 1.0 / current_scale if current_scale > 0 else 1.0
            
            # Update polygon strokes
            for shape, base_thickness in shape_thickness_map.items():
                shape.StrokeThickness = base_thickness * inverse
            
            # Update centroid circles and labels
            for element, base_size, cx, cy in centroid_elements:
                if isinstance(element, Ellipse):
                    # Scale circle size
                    new_size = base_size * inverse
                    element.Width = new_size
                    element.Height = new_size
                    element.StrokeThickness = 2 * inverse
                    # Recalculate position
                    sx, sy = transform(cx, cy)
                    Canvas.SetLeft(element, sx - new_size / 2.0)
                    Canvas.SetTop(element, sy - new_size / 2.0)
                elif isinstance(element, TextBlock):
                    # Scale font size
                    element.FontSize = base_size * inverse
                    # Recalculate position
                    sx, sy = transform(cx, cy)
                    Canvas.SetLeft(element, sx + 6 * inverse)
                    Canvas.SetTop(element, sy - 5 * inverse)
        except Exception:
            pass
    
    # Helper to constrain pan within bounds
    def constrain_pan():
        """Keep the canvas content visible - don't pan off screen."""
        try:
            current_scale = scale_transform.ScaleX
            
            # Calculate how much the scaled canvas exceeds the window
            scaled_width = window_width * current_scale
            scaled_height = window_height * current_scale
            
            # Calculate pan limits to keep content on screen
            # Allow panning the full extent of the scaled canvas
            max_pan_x = max(0, (scaled_width - window_width))
            max_pan_y = max(0, (scaled_height - window_height))
            
            # Constrain translation (can pan from -max to 0)
            # Negative translation moves content left/up (shows right/bottom)
            # Positive translation moves content right/down (shows left/top)
            translate_transform.X = max(-max_pan_x, min(0, translate_transform.X))
            translate_transform.Y = max(-max_pan_y, min(0, translate_transform.Y))
        except Exception:
            pass
    
    # Mouse wheel zoom
    def on_mouse_wheel(sender, e):
        try:
            delta = e.Delta
            zoom_factor = 1.15 if delta > 0 else 0.87  # Smoother zoom steps
            
            # Get mouse position relative to canvas
            pos = e.GetPosition(canvas)
            
            # Calculate new scale - min 1.0 (entire canvas visible), max 20.0
            new_scale = scale_transform.ScaleX * zoom_factor
            new_scale = max(1.0, min(new_scale, 20.0))  # Limit: 1.0x to 20.0x
            
            # Zoom towards mouse position
            old_scale = scale_transform.ScaleX
            scale_transform.ScaleX = new_scale
            scale_transform.ScaleY = new_scale
            
            # Adjust translation to zoom towards mouse
            scale_change = new_scale / old_scale
            translate_transform.X = pos.X - (pos.X - translate_transform.X) * scale_change
            translate_transform.Y = pos.Y - (pos.Y - translate_transform.Y) * scale_change
            
            # Constrain pan to keep content visible
            constrain_pan()
            
            # Update stroke thickness for new zoom level
            update_stroke_thickness()
            
            e.Handled = True
        except Exception:
            pass
    
    # Mouse drag to pan
    def on_mouse_down(sender, e):
        if e.LeftButton == MouseButtonState.Pressed:
            mouse_state.is_dragging = True
            # Get position relative to WINDOW, not transformed canvas
            mouse_state.last_pos = e.GetPosition(window)
            canvas.CaptureMouse()
            e.Handled = True
    
    def on_mouse_up(sender, e):
        if mouse_state.is_dragging:
            mouse_state.is_dragging = False
            canvas.ReleaseMouseCapture()
            e.Handled = True
    
    def on_mouse_move(sender, e):
        try:
            if mouse_state.is_dragging and mouse_state.last_pos:
                # Get position relative to WINDOW, not transformed canvas
                current_pos = e.GetPosition(window)
                delta_x = current_pos.X - mouse_state.last_pos.X
                delta_y = current_pos.Y - mouse_state.last_pos.Y
                
                translate_transform.X += delta_x
                translate_transform.Y += delta_y
                
                # Constrain pan to keep content visible
                constrain_pan()
                
                mouse_state.last_pos = current_pos
                e.Handled = True
        except Exception:
            pass
    
    # Attach events to canvas
    canvas.MouseWheel += on_mouse_wheel
    canvas.MouseLeftButtonDown += on_mouse_down
    canvas.MouseLeftButtonUp += on_mouse_up
    canvas.MouseMove += on_mouse_move
    
    # Draw area polygons (blue, semi-transparent)
    area_brush = SolidColorBrush(Color.FromArgb(100, 0, 100, 255))
    area_stroke = SolidColorBrush(Color.FromArgb(255, 0, 50, 150))
    
    for poly in polygons:
        for contour in poly.get_contours():
            if len(contour) < 3:
                continue
            wpf_poly = WpfPolygon()
            wpf_poly.Fill = area_brush
            wpf_poly.Stroke = area_stroke
            wpf_poly.StrokeThickness = 1
            wpf_poly.UseLayoutRounding = True
            
            points = PointCollection()
            for x, y in contour:
                sx, sy = transform(x, y)
                points.Add(WpfPoint(sx, sy))
            wpf_poly.Points = points
            canvas.Children.Add(wpf_poly)
            shape_thickness_map[wpf_poly] = 1  # Store base thickness in dict
    
    # Draw gap contours (red, highlighted)
    gap_brush = SolidColorBrush(Color.FromArgb(180, 255, 0, 0))
    gap_stroke = SolidColorBrush(Color.FromArgb(255, 200, 0, 0))
    
    for contour in gap_contours:
        if len(contour) < 3:
            continue
        wpf_poly = WpfPolygon()
        wpf_poly.Fill = gap_brush
        wpf_poly.Stroke = gap_stroke
        wpf_poly.StrokeThickness = 2
        wpf_poly.UseLayoutRounding = True
        
        points = PointCollection()
        for x, y in contour:
            sx, sy = transform(x, y)
            points.Add(WpfPoint(sx, sy))
        wpf_poly.Points = points
        canvas.Children.Add(wpf_poly)
        shape_thickness_map[wpf_poly] = 2  # Store base thickness in dict
    
    # Draw centroids (green circles with labels)
    for cx, cy in gap_centroids:
        sx, sy = transform(cx, cy)
        
        # Circle - base 8px diameter
        ellipse = Ellipse()
        ellipse.Width = 8
        ellipse.Height = 8
        ellipse.Fill = Brushes.Green
        ellipse.Stroke = Brushes.DarkGreen
        ellipse.StrokeThickness = 2
        ellipse.UseLayoutRounding = True
        Canvas.SetLeft(ellipse, sx - 4)
        Canvas.SetTop(ellipse, sy - 4)
        canvas.Children.Add(ellipse)
        centroid_elements.append((ellipse, 8, cx, cy))  # Store for inverse scaling
        
        # Coordinate label - base 10pt font
        label = TextBlock()
        label.Text = "({:.1f}, {:.1f})".format(cx, cy)
        label.FontSize = 10
        label.FontWeight = System.Windows.FontWeights.Bold
        label.Foreground = Brushes.DarkGreen
        label.Background = SolidColorBrush(Color.FromArgb(200, 255, 255, 255))
        label.UseLayoutRounding = True
        Canvas.SetLeft(label, sx + 6)
        Canvas.SetTop(label, sy - 5)
        canvas.Children.Add(label)
        centroid_elements.append((label, 10, cx, cy))  # Store for inverse scaling
    
    # Add legend at top - on grid, not canvas, so it stays fixed
    legend = TextBlock()
    legend.Text = "Blue=Areas ({}) | Red=Failed Gaps ({}) | Green=Centroids ({})\\nMouse Wheel=Zoom | Drag=Pan".format(
        len(polygons), len(gap_contours), len(gap_centroids))
    legend.FontSize = 12
    legend.FontWeight = System.Windows.FontWeights.Bold
    legend.Foreground = Brushes.Black
    legend.Background = SolidColorBrush(Color.FromArgb(220, 255, 255, 200))
    legend.Padding = Thickness(8)
    legend.HorizontalAlignment = System.Windows.HorizontalAlignment.Left
    legend.VerticalAlignment = System.Windows.VerticalAlignment.Top
    legend.Margin = Thickness(10, 10, 0, 0)
    grid.Children.Add(legend)  # Add to grid, not canvas
    
    # Show window (non-modal - stays open, doesn't block Revit)
    window.Show()


def find_gap_points_2d(curve_loops, debug=False):
    """Find gap/hole points using 2D polygon boolean operations.
    
    This is the main entry point for hole detection using 2D analysis.
    
    Args:
        curve_loops: List of Revit CurveLoop objects representing area boundaries
        debug: If True, print debug information
    
    Returns:
        List of (x, y, z) tuples representing centroids of detected holes
    """
    if not curve_loops:
        return []
    
    # Convert curve loops to Polygon2D objects
    polygons = []
    for i, cl in enumerate(curve_loops):
        try:
            poly = Polygon2D.from_curveloop(cl)
            if not poly.is_empty:
                polygons.append(poly)
                if debug:
                    print("  [2D] Polygon {}: bounds={}".format(i, poly.bounds))
        except Exception as e:
            if debug:
                print("  [2D] Failed to convert curveloop {}: {}".format(i, e))
            continue
    
    if debug:
        print("  [2D] Converted {} of {} curve loops to polygons".format(
            len(polygons), len(curve_loops)))
    
    if not polygons:
        return []
    
    # Find holes/gaps
    try:
        holes = Polygon2D.find_holes(polygons, margin=0.5)
        
        if holes.is_empty:
            if debug:
                print("  [2D] No holes found in polygon union")
            return []
        
        # Get interior contours (the actual holes, not the outer boundary)
        # Use a small area threshold to filter noise
        interior_contours = holes.get_interior_contours(area_threshold=1.0)
        
        if debug:
            print("  [2D] Found {} interior hole contours".format(len(interior_contours)))
        
        # Get centroids of holes
        gap_points = []
        for contour in interior_contours:
            if contour:
                cx = sum(p[0] for p in contour) / len(contour)
                cy = sum(p[1] for p in contour) / len(contour)
                gap_points.append((cx, cy, 0.0))  # Z=0 for 2D
                if debug:
                    area = abs(Polygon2D._calculate_contour_area(contour))
                    print("    [2D] Hole centroid: ({:.2f}, {:.2f}), area={:.2f} sqft".format(
                        cx, cy, area))
        
        return gap_points
    
    except Exception as e:
        if debug:
            print("  [2D] Error finding holes: {}".format(e))
        return []


def _contour_is_outer_margin(contour, bbox, union_polygon, tolerance=0.1):
    """Check if a contour is an outer margin region (not an interior gap).
    
    A margin region:
    1. Touches the bounding box edge
    2. Its centroid is OUTSIDE the original union of polygons
    
    An interior gap:
    1. May or may not touch the bbox
    2. Its centroid is INSIDE the original union of polygons
    
    Args:
        contour: List of (x, y) points
        bbox: Tuple of (min_x, min_y, max_x, max_y)
        union_polygon: The Polygon2D union of all original areas
        tolerance: Distance from edge to consider "touching"
    
    Returns:
        True if this is an outer margin (should be filtered out)
    """
    min_x, min_y, max_x, max_y = bbox
    
    # Check if contour touches bounding box edge
    touches_bbox = False
    for x, y in contour:
        if (abs(x - min_x) < tolerance or 
            abs(x - max_x) < tolerance or
            abs(y - min_y) < tolerance or 
            abs(y - max_y) < tolerance):
            touches_bbox = True
            break
    
    if not touches_bbox:
        # If it doesn't touch the bbox, it's definitely an interior gap
        return False
    
    # It touches the bbox - check if the centroid is inside the union
    # If the centroid is INSIDE the union, it's a true interior gap
    # If OUTSIDE the union, it's a margin region
    cx = sum(p[0] for p in contour) / len(contour)
    cy = sum(p[1] for p in contour) / len(contour)
    
    # A gap's centroid should be in a "hole" area - OUTSIDE the original polygons union
    # but INSIDE the overall bounds. A margin's centroid would also be outside the union.
    # The key difference: a margin wraps AROUND the union's exterior.
    
    # Better check: count how many bbox edges this contour touches.
    # A margin typically touches multiple edges (wraps around corners).
    # An interior gap near an edge typically touches only ONE edge.
    edges_touched = 0
    touches_left = any(abs(p[0] - min_x) < tolerance for p in contour)
    touches_right = any(abs(p[0] - max_x) < tolerance for p in contour)
    touches_bottom = any(abs(p[1] - min_y) < tolerance for p in contour)
    touches_top = any(abs(p[1] - max_y) < tolerance for p in contour)
    
    edges_touched = sum([touches_left, touches_right, touches_bottom, touches_top])
    
    # If it touches 3+ edges, it's definitely wrapping around (margin)
    if edges_touched >= 3:
        return True
    
    # If it touches only 1 edge and has reasonable size, likely interior gap
    if edges_touched == 1:
        return False
    
    # For 2 edges, check if they're opposite (spans entire dimension = margin)
    # or adjacent (corner = might be margin)
    if edges_touched == 2:
        if (touches_left and touches_right) or (touches_top and touches_bottom):
            # Spans entire dimension - likely margin
            return True
        # Adjacent edges - could be margin or gap, keep it as gap
        return False
    
    return False


def _point_to_segment_distance(px, py, x1, y1, x2, y2):
    """Calculate minimum distance from point (px, py) to line segment (x1,y1)-(x2,y2)."""
    dx = x2 - x1
    dy = y2 - y1
    
    if dx == 0 and dy == 0:
        # Degenerate segment (point)
        return ((px - x1)**2 + (py - y1)**2)**0.5
    
    # Parameter t for closest point on line
    t = max(0, min(1, ((px - x1) * dx + (py - y1) * dy) / (dx * dx + dy * dy)))
    
    # Closest point on segment
    closest_x = x1 + t * dx
    closest_y = y1 + t * dy
    
    return ((px - closest_x)**2 + (py - closest_y)**2)**0.5


def _find_bottleneck_points(contour, threshold=0.5):
    """Find points where boundary is close to another part of itself.
    
    These are bottleneck points where the polygon is very narrow.
    
    Args:
        contour: List of (x, y) polygon vertices
        threshold: Maximum distance to consider a bottleneck (feet)
                   Default 0.5 ft = ~15cm - narrow corridors that won't fit an area
    
    Returns:
        List of indices where bottlenecks occur
    """
    n = len(contour)
    if n < 6:  # Need at least 6 points for meaningful bottleneck
        return []
    
    bottleneck_indices = []
    
    for i in range(n):
        px, py = contour[i]
        is_narrow = False
        
        # Check distance to all non-adjacent edges
        for j in range(n):
            # Skip adjacent edges (within 3 indices to avoid false positives)
            dist_in_ring = min(abs(i - j), n - abs(i - j))
            if dist_in_ring <= 3:
                continue
            
            # Get edge j -> j+1
            x1, y1 = contour[j]
            x2, y2 = contour[(j + 1) % n]
            
            # Distance from point to edge
            dist = _point_to_segment_distance(px, py, x1, y1, x2, y2)
            
            if dist < threshold:
                is_narrow = True
                break
        
        if is_narrow:
            bottleneck_indices.append(i)
    
    return bottleneck_indices


def _split_contour_at_bottlenecks(contour, bottleneck_threshold=0.5, min_region_area=1.0, _depth=0):
    """Split a contour into separate regions at narrow bottlenecks.
    
    Detects where the boundary comes very close to itself (< threshold),
    indicating a narrow corridor. Splits the contour at these bottlenecks
    to create separate regions. Recursively splits to handle multiple bottlenecks.
    
    Args:
        contour: List of (x, y) polygon vertices
        bottleneck_threshold: Distance threshold for bottleneck detection (feet)
                             Default 0.5 ft = ~15cm - narrow corridors won't fit an area
        min_region_area: Minimum area for a valid split region (sqft)
        _depth: Internal recursion depth (prevents infinite recursion)
    
    Returns:
        List of contours (each is a list of (x,y) points)
    """
    MAX_RECURSION = 10  # Prevent infinite recursion
    
    n = len(contour)
    
    if n < 6 or _depth > MAX_RECURSION:
        return [contour]
    
    # Find bottleneck points (where boundary is close to itself)
    bottleneck_indices = _find_bottleneck_points(contour, bottleneck_threshold)
    
    if not bottleneck_indices:
        return [contour]
    
    # Group consecutive bottleneck indices into zones
    zones = []
    current_zone = [bottleneck_indices[0]]
    
    for i in range(1, len(bottleneck_indices)):
        curr = bottleneck_indices[i]
        prev = bottleneck_indices[i - 1]
        
        if curr == prev + 1 or (prev == n - 1 and curr == 0):
            current_zone.append(curr)
        else:
            zones.append(current_zone)
            current_zone = [curr]
    zones.append(current_zone)
    
    # Need at least 2 separate zones to split
    if len(zones) < 2:
        return [contour]
    
    # Find the best pair of zones to split at (most "opposite" = roughly half-contour apart)
    # Each pair of opposing zones represents one bottleneck
    best_pair = None
    best_score = -1
    
    for i in range(len(zones)):
        for j in range(i + 1, len(zones)):
            zone_i = zones[i]
            zone_j = zones[j]
            
            # Calculate how "opposite" these zones are (ideally ~n/2 apart)
            mid_i = (zone_i[0] + zone_i[-1]) / 2.0
            mid_j = (zone_j[0] + zone_j[-1]) / 2.0
            
            # Distance in ring (accounting for wrap-around)
            dist = abs(mid_j - mid_i)
            if dist > n / 2:
                dist = n - dist
            
            # Score: closer to n/2 is better (more opposite = true bottleneck pair)
            # Also prefer larger zones (more confident bottleneck)
            opposition_score = 1.0 - abs(dist - n / 2.0) / (n / 2.0)
            size_score = (len(zone_i) + len(zone_j)) / float(n)
            score = opposition_score + size_score * 0.5
            
            if score > best_score:
                best_score = score
                best_pair = (zone_i, zone_j)
    
    if best_pair is None:
        return [contour]
    
    zone0, zone1 = best_pair
    # Ensure zone0 comes before zone1 in index order
    if zone0[0] > zone1[0]:
        zone0, zone1 = zone1, zone0
    
    # The two zones represent OPPOSITE SIDES of the narrow corridor
    # Cut ACROSS the bottleneck to separate the wide regions
    zone0_end = zone0[-1]
    zone1_start = zone1[0]
    zone1_end = zone1[-1]
    zone0_start = zone0[0]
    
    # Calculate index ranges for each wide region
    if zone0_end < zone1_start:
        region_a_indices = list(range(zone0_end, zone1_start + 1))
    else:
        region_a_indices = list(range(zone0_end, n)) + list(range(0, zone1_start + 1))
    
    if zone1_end < zone0_start:
        region_b_indices = list(range(zone1_end, zone0_start + 1))
    else:
        region_b_indices = list(range(zone1_end, n)) + list(range(0, zone0_start + 1))
    
    # Extract contours
    contour1 = [contour[i] for i in region_a_indices]
    contour2 = [contour[i] for i in region_b_indices]
    
    # Validate and recursively split each resulting contour
    result = []
    for c in [contour1, contour2]:
        if len(c) >= 3:
            area = abs(Polygon2D._calculate_contour_area(c))
            if area >= min_region_area:
                # Recursively check for more bottlenecks
                sub_results = _split_contour_at_bottlenecks(
                    c, bottleneck_threshold, min_region_area, _depth + 1
                )
                result.extend(sub_results)
    
    if len(result) >= 1:
        return result
    
    return [contour]


def _point_in_polygon(x, y, contour):
    """Check if point (x, y) is inside a polygon using ray casting.
    
    Args:
        x, y: Point coordinates
        contour: List of (x, y) polygon vertices
    
    Returns:
        True if point is inside polygon
    """
    n = len(contour)
    inside = False
    
    j = n - 1
    for i in range(n):
        xi, yi = contour[i]
        xj, yj = contour[j]
        
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi) + xi):
            inside = not inside
        j = i
    
    return inside


def _find_interior_point(contour, debug=False, gap_polygon=None):
    """Find a point guaranteed to be inside a polygon, preferring the widest area.
    
    For non-convex polygons, the centroid may fall outside.
    This function tries multiple strategies to find an interior point,
    preferring points in wider areas (more clearance from boundary).
    
    Args:
        contour: List of (x, y) polygon vertices
        debug: Print debug info
        gap_polygon: Optional Polygon2D representing the actual gap geometry.
                    Used to validate points for donut-shaped gaps where the
                    contour alone doesn't capture interior holes.
    
    Returns:
        (x, y) point inside the polygon, or centroid as fallback
    """
    if len(contour) < 3:
        return None
    
    def is_valid_point(x, y):
        """Check if point is valid - inside contour AND inside gap_polygon if provided."""
        if not _point_in_polygon(x, y, contour):
            return False
        # For donut-shaped gaps, also check against the actual gap geometry
        if gap_polygon is not None:
            return gap_polygon.contains_point(x, y)
        return True
    
    # Calculate centroid
    cx = sum(p[0] for p in contour) / len(contour)
    cy = sum(p[1] for p in contour) / len(contour)
    
    # Get bounding box of contour
    min_x = min(p[0] for p in contour)
    max_x = max(p[0] for p in contour)
    min_y = min(p[1] for p in contour)
    max_y = max(p[1] for p in contour)
    
    # Collect ALL valid interior points with their clearance (segment width)
    # We want to pick the point in the WIDEST part of the polygon
    candidate_points = []  # List of (x, y, clearance)
    
    # Check centroid first
    if is_valid_point(cx, cy):
        # Estimate clearance at centroid by checking distance to boundary
        centroid_clearance = _estimate_clearance(cx, cy, contour)
        candidate_points.append((cx, cy, centroid_clearance))
    
    # Scan at multiple Y levels to find wide interior segments
    for y_ratio in [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
        y = min_y + (max_y - min_y) * y_ratio
        
        # Find all X intersections with the polygon edges at this Y
        intersections = []
        n = len(contour)
        for i in range(n):
            x1, y1 = contour[i]
            x2, y2 = contour[(i + 1) % n]
            
            # Check if edge crosses this Y level
            if (y1 <= y < y2) or (y2 <= y < y1):
                if y2 != y1:
                    x_intersect = x1 + (y - y1) * (x2 - x1) / (y2 - y1)
                    intersections.append(x_intersect)
        
        if len(intersections) >= 2:
            intersections.sort()
            # Check each pair of intersections (inside segments)
            for i in range(0, len(intersections) - 1, 2):
                x_left = intersections[i]
                x_right = intersections[i + 1]
                segment_width = x_right - x_left
                mid_x = (x_left + x_right) / 2
                
                if is_valid_point(mid_x, y):
                    # Use segment width as clearance estimate
                    candidate_points.append((mid_x, y, segment_width))
    
    # Also scan at multiple X levels (vertical scan lines) for irregular shapes
    for x_ratio in [0.1, 0.2, 0.3, 0.4, 0.5, 0.6, 0.7, 0.8, 0.9]:
        x = min_x + (max_x - min_x) * x_ratio
        
        # Find all Y intersections
        intersections = []
        n = len(contour)
        for i in range(n):
            x1, y1 = contour[i]
            x2, y2 = contour[(i + 1) % n]
            
            if (x1 <= x < x2) or (x2 <= x < x1):
                if x2 != x1:
                    y_intersect = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
                    intersections.append(y_intersect)
        
        if len(intersections) >= 2:
            intersections.sort()
            for i in range(0, len(intersections) - 1, 2):
                y_bottom = intersections[i]
                y_top = intersections[i + 1]
                segment_height = y_top - y_bottom
                mid_y = (y_bottom + y_top) / 2
                
                if is_valid_point(x, mid_y):
                    candidate_points.append((x, mid_y, segment_height))
    
    # Pick the point with maximum clearance (widest area)
    if candidate_points:
        candidate_points.sort(key=lambda p: p[2], reverse=True)
        best_x, best_y, best_clearance = candidate_points[0]
        if debug:
            print("    [2D] Found interior point in widest area: ({:.2f}, {:.2f}) clearance={:.2f}".format(
                best_x, best_y, best_clearance))
        return (best_x, best_y)
    
    # Last resort: try midpoints of edges offset inward
    for i in range(len(contour)):
        x1, y1 = contour[i]
        x2, y2 = contour[(i + 1) % len(contour)]
        mid_x = (x1 + x2) / 2
        mid_y = (y1 + y2) / 2
        # Offset slightly inward (toward centroid)
        offset_x = mid_x + (cx - mid_x) * 0.1
        offset_y = mid_y + (cy - mid_y) * 0.1
        if is_valid_point(offset_x, offset_y):
            if debug:
                print("    [2D] Found interior point near edge: ({:.2f}, {:.2f})".format(offset_x, offset_y))
            return (offset_x, offset_y)
    
    if debug:
        print("    [2D] Warning: Could not find interior point, using centroid as fallback")
    
    # Fallback to centroid even though it's outside
    return (cx, cy)


def _estimate_clearance(x, y, contour):
    """Estimate clearance (distance to nearest boundary) at a point."""
    min_dist = float('inf')
    n = len(contour)
    for i in range(n):
        x1, y1 = contour[i]
        x2, y2 = contour[(i + 1) % n]
        dist = _point_to_segment_distance(x, y, x1, y1, x2, y2)
        if dist < min_dist:
            min_dist = dist
    return min_dist


def _compute_convex_hull(points):
    """Compute convex hull of a set of points using Graham scan.
    
    Args:
        points: List of (x, y) tuples
    
    Returns:
        List of (x, y) points forming the convex hull (counterclockwise)
    """
    if len(points) < 3:
        return list(points)
    
    # Find the bottom-most point (and leftmost if tie)
    start = min(points, key=lambda p: (p[1], p[0]))
    
    def cross(o, a, b):
        """Cross product of vectors OA and OB."""
        return (a[0] - o[0]) * (b[1] - o[1]) - (a[1] - o[1]) * (b[0] - o[0])
    
    def dist_sq(a, b):
        """Squared distance between two points."""
        return (a[0] - b[0])**2 + (a[1] - b[1])**2
    
    # Sort points by polar angle with respect to start
    import math
    def angle_key(p):
        if p == start:
            return (-float('inf'), 0)
        angle = math.atan2(p[1] - start[1], p[0] - start[0])
        return (angle, dist_sq(start, p))
    
    sorted_points = sorted(points, key=angle_key)
    
    # Build hull
    hull = []
    for p in sorted_points:
        while len(hull) >= 2 and cross(hull[-2], hull[-1], p) <= 0:
            hull.pop()
        hull.append(p)
    
    return hull


def find_all_gap_regions_2d(curve_loops, debug=False):
    """Find all gap regions with their contours and centroids.
    
    Algorithm:
    1. Convert all area boundaries to polygons
    2. Union all polygons together
    3. Get the outer boundary of the union (the actual perimeter, not convex hull)
    4. Subtract the union from its own outer boundary filled
    5. Also check for interior holes within the union
    
    This handles both:
    - Gaps between areas (enclosed by the overall boundary)
    - Interior holes within individual areas (donut holes)
    
    Args:
        curve_loops: List of Revit CurveLoop objects
        debug: If True, print debug information
    
    Returns:
        List of dicts with 'contour', 'centroid', 'area' keys
    """
    if not curve_loops:
        return []
    
    # Convert curve loops to Polygon2D objects
    polygons = []
    failed_conversions = 0
    empty_polygons = 0
    
    for i, cl in enumerate(curve_loops):
        try:
            poly = Polygon2D.from_curveloop(cl)
            if not poly.is_empty:
                polygons.append(poly)
                # Debug: check polygon validity
                contours = poly.get_contours()
                if debug and (not contours or len(contours[0]) < 3):
                    print("  [2D] WARNING: Polygon {} has invalid contours: {}".format(i, len(contours) if contours else 0))
            else:
                empty_polygons += 1
                if debug:
                    print("  [2D] CurveLoop {} converted to empty polygon".format(i))
        except Exception as e:
            failed_conversions += 1
            if debug:
                print("  [2D] CurveLoop {} failed to convert: {}".format(i, e))
            continue
    
    if not polygons:
        if debug:
            print("  [2D] No polygons created! Failed:{} Empty:{}".format(failed_conversions, empty_polygons))
        return []
    
    if debug:
        print("  [2D] Converted {} curve loops to polygons (failed:{}, empty:{})".format(
            len(polygons), failed_conversions, empty_polygons))
    
    try:
        # APPROACH: Subtract each polygon from a bounding rectangle
        # This is more reliable than union because:
        # 1. Union can merge adjacent areas and lose gaps
        # 2. Individual subtraction preserves all non-covered regions
        
        # Get bounding box of all polygons
        all_min_x, all_min_y = float('inf'), float('inf')
        all_max_x, all_max_y = float('-inf'), float('-inf')
        
        for poly in polygons:
            bounds = poly.bounds
            if bounds:
                min_x, min_y, max_x, max_y = bounds
                all_min_x = min(all_min_x, min_x)
                all_min_y = min(all_min_y, min_y)
                all_max_x = max(all_max_x, max_x)
                all_max_y = max(all_max_y, max_y)
        
        if all_min_x == float('inf'):
            return []
        
        # Add small margin
        margin = 1.0
        all_min_x -= margin
        all_min_y -= margin
        all_max_x += margin
        all_max_y += margin
        
        if debug:
            print("  [2D] Bounding box: ({:.1f}, {:.1f}) to ({:.1f}, {:.1f})".format(
                all_min_x, all_min_y, all_max_x, all_max_y))
        
        # Create bounding rectangle
        bbox_rect = Polygon2D.create_rectangle(all_min_x, all_min_y, all_max_x, all_max_y)
        
        # Subtract each polygon individually from the bounding rectangle
        # This preserves gaps that would be lost in a union
        result = bbox_rect
        for i, poly in enumerate(polygons):
            result = result.difference(poly)
            if result.is_empty:
                if debug:
                    print("  [2D] Warning: result became empty after subtracting polygon {}".format(i))
                break
        
        if result.is_empty:
            if debug:
                print("  [2D] No gaps found (areas cover entire bounding box)")
            return []
        
        # Get all contours from the result (gaps)
        all_contours_raw = result.get_contours()
        
        if debug:
            print("  [2D] Found {} raw contours after subtraction".format(len(all_contours_raw)))
        
        # Filter out the outer bounding box contour (the margin around everything)
        # It will be the largest contour
        contour_data = []
        for contour in all_contours_raw:
            if len(contour) < 3:
                continue
            area = abs(Polygon2D._calculate_contour_area(contour))
            contour_data.append({
                'contour': contour,
                'area': area
            })
        
        if not contour_data:
            return []
        
        # Sort by area descending
        contour_data.sort(key=lambda x: x['area'], reverse=True)
        
        # The largest contour is the outer margin (bbox - all areas) - skip it
        # All others are interior gaps
        if len(contour_data) > 1:
            outer_margin_area = contour_data[0]['area']
            if debug:
                print("  [2D] Outer margin area: {:.2f} sqft (filtering out)".format(outer_margin_area))
            all_contours = [cd['contour'] for cd in contour_data[1:]]
        else:
            # Only one contour - it might be all gaps connected
            all_contours = [contour_data[0]['contour']]
        
        if debug:
            print("  [2D] Found {} interior gap contours".format(len(all_contours)))
        
        # Process each contour - all regions are valid interior gaps
        regions = []
        for contour in all_contours:
            if len(contour) < 3:
                continue
            
            area = abs(Polygon2D._calculate_contour_area(contour))
            
            # Skip very small regions (noise)
            if area < 1.0:  # 1 sqft threshold (increased slightly)
                if debug:
                    print("  [2D] Skipping tiny region: area={:.2f} sqft".format(area))
                continue
            
            # Find a point guaranteed to be inside the polygon
            # (centroid may be outside for non-convex shapes)
            interior_point = _find_interior_point(contour, debug=debug)
            if interior_point is None:
                if debug:
                    print("  [2D] Could not find interior point for contour")
                continue
            
            cx, cy = interior_point
            
            regions.append({
                'contour': contour,
                'centroid': (cx, cy, 0.0),
                'area': area
            })
        
        # Sort by area descending
        regions.sort(key=lambda r: r['area'], reverse=True)
        
        if debug:
            print("  [2D] Identified {} gap regions:".format(len(regions)))
            for i, r in enumerate(regions):
                print("    Region {}: centroid=({:.2f}, {:.2f}), area={:.2f} sqft".format(
                    i, r['centroid'][0], r['centroid'][1], r['area']))
        
        return regions
    
    except Exception as e:
        if debug:
            print("  [2D] Error finding gap regions: {}".format(e))
        return []


def find_all_gap_regions_2d_from_polygons(polygons, debug=False):
    """Find all gap regions using pre-created Polygon2D objects.
    
    This version takes Polygon2D objects directly, bypassing CurveLoop conversion.
    
    Algorithm:
    1. Create bounding box around all polygons
    2. Subtract each polygon from the bounding box
    3. The remaining regions are gaps between areas
    
    Args:
        polygons: List of Polygon2D objects
        debug: If True, print debug information
    
    Returns:
        List of dicts with 'contour', 'centroid', 'area' keys
    """
    if not polygons:
        return []
    
    if debug:
        print("  [2D] Processing {} polygons for gap detection".format(len(polygons)))
    
    try:
        # Get bounding box of all polygons
        all_min_x, all_min_y = float('inf'), float('inf')
        all_max_x, all_max_y = float('-inf'), float('-inf')
        
        for poly in polygons:
            bounds = poly.bounds
            if bounds:
                min_x, min_y, max_x, max_y = bounds
                all_min_x = min(all_min_x, min_x)
                all_min_y = min(all_min_y, min_y)
                all_max_x = max(all_max_x, max_x)
                all_max_y = max(all_max_y, max_y)
        
        if all_min_x == float('inf'):
            return []
        
        # Add small margin
        margin = 1.0
        all_min_x -= margin
        all_min_y -= margin
        all_max_x += margin
        all_max_y += margin
        
        if debug:
            print("  [2D] Bounding box: ({:.1f}, {:.1f}) to ({:.1f}, {:.1f})".format(
                all_min_x, all_min_y, all_max_x, all_max_y))
        
        # Create bounding rectangle
        bbox_rect = Polygon2D.create_rectangle(all_min_x, all_min_y, all_max_x, all_max_y)
        
        # Subtract each polygon from the bounding rectangle
        result = bbox_rect
        for i, poly in enumerate(polygons):
            result = result.difference(poly)
            if result.is_empty:
                if debug:
                    print("  [2D] Warning: result became empty after subtracting polygon {}".format(i))
                break
        
        if result.is_empty:
            if debug:
                print("  [2D] No gaps found (areas cover entire bounding box)")
            return []
        
        # Get all contours from the result (gaps)
        all_contours_raw = result.get_contours()
        
        if debug:
            print("  [2D] Found {} raw contours after subtraction".format(len(all_contours_raw)))
        
        # Filter out the outer bounding box contour (the margin around everything)
        contour_data = []
        for contour in all_contours_raw:
            if len(contour) < 3:
                continue
            area = abs(Polygon2D._calculate_contour_area(contour))
            contour_data.append({
                'contour': contour,
                'area': area
            })
        
        if not contour_data:
            return []
        
        # Sort by area descending
        contour_data.sort(key=lambda x: x['area'], reverse=True)
        
        # The largest contour is the outer margin (bbox - all areas) - skip it
        if len(contour_data) > 1:
            outer_margin_area = contour_data[0]['area']
            if debug:
                print("  [2D] Outer margin area: {:.2f} sqft (filtering out)".format(outer_margin_area))
            all_contours = [cd['contour'] for cd in contour_data[1:]]
        else:
            all_contours = [contour_data[0]['contour']]
        
        if debug:
            print("  [2D] Found {} interior gap contours".format(len(all_contours)))
        
        # Process each contour
        regions = []
        for contour in all_contours:
            if len(contour) < 3:
                continue
            
            area = abs(Polygon2D._calculate_contour_area(contour))
            
            # Filter by area:
            # - Skip very small regions (< 0.5 sqft ≈ 0.05 sqm) - likely tessellation noise
            # - Skip very large regions (> 10,000 sqft) - likely the outer margin (bbox - areas)
            if area < 0.5:
                if debug:
                    print("  [2D] Skipping tiny region: area={:.2f} sqft".format(area))
                continue
            
            if area > 10000.0:
                if debug:
                    print("  [2D] Skipping huge region (outer margin): area={:.2f} sqft".format(area))
                continue
            
            # Find interior point
            interior_point = _find_interior_point(contour, debug=debug)
            if interior_point is None:
                if debug:
                    print("  [2D] Could not find interior point for region: area={:.2f} sqft".format(area))
                continue
            
            cx, cy = interior_point
            
            regions.append({
                'contour': contour,
                'centroid': (cx, cy, 0.0),
                'area': area
            })
        
        # Sort by area descending
        regions.sort(key=lambda r: r['area'], reverse=True)
        
        if debug:
            print("  [2D] Identified {} gap regions:".format(len(regions)))
            for i, r in enumerate(regions):
                print("    Region {}: centroid=({:.2f}, {:.2f}), area={:.2f} sqft".format(
                    i, r['centroid'][0], r['centroid'][1], r['area']))
        
        return regions
    
    except Exception as e:
        if debug:
            print("  [2D] Error finding gap regions: {}".format(e))
        return []
