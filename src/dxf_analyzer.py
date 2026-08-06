import os
import ezdxf
from ezdxf.math import Vec2

def detect_dxf_encoding(dxf_path):
    """
    Checks if the DXF file is encoded in UTF-8 by trying to decode a 256KB block.
    If it fails, falls back to CP949 (Korean legacy encoding).
    """
    try:
        with open(dxf_path, 'rb') as f:
            chunk = f.read(256 * 1024)
        # Try decoding as UTF-8
        chunk.decode('utf-8')
        return 'utf-8'
    except UnicodeDecodeError:
        return 'cp949'

def get_hatch_geometry_area(hatch):
    """
    Directly extracts vertices from Hatch boundaries and calculates the area using the Shoelace formula.
    Extremely fast and avoids importing ezdxf.path or font caching.
    """
    total_area = 0.0
    try:
        for path in hatch.paths:
            pts = []
            if hasattr(path, 'vertices') and path.vertices:
                # PolylinePath
                for v in path.vertices:
                    pts.append(Vec2(v[0], v[1]))
            elif hasattr(path, 'edges') and path.edges:
                # EdgePath
                for edge in path.edges:
                    edge_type = getattr(edge, 'EDGE_TYPE', None)
                    if edge_type == 'LineEdge':
                        pts.append(Vec2(edge.start[0], edge.start[1]))
                        pts.append(Vec2(edge.end[0], edge.end[1]))
                    elif edge_type == 'ArcEdge':
                        pts.append(Vec2(edge.start[0], edge.start[1]))
                        pts.append(Vec2(edge.end[0], edge.end[1]))
                    elif edge_type == 'SplineEdge':
                        for p in getattr(edge, 'control_points', []):
                            pts.append(Vec2(p[0], p[1]))
                    else:
                        if hasattr(edge, 'start'):
                            pts.append(Vec2(edge.start[0], edge.start[1]))
                        if hasattr(edge, 'end'):
                            pts.append(Vec2(edge.end[0], edge.end[1]))
            
            # Clean consecutive duplicates
            cleaned_pts = []
            for p in pts:
                if not cleaned_pts or cleaned_pts[-1] != p:
                    cleaned_pts.append(p)
            
            if len(cleaned_pts) >= 3:
                area_val = 0.0
                n = len(cleaned_pts)
                for i in range(n):
                    j = (i + 1) % n
                    area_val += cleaned_pts[i].x * cleaned_pts[j].y
                    area_val -= cleaned_pts[j].x * cleaned_pts[i].y
                total_area += abs(area_val) / 2.0
    except Exception as e:
        print(f"Error calculating hatch area: {e}")
    return total_area

def get_hatch_layers(dxf_path):
    """
    Returns a dict mapping layer_name -> hatch_count for layers that contain HATCH entities.
    Also returns the full list of all layers in the file.
    """
    if not os.path.exists(dxf_path):
        return {}, []
    
    encoding = detect_dxf_encoding(dxf_path)
    doc = ezdxf.readfile(dxf_path, encoding=encoding)
    msp = doc.modelspace()
    
    hatch_layer_counts = {}
    hatches = msp.query('HATCH')
    for hatch in hatches:
        layer_name = hatch.dxf.layer
        hatch_layer_counts[layer_name] = hatch_layer_counts.get(layer_name, 0) + 1
    
    all_layers = sorted(set(l.dxf.name for l in doc.layers))
    return hatch_layer_counts, all_layers


def analyze_dxf_hatch_only(dxf_path, boundary_layer=None, selected_layers=None):
    """
    Analyzes a DXF file to compute HATCH areas only (no polylines).
    If selected_layers is provided, only hatches on those layers are included.
    """
    if not os.path.exists(dxf_path):
        raise FileNotFoundError(f"File not found: {dxf_path}")
    
    encoding = detect_dxf_encoding(dxf_path)
    doc = ezdxf.readfile(dxf_path, encoding=encoding)
    msp = doc.modelspace()
    
    layer_areas = {}
    
    hatches = msp.query('HATCH')
    for hatch in hatches:
        layer_name = hatch.dxf.layer
        
        # Skip layers not in the selected list (if provided)
        if selected_layers is not None and layer_name not in selected_layers:
            continue
        
        area = get_hatch_geometry_area(hatch)
        if area > 0:
            layer_areas[layer_name] = layer_areas.get(layer_name, 0.0) + area
    
    # Calculate Boundary Area (구역계)
    boundary_area_val = 0.0
    if boundary_layer and boundary_layer in layer_areas:
        boundary_area_val = layer_areas[boundary_layer]
    else:
        # Fallback: try to find polylines on boundary layer for boundary area only
        if boundary_layer:
            boundary_polys = msp.query(f'*[layer=="{boundary_layer}"]')
            for pl in boundary_polys:
                if pl.dxftype() in ('LWPOLYLINE', 'POLYLINE') and pl.is_closed:
                    try:
                        pts = [Vec2(p[0], p[1]) for p in pl.get_points(format='xy')]
                        if len(pts) >= 3:
                            area_val = 0.0
                            n = len(pts)
                            for i in range(n):
                                j = (i + 1) % n
                                area_val += pts[i].x * pts[j].y
                                area_val -= pts[j].x * pts[i].y
                            boundary_area_val += abs(area_val) / 2.0
                    except Exception:
                        pass
    
    # If boundary area is still 0, sum up all layers (excluding boundary layer)
    if boundary_area_val == 0.0:
        boundary_area_val = sum(area for name, area in layer_areas.items() if name != boundary_layer)
    
    if boundary_area_val == 0.0:
        boundary_area_val = 1.0  # Avoid division by zero
    
    return {
        'layers': layer_areas,
        'boundary_area': boundary_area_val
    }


def analyze_dxf(dxf_path, boundary_layer=None):
    """
    Analyzes a DXF file to compute hatch and closed polyline areas by layer.
    """
    if not os.path.exists(dxf_path):
        raise FileNotFoundError(f"File not found: {dxf_path}")
        
    encoding = detect_dxf_encoding(dxf_path)
    doc = ezdxf.readfile(dxf_path, encoding=encoding)
    msp = doc.modelspace()
    
    layer_areas = {}
    hatch_layers = set()
    
    # 1. Analyze Hatches (Preferred for zones)
    hatches = msp.query('HATCH')
    for hatch in hatches:
        layer_name = hatch.dxf.layer
        area = get_hatch_geometry_area(hatch)
        if area > 0:
            layer_areas[layer_name] = layer_areas.get(layer_name, 0.0) + area
            hatch_layers.add(layer_name)
            
    # 2. Analyze Closed Polylines (For layers without hatches, e.g., outlines or roads)
    polylines = msp.query('LWPOLYLINE POLYLINE')
    for pl in polylines:
        if pl.is_closed:
            layer_name = pl.dxf.layer
            # Skip if we already have hatch areas for this layer to avoid double counting
            if layer_name not in hatch_layers:
                try:
                    pts = [Vec2(p[0], p[1]) for p in pl.get_points(format='xy')]
                    if len(pts) >= 3:
                        area_val = 0.0
                        n = len(pts)
                        for i in range(n):
                            j = (i + 1) % n
                            area_val += pts[i].x * pts[j].y
                            area_val -= pts[j].x * pts[i].y
                        area = abs(area_val) / 2.0
                        if area > 0:
                            layer_areas[layer_name] = layer_areas.get(layer_name, 0.0) + area
                except Exception as e:
                    print(f"Warning: Failed to calculate polyline area on layer {layer_name}: {e}")
                    
    # Calculate Boundary Area (구역계)
    boundary_area_val = 0.0
    if boundary_layer and boundary_layer in layer_areas:
        boundary_area_val = layer_areas[boundary_layer]
    else:
        # Fallback: if boundary layer is not in layer_areas, try to find polylines on it directly
        if boundary_layer:
            boundary_polys = msp.query(f'*[layer=="{boundary_layer}"]')
            for pl in boundary_polys:
                if pl.dxftype() in ('LWPOLYLINE', 'POLYLINE') and pl.is_closed:
                    try:
                        pts = [Vec2(p[0], p[1]) for p in pl.get_points(format='xy')]
                        if len(pts) >= 3:
                            area_val = 0.0
                            n = len(pts)
                            for i in range(n):
                                j = (i + 1) % n
                                area_val += pts[i].x * pts[j].y
                                area_val -= pts[j].x * pts[i].y
                            boundary_area_val += abs(area_val) / 2.0
                    except Exception:
                        pass
        
    # If boundary area is still 0, sum up all layers (excluding boundary layer itself if it's there)
    if boundary_area_val == 0.0:
        boundary_area_val = sum(area for name, area in layer_areas.items() if name != boundary_layer)
        
    # If there are no layers calculated, boundary is 0
    if boundary_area_val == 0.0:
        boundary_area_val = 1.0  # Avoid division by zero
        
    return {
        'layers': layer_areas,
        'boundary_area': boundary_area_val
    }

def get_all_layers(dxf_path):
    """
    Extremely fast text-based layer extractor. Parses only the LAYER table in < 0.2s.
    """
    if not os.path.exists(dxf_path):
        return []
        
    encoding = detect_dxf_encoding(dxf_path)
        
    layers = set()
    try:
        with open(dxf_path, 'r', encoding=encoding, errors='ignore') as f:
            group_code = None
            is_layer_record = False
            for line in f:
                line = line.strip()
                if group_code is None:
                    try:
                        group_code = int(line)
                    except ValueError:
                        pass
                else:
                    if group_code == 0:
                        is_layer_record = (line == "LAYER")
                    elif group_code == 2 and is_layer_record:
                        layers.add(line)
                    group_code = None
    except Exception as e:
        print(f"Error in fast layer read: {e}")
        
    return sorted(list(layers))
