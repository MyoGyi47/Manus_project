"""
Utility functions for Geo Processor
"""
import re
import os
import zipfile
import xml.etree.ElementTree as ET
from shapely.geometry import Point, LineString, Polygon
import pandas as pd
from config import MM_DIGITS

# ==================== TEXT UTILITIES ====================
def convert_to_mm_digits(number_str):
    """Convert numbers to Burmese digits"""
    return ''.join(MM_DIGITS.get(char, char) for char in str(number_str))

def round_coordinate_for_phrase(coord_str, coord_type=None):
    """Apply custom rounding and slicing logic for coordinates"""
    coord_str = str(coord_str).strip()
    extracted_four_digits = None

    if coord_type == "easting":
        if len(coord_str) < 5:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[1:5]
    elif coord_type == "northing":
        if len(coord_str) < 6:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[2:6]
    else:
        if len(coord_str) < 6:
            return convert_to_mm_digits(coord_str)
        extracted_four_digits = coord_str[2:6]
    
    if not extracted_four_digits or not extracted_four_digits.isdigit():
        return convert_to_mm_digits(coord_str)

    try:
        num_with_decimal = float(extracted_four_digits[:-1] + '.' + extracted_four_digits[-1])
        rounded_val = round(num_with_decimal)
        result_str = str(int(rounded_val)).zfill(3)
        return convert_to_mm_digits(result_str)
    except (ValueError, IndexError):
        return convert_to_mm_digits(coord_str)

def move_del_to_remark(name, remark):
    """Move 'del' from name to remark"""
    match = re.search(r'\b(del|delete)\b', name, re.IGNORECASE)
    if match:
        split_index = match.start()
        moved_text = name[split_index:].strip()
        name = name[:split_index].strip()
        remark = f"{remark.strip()} | {moved_text}" if remark else moved_text
    return name, remark

def normalize_for_comparison(text):
    """Normalize text for comparison (remove spaces)"""
    return re.sub(r'\s+', '', text.lower()) if text else ""

# ==================== GEOMETRY UTILITIES ====================
def get_all_vertices(geometry):
    """Extract all vertices from Shapely geometry"""
    if geometry is None:
        return []
    
    geom_type = geometry.geom_type
    
    if geom_type in ['Point', 'MultiPoint']:
        points = geometry.geoms if geom_type == 'MultiPoint' else [geometry]
        return [(p.x, p.y) for p in points]
    
    elif geom_type in ['LineString', 'MultiLineString']:
        all_coords = []
        geometries = geometry.geoms if geom_type == 'MultiLineString' else [geometry]
        for line in geometries:
            coords = list(line.coords)
            all_coords.extend(coords[:-1] if len(coords) > 1 and coords[0] == coords[-1] else coords)
        return all_coords
    
    elif geom_type in ['Polygon', 'MultiPolygon']:
        all_coords = []
        geometries = geometry.geoms if geom_type == 'MultiPolygon' else [geometry]
        for poly in geometries:
            coords = list(poly.exterior.coords)
            all_coords.extend(coords[:-1] if coords and coords[0] == coords[-1] else coords)
        return all_coords
    
    return []

# ==================== FILE UTILITIES ====================
def extract_kml_from_kmz(kmz_path):
    """Extract KML content from KMZ file"""
    with zipfile.ZipFile(kmz_path, 'r') as kmz:
        for file_name in kmz.namelist():
            if file_name.endswith('.kml'):
                return kmz.read(file_name).decode('utf-8')
    raise FileNotFoundError("KMZ ဖိုင်အတွင်း KML ဖိုင်ကို ရှာမတွေ့ပါ။")

def kml_coords_to_list(coords_str):
    """Convert KML coordinates string to list of (lon, lat) tuples"""
    coords = []
    for point in coords_str.strip().split(' '):
        if point:
            lon, lat, *_ = map(float, point.split(","))
            coords.append((lon, lat))
    return coords

def detect_zone(lon, proj_choice):
    """Detect UTM zone based on longitude"""
    if proj_choice == "Custom_UTM":  # MMD2000
        return 47 if lon >= 95.9968784133 else 46
    else:  # WGS84 UTM or WGS84 LatLon
        return 47 if lon >= 96 else 46

def detect_zone_from_filename(filename, lon, projection, datum):
    """
    Detect zone from filename for UTM projections only
    """
    # Geographic မှာ zone မလိုဘူး
    if projection == "Geographic":
        return None
    
    # Same filename pattern logic as before
    import re
    import os
    
    base_name = os.path.basename(filename)
    name_without_ext = os.path.splitext(base_name)[0]
    clean_name = re.sub(r'[^a-zA-Z0-9\s\-]', ' ', name_without_ext)
    
    patterns = [
        r'(\d{4})[\s\-_](\d{1,2})',
        r'(\d{4})(\d{2})',
        r'(\d{2})[\s\-_](\d{1,2})',
        r'(\d{2})(\d{2})',
    ]
    
    for pattern in patterns:
        match = re.search(pattern, clean_name)
        if match:
            year_part = match.group(1)
            number_part = match.group(2)
            
            if len(year_part) == 4:
                year_last_two = year_part[2:]
            else:
                year_last_two = year_part
            
            try:
                year_num = int(year_last_two)
                num = int(number_part)
                
                if year_num == 95 and num in [13, 14, 15, 16]:
                    return 46
                elif year_num == 96 and num in [1, 2, 3, 4]:
                    return 47
                else:
                    return None  # Auto detection ကိုသွားမယ်
                    
            except ValueError:
                continue
    
    return None  # No pattern found, use auto detection
    
    return detect_zone(lon, proj_choice)  # Auto detection if no pattern matched

def get_zone_info_from_filename(filename, proj_choice):
    """
    Get human-readable zone info from filename
    Returns: (zone_number, message)
    """
    if not filename:
        return None, ""
    
    try:
        base_name = os.path.basename(filename)
        name_without_ext = os.path.splitext(base_name)[0]
        clean_name = re.sub(r'[^a-zA-Z0-9\s\-]', ' ', name_without_ext)
        
        patterns = [
            r'(\d{4})[\s\-_](\d{1,2})',
            r'(\d{4})(\d{2})',
            r'(\d{2})[\s\-_](\d{1,2})',
            r'(\d{2})(\d{2})',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, clean_name)
            if match:
                year_part = match.group(1)
                number_part = match.group(2)
                
                if len(year_part) == 4:
                    year_last_two = year_part[2:]
                else:
                    year_last_two = year_part
                
                try:
                    year_num = int(year_last_two)
                    num = int(number_part)
                    
                    # Check for specific patterns
                    if year_num == 95 and num in [13, 14, 15, 16]:
                        return 46, f"{'MMD' if proj_choice == 'Custom_UTM' else 'WGS84'} Zone 46 ဖြင့်အလုပ်လုပ်နေပါသည်"
                    
                    elif year_num == 96 and num in [1, 2, 3, 4]:
                        return 47, f"{'MMD' if proj_choice == 'Custom_UTM' else 'WGS84'} Zone 47 ဖြင့်အလုပ်လုပ်နေပါသည်"
                    
                except ValueError:
                    continue
    
    except Exception:
        pass
    
    return None, ""

# utils.py - zone detection function ကို update
def determine_zone(lon, projection, datum, zone_mode="auto", manual_zone=None, filename=None):
    """
    Determine zone based on multiple parameters
    Returns: zone number or None (for Geographic)
    """
    # Geographic ဆိုရင် zone မလိုဘူး
    if projection == "Geographic":
        return None
    
    # Manual zone ရွေးထားရင်
    if zone_mode == "manual" and manual_zone in [46, 47]:
        return manual_zone
    
    # Auto mode with filename
    if zone_mode == "auto" and filename:
        zone = detect_zone_from_filename(filename, lon, projection, datum)
        if zone:
            return zone
    
    # Default auto detection
    if datum == "MMD2000":
        return 47 if lon >= 95.9968784133 else 46
    else:  # WGS84
        return 47 if lon >= 96 else 46