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

def detect_zone_from_filename(filename, lon, proj_choice):
    """
    Detect zone from filename patterns
    2195 13-16 → Zone 46
    2196 01-04 → Zone 47
    Others → Auto detection
    """
    base_name = os.path.basename(filename)
    name_without_ext = os.path.splitext(base_name)[0]
    clean_name = re.sub(r'[^a-zA-Z0-9\s\-]', ' ', name_without_ext)
    
    patterns = [
        r'(\d{4})[\s\-_](\d{1,2})',      # 2195 15, 2195-15
        r'(\d{4})(\d{2})',               # 219515
        r'(\d{2})[\s\-_](\d{1,2})',      # 95 15, 95-15
        r'(\d{2})(\d{2})',               # 9515
    ]
    
    for pattern in patterns:
        match = re.search(pattern, clean_name)
        if match:
            year_part = match.group(1)
            number_part = match.group(2)
            
            if len(year_part) == 4:
                year_last_two = year_part[2:]  # 2195 -> 95
            else:
                year_last_two = year_part
            
            try:
                year_num = int(year_last_two)
                num = int(number_part)
                
                # Simplified logic as requested
                if year_num == 95:
                    if num in [13, 14, 15, 16]:
                        return 46  # Zone 46
                    else:
                        return detect_zone(lon, proj_choice)  # Auto detection
                
                elif year_num == 96:
                    if num in [1, 2, 3, 4]:
                        return 47  # Zone 47
                    else:
                        return detect_zone(lon, proj_choice)  # Auto detection
                
                else:
                    return detect_zone(lon, proj_choice)  # Auto detection
                    
            except ValueError:
                continue
    
    return detect_zone(lon, proj_choice)  # Auto detection if no pattern matched