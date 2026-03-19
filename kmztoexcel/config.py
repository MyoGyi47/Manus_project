"""
Configuration and constants for Geo Processor
"""

# Custom Datum PROJ strings
PROJ_46 = "+proj=utm +zone=46 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"
PROJ_47 = "+proj=utm +zone=47 +a=6377276.345 +rf=300.801698010253 +towgs84=246.632,784.833,276.923,0,0,0,0 +units=m +no_defs"

# Keywords for splitting logic
ALLOWED_KEYWORDS = ['ew', 'mr', 'sr', 'or', 'ct', 'pt', 'fp']

POLYGON_KEYWORDS = [
    'bua', 'lake', 'pond', 'religious area', 'sport field', 'martyrs temple',
    'dam', 'fish farm', 'swamp area', 'cultivation area', 'reservoir',
    'highway bus terminal compound', 'myit kyo in', 'solar panel', 'spill way',
    'river', 'cemetery area', 'water area repair', 'golf course', 'livestock farm'
]

# Remove duplicate "Ta da_U"
PROTECTED_SPLIT_NAMES_WORD = [
    'Kyauk_O', 'U_yin', 'Nga ku_Oh', 'Ta da_U', 'Tha bye_U',
    'Kyauk_O(Kyauk kon)', 'Kyun_U', 'San_U', 'Le gyin_U', 'Kan_U',
    'Laung daw_U', 'O_gyi gwe', 'O_yin', 'Daung_U', 'Chaung_U',
    'O_pon daw', 'Chaung zon', "Chaung gwa", "Chaung bat", "U_hmin", "TADA_U"
]

ROAD_KEYWORDS = [
    "main road", "main rd", "mainroad", "main-road",
    "secondary road", "secondary rd", "secondary-road",
    "cart track", "cart-track", "carttrack",
    "other road", "other rd", "other-road",
    "pack track", "packtrack", "pack-track",
    "footpath", "foot path", "foot-path",
    'canal', 'stream', 'chaung', 'embankment',
    'river', 'fish farm', "express way", "expressway", "express-way", "zaung dan"
]

# Burmese digit mapping
MM_DIGITS = {
    '0': '၀', '1': '၁', '2': '၂', '3': '၃', '4': '၄',
    '5': '၅', '6': '၆', '7': '၇', '8': '၈', '9': '၉'
}