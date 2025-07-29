#!/usr/bin/env python3
"""
Simple Lens Parser
==================

A lightweight lens name parser that extracts structured data from raw lens names
using regex patterns and string matching.
"""

import csv
import re
import pandas as pd
from pathlib import Path
from typing import Tuple, Dict, List, Optional
from dataclasses import dataclass
import json  # Added for dynamic loading of learned patterns

@dataclass
class ParsedLens:
    """Data class for parsed lens information"""
    manufacturer: str = ""
    series: str = ""
    focal_length: str = ""
    t_stop: str = ""
    lens_type: str = ""
    format: str = ""
    mount: str = ""
    anamorphic_spherical: str = ""
    anamorphic_squeeze: str = ""
    anamorphic_location: str = ""
    housing: str = ""
    front_diameter: str = ""
    close_focus: str = ""
    length: str = ""
    film_compatibility: str = ""
    image_circle: str = ""
    iris_blade_count: str = ""
    extender: str = ""
    lds: str = ""
    idata: str = ""
    support_recommended: str = ""
    support_post_length: str = ""
    weight: str = ""
    manufacture_year: str = ""
    expander: str = ""
    heden_motor_size: str = ""
    size: str = ""
    notes: str = ""
    look: str = ""
    use_case: str = ""
    bokeh: str = ""
    flare: str = ""
    focus_falloff: str = ""
    breathing: str = ""
    focus_scale: str = ""
    original_name: str = ""
    needs_review: bool = False
    confidence_score: float = 0.0

class SimpleLensParser:
    """Simple lens parser using regex and string matching"""
    
    def __init__(self):
        # Manufacturer patterns
        self.manufacturers = {
            'angenieux': ['angenieux'],
            'canon': ['canon'],
            'cooke': ['cooke'],
            'zeiss': ['zeiss', 'arri/zeiss', 'arri / zeiss'],
            'leica': ['leica'],
            'leitz': ['leitz'],
            'sigma': ['sigma'],
            'laowa': ['laowa'],
            'atlas': ['atlas'],
            'caldwell': ['caldwell'],
            'master': ['master'],
            'ancient optics': ['ancient optics'],
            'zero optik': ['zero optik'],
            'tls': ['tls'],
            'optex': ['optex'],
            'infinity': ['infinity'],
            'lindsey': ['lindsey'],
            'century': ['century'],
            'sony': ['sony'],
            'dzofilm': ['dzofilm', 'dzofilms'],
            'gecko-cam': ['gecko-cam'],
            'kowa': ['kowa'],
            'scorpio': ['scorpio'],
            'masterbuilt': ['masterbuilt'],
            'tribe7': ['tribe7'],
            'hawk': ['hawk'],
            'petzval': ['petzval'],
            'fujinon': ['fujinon'],
            'lensbaby': ['lensbaby'],
            'cci': ['cci'],
            'lomo': ['lomo'],
            'arri': ['arri'],
            'optika': ['optika'],
            'konica': ['konica'],
            'fuji': ['fuji'],
            'duclos': ['duclos'],
            'nikon': ['nikon'],
            'swift 960': ['swift 960'],
            'schneider kreuznach': ['schneider kreuznach'],
            'p+s technik': ['p+s technik'],
            'astroscope': ['astroscope'],
            'nanmorph': ['nanmorph'],
            'infiniprobe': ['infiniprobe'],
            'rodenstock': ['rodenstock'],
            'kish': ['kish'],
            'keslow': ['keslow', 'kes-low'],
            'hawk': ['hawk'],
            'second reef': ['second reef'],
            'ironglass': ['ironglass'],
            'lensworks': ['lensworks'],
            'xelmus': ['xelmus'],
            'voigtlander': ['voigtlander'],
            'praxis': ['praxis']
        }
        
        # Series patterns
        self.series_patterns = {
            'master prime': ['master prime'],
            'ultra prime': ['ultra prime'],
            'super speed': ['super speed'],
            'standard speed': ['standard speed'],
            'signature prime': ['signature prime'],
            's4/i': ['s4/i', 's4i'],
            's5/i': ['s5/i', 's5i'],
            's7/i': ['s7/i', 's7i'],
            'panchro/i': ['panchro/i', 'panchroi'],
            'anamorphic/i': ['anamorphic/i', 'anamorphici'],
            'anamorphic ff plus': ['anamorphic ff plus'],
            'anamorphic sf ff plus': ['anamorphic sf ff plus'],
            'sp3': ['sp3'],
            'optimo': ['optimo'],
            'optimo ultra': ['optimo ultra'],
            'optimo ultra compact': ['optimo ultra compact'],
            'optimo dp': ['optimo dp'],
            'optimo prime': ['optimo prime'],
            'optimo style': ['optimo style'],
            'optimo vintage': ['optimo vintage'],
            'optimo hr': ['optimo hr'],
            'optimo anamorphic': ['optimo anamorphic'],
            'optimo anamorphic hr': ['optimo anamorphic hr'],
            'ez-1': ['ez-1', 'ez1'],
            'ez-2': ['ez-2', 'ez2'],
            'a-2s': ['a-2s', 'a2s'],
            's2': ['s2'],
            'alura': ['alura'],
            'cabrio': ['cabrio'],
            'premier': ['premier'],
            'rangefinder': ['rangefinder'],
            'fd': ['fd'],
            'ef': ['ef'],
            'ef-s': ['ef-s', 'efs'],
            'l series': ['l series'],
            'k-35': ['k-35', 'k35'],
            'nikkor': ['nikkor'],
            'fisheye': ['fisheye'],
            'hawk 65': ['hawk 65'],
            'orion': ['orion'],
            'mercury': ['mercury'],
            'silver edition': ['silver edition'],
            'chameleon': ['chameleon'],
            'nanomorph': ['nanomorph'],
            'proteus': ['proteus'],
            'spherical primes': ['spherical primes'],
            'vintage primes': ['vintage primes'],
            'tegea': ['tegea'],
            'super cine': ['super cine'],
            'elite': ['elite'],
            'illumina': ['illumina'],
            'varopanchr': ['varopanchr'],
            'vario-sonnar': ['vario-sonnar'],
            'lwz.1': ['lwz.1', 'lwz1'],
            'lwz.2': ['lwz.2', 'lwz2'],
            'master zoom': ['master zoom'],
            'ultra wide zoom': ['ultra wide zoom'],
            'variable prime': ['variable prime'],
            's16 zooms': ['s16 zooms'],
            'cinema zoom': ['cinema zoom'],
            'cine zoom': ['cine zoom'],
            'broadcast zoom': ['broadcast zoom'],
            'eng zoom': ['eng zoom'],
            'efp zoom': ['efp zoom'],
            'studio zoom': ['studio zoom'],
            'field zoom': ['field zoom'],
            'portrait': ['portrait'],
            'macro': ['macro'],
            'fisheye': ['fisheye'],
            'tilt-shift': ['tilt-shift', 'tilt shift'],
            'lensbaby': ['lensbaby'],
            'holga': ['holga'],
            'diana': ['diana'],
            'lomo': ['lomo'],
            'summilux': ['summilux', 'summilux c', 'summilux-c'],
            'summicron': ['summicron', 'summicron c', 'summicron-c'],
            'varotal': ['varotal'],
            'sk4': ['sk4'],
            's2000': ['s2000'],
            'fe': ['fe'],
            'gm': ['gm'],
            'vario-tessar': ['vario-tessar'],
            'genesis': ['genesis'],
            'vespid': ['vespid'],
            'pavo': ['pavo'],
            'arles': ['arles'],
            'x-tract': ['x-tract', 'x tract'],
            'signature zoom': ['signature zoom'],
            'variable zoom': ['variable zoom'],
            'swing shift': ['swing shift'],
            'special flare': ['sf'],
            'gnosis': ['gnosis'],
            'pro2be': ['pro2be'],
            'shift and tilt': ['shift and tilt'],
            'phenix': ['phenix'],
            'petzvalux': ['petzvalux'],
            'compact zoom': ['compact zoom'],
            'supreme prime radiance': ['supreme prime radiance'],
            'supreme prime': ['supreme prime'],
            'hexanon': ['hexanon'],
            'compact prime cp2': ['compact prime cp2'],
            'compact prime cp3': ['compact prime cp3'],
            'ebc': ['ebc'],
            'k-35': ['k-35', 'k35'],
            'elite': ['elite'],
            'optex': ['optex'],
            'chameleon sc/xc': ['chameleon sc/xc'],
            'chameleon xc': ['chameleon xc'],
            'chameleon uw sc': ['chameleon uw sc'],
            'vista one': ['vista one'],
            'vespid retro': ['vespid retro'],
            't-rex': ['t-rex'],
            'rangefinder': ['rangefinder'],
            'one': ['one'],
            'genesis g35': ['genesis g35'],
            'genesis g65': ['genesis g65'],
            'cine orange flare': ['cine orange flare'],
            'cine blue flare': ['cine blue flare'],
            'cine gold flare': ['cine gold flare'],
            'coral': ['coral'],
            'neo-ao': ['neo-ao'],
            'hugo': ['hugo'],
            'proteus': ['proteus'],
            'v-lite': ['v-lite'],
            'thalia': ['thalia'],
            'snorricam': ['snorricam'],
            'peephole lens': ['peephole lens'],
            'kaleidoscope lens': ['kaleidoscope lens'],
            'rifle scope': ['rifle scope'],
            'squishy lens': ['squishy lens'],
            'image shaker': ['image shaker'],
            'flow motion lens system': ['flow motion lens system'],
            'low angle mirror': ['low angle mirror'],
            'sf': ['sf'],
            'apollo': ['apollo'],
            'noctilux': ['noctilux'],
            'pure reach periscope': ['pure reach periscope']
        }
        
        # Format patterns - improved based on manual corrections
        self.format_patterns = {
            's35': ['s35', 'super 35', 'super35'],
            'full frame': ['full frame', 'ff', '35mm'],
            '16mm': ['16mm', '16 mm'],
            's16': ['s16', 'super 16', 'super16'],
            'aps-c': ['aps-c', 'apsc'],
            'm43': ['m43', 'micro four thirds', 'micro 4/3'],
            'medium format': ['medium format', 'mf'],
            'ff': ['ff'],
            'vv': ['vv', 'ff/vv']
        }
        
        # Mount patterns - improved based on manual corrections
        self.mount_patterns = {
            'pl': ['pl', 'pl mount'],
            'lpl': ['lpl', 'lpl mount'],
            'ef': ['ef', 'ef mount'],
            'rf': ['rf', 'rf mount'],
            'e-mount': ['e-mount', 'e mount', 'sony e'],
            'z-mount': ['z-mount', 'z mount', 'nikon z'],
            'f-mount': ['f-mount', 'f mount', 'nikon f'],
            'bayonet': ['bayonet'],
            'm42': ['m42', 'm42 mount'],
            'm39': ['m39', 'm39 mount'],
            'e': ['e'],
            'eos': ['eos']
        }
        
        # Anamorphic patterns
        self.anamorphic_patterns = {
            'anamorphic': ['anamorphic', 'ana'],
            'spherical': ['spherical', 'sph']
        }
        
        # Squeeze factor patterns - improved based on manual corrections
        self.squeeze_patterns = {
            '1.3x': ['1.3x', '1.3'],
            '1.5x': ['1.5x', '1.5'],
            '1.8x': ['1.8x', '1.8'],
            '2x': ['2x', '2.0x', '2.0'],
            '2.4x': ['2.4x', '2.4']
        }

        # --- Auto-merge patterns learned from Manual Edits ---
        patterns_file = Path(__file__).with_name('learned_patterns.json')
        if patterns_file.exists():
            try:
                with patterns_file.open() as fp:
                    learned_data = json.load(fp).get('manual_patterns', {})
                # Merge manufacturers
                for manu in learned_data.get('manufacturers', {}):
                    if manu not in self.manufacturers:
                        self.manufacturers[manu] = [manu]
                # Merge series
                for series in learned_data.get('series', {}):
                    if series not in self.series_patterns:
                        self.series_patterns[series] = [series]
                # Merge mounts
                for mnt in learned_data.get('mounts', {}):
                    if mnt not in self.mount_patterns:
                        self.mount_patterns[mnt] = [mnt]
                # Merge formats
                for fmt in learned_data.get('formats', {}):
                    if fmt not in self.format_patterns:
                        self.format_patterns[fmt] = [fmt]
            except Exception as exc:
                print(f"[SimpleLensParser] Warning: could not merge learned patterns: {exc}")

    def preprocess_text(self, text: str) -> str:
        """Preprocess the lens name text"""
        if not text:
            return ""
        
        # Convert to lowercase and normalize spacing
        text = text.lower().strip()
        text = re.sub(r'\s+', ' ', text)
        
        return text

    def identify_manufacturer(self, text: str) -> Tuple[str, float]:
        """Identify manufacturer from text"""
        best_match = None
        best_score = 0
        
        for manufacturer, aliases in self.manufacturers.items():
            for alias in aliases:
                if alias in text:
                    score = len(alias) / len(text) * 100
                    if score > best_score:
                        best_score = score
                        best_match = manufacturer
        
        if best_score >= 3 and best_match:  # Lowered threshold from 10 to 3
            return best_match.title(), min(best_score / 100, 0.9)
        return "", 0.0

    def identify_series(self, text: str) -> Tuple[str, float]:
        """Identify series from text - improved to handle complex series names"""
        best_match = None
        best_score = 0
        
        # Check for exact matches first (higher priority)
        for series, patterns in self.series_patterns.items():
            for pattern in patterns:
                if pattern in text:
                    score = len(pattern) / len(text) * 100
                    if score > best_score:
                        best_score = score
                        best_match = series

        
        # Special handling for complex series names
        if 'master anamorphic' in text.lower():
            return "Master Anamorphic", 0.9
        elif 'optimo ultra compact' in text.lower():
            return "Optimo Ultra Compact", 0.9
        elif 'optimo ultra' in text.lower():
            return "Optimo Ultra", 0.9
        elif 'optimo dp' in text.lower():
            return "Optimo DP", 0.9
        elif 'optimo prime' in text.lower():
            return "Optimo Prime", 0.9
        elif 'optimo style' in text.lower():
            return "Optimo Style", 0.9
        elif 'optimo vintage' in text.lower():
            return "Optimo Vintage", 0.9
        elif 'optimo hr' in text.lower():
            return "Optimo HR", 0.9
        elif 'optimo anamorphic' in text.lower():
            return "Optimo Anamorphic", 0.9
        elif 'optimo anamorphic hr' in text.lower():
            return "Optimo Anamorphic HR", 0.9
        elif 'chameleon sc/xc' in text.lower():
            return "Chameleon SC/XC", 0.9
        elif 'chameleon xc' in text.lower():
            return "Chameleon XC", 0.9
        elif 'chameleon uw sc' in text.lower():
            return "Chameleon UW SC", 0.9
        elif 'nanomorph' in text.lower():
            return "Nanomorph", 0.9
        elif 'genesis g35' in text.lower():
            return "Genesis G35", 0.9
        elif 'genesis g65' in text.lower():
            return "Genesis G65", 0.9
        elif 'vespid retro' in text.lower():
            return "Vespid Retro", 0.9
        elif 'pavo' in text.lower():
            return "Pavo", 0.9
        elif 'arles' in text.lower():
            return "Arles", 0.9
        elif 'x-tract' in text.lower():
            return "X-Tract", 0.9
        elif 'signature zoom' in text.lower():
            return "Signature Zoom", 0.9
        elif 'variable zoom' in text.lower():
            return "Variable Zoom", 0.9
        elif 'ranger' in text.lower():
            return "Ranger", 0.9
        elif 'orion' in text.lower():
            return "Orion", 0.9
        elif 'mercury' in text.lower():
            return "Mercury", 0.9
        elif 'silver edition' in text.lower():
            return "Silver Edition", 0.9
        # Add EZ-2 and EZ-1 detection
        elif 'ez-2' in text.lower() or 'ez2' in text.lower():
            return "EZ-2", 0.9
        elif 'ez-1' in text.lower() or 'ez1' in text.lower():
            return "EZ-1", 0.9
        # Removed Leica R special handling as 'R' is no longer treated as a valid series
        
        # Special handling for EBC false positive
        elif 'ebc' in text.lower() and not any(word in text.lower() for word in ['fujinon', 'fuji']):
            return "EBC", 0.9
        
        # Special handling for Shift and Tilt
        elif 'shift and tilt' in text.lower():
            return "Shift and Tilt", 0.9
        
        # Special handling for CINE FLARE series to preserve case
        elif 'cine orange flare' in text.lower():
            return "CINE ORANGE FLARE", 0.9
        elif 'cine blue flare' in text.lower():
            return "CINE BLUE FLARE", 0.9
        elif 'cine gold flare' in text.lower():
            return "CINE GOLD FLARE", 0.9
        
        if best_score >= 2 and best_match:  # Lowered threshold from 5 to 2
            return best_match.title(), min(best_score / 100, 0.9)
        return "", 0.0

    def extract_focal_length(self, text: str) -> Tuple[str, float]:
        """Extract focal length from text - improved to handle ranges and complex patterns"""
        # Pattern for focal length (e.g., 50mm, 24-70mm, 15.5-45mm, 100mm/150mm, 20mm-105mm, 24-290/26-320/36-435)
        patterns = [
            r'(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)\s*mm',  # Standard mm format with ranges
            r'(\d+(?:\.\d+)?(?:/\d+(?:\.\d+)?)?)\s*mm',  # Slash format like 100mm/150mm
            r'(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?(?:/\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)*)',  # Complex format like 24-290/26-320/36-435
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text)
            if match:
                focal_length = match.group(1)
                # Clean up the focal length
                focal_length = re.sub(r'\s+', '', focal_length)
                return focal_length, 0.9
        
        return "", 0.0

    def extract_t_stop(self, text: str) -> Tuple[str, float]:
        """Extract T-stop from text - improved to handle more patterns"""
        # Pattern for T-stop and F-stop (e.g., T2.8, T1.4, F2.8, F4.5-5.6, f/4.5-5.6, N/A)
        patterns = [
            r't(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)',  # T2.8, T1.4, T3.5-4.5
            r'f(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)',  # F2.8, F1.4, F4.5-5.6
            r'f/(\d+(?:\.\d+)?(?:-\d+(?:\.\d+)?)?)',  # f/2.8, f/4.5-5.6
            r'n/a',  # N/A values
        ]
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                if pattern == r'n/a':
                    return "N/A", 0.9
                else:
                    # Determine if it's T or F based on the pattern
                    if pattern.startswith('f'):
                        t_stop = f"F{match.group(1)}"
                    else:
                        t_stop = f"T{match.group(1)}"
                    return t_stop, 0.9
        
        return "", 0.0

    def determine_lens_type(self, text: str, focal_length: str) -> Tuple[str, float]:
        """Determine if lens is prime, zoom, or special - improved to detect ranges"""
        # Check for zoom indicators in text
        zoom_indicators = ['zoom', 'varotal', 'cabrio', 'servo', 'cine-servo']
        for indicator in zoom_indicators:
            if indicator in text.lower():
                return "Zoom", 0.9
        
        # Check if focal length contains a range (indicating zoom)
        if focal_length and ('-' in focal_length or '/' in focal_length):
            return "Zoom", 0.9
        
        # Check for special lens types
        special_indicators = ['tilt-shift', 'shift', 'swing', 'lensbaby', 'composer']
        for indicator in special_indicators:
            if indicator in text.lower():
                return "Special", 0.9
        
        # Default to prime if no other indicators found
        return "Prime", 0.7

    def identify_format(self, text: str) -> Tuple[str, float]:
        """Identify format from text - improved to avoid false positives"""
        # Check for specific format patterns in order of priority
        text_lower = text.lower()
        
        # Exception: don't set format if "doesn't cover" is mentioned
        if 'doesn\'t cover' in text_lower:
            return "", 0.0
        
        # Check for FF first (higher priority)
        if 'ff' in text_lower:
            return "FF", 0.9
        
        # Check for other format patterns with more specific matching
        format_patterns = {
            's35': ['s35', 'super 35', 'super35'],
            'full frame': ['full frame'],
            's16': ['s16', 'super 16', 'super16'],
            'aps-c': ['aps-c', 'apsc'],
            'm43': ['m43', 'micro four thirds', 'micro 4/3'],
            'vv': ['vv', 'ff/vv']
        }
        
        # Check for 16mm format more carefully to avoid focal length false positives
        if '16mm' in text_lower or '16 mm' in text_lower:
            # Only consider it a format if it's not part of a focal length range
            # Look for patterns like "16mm" that are not preceded by numbers or followed by focal length indicators
            if not re.search(r'\d+\s*-\s*16mm', text_lower) and not re.search(r'16mm\s*-\s*\d+', text_lower):
                return "16MM", 0.9
        
        for format_name, patterns in format_patterns.items():
            for pattern in patterns:
                if pattern in text_lower:
                    return format_name.upper(), 0.9
        
        # Don't assume any format if not explicitly mentioned
        return "", 0.0

    def identify_mount(self, text: str) -> Tuple[str, float]:
        """Identify mount from text - improved to avoid false positives"""
        # Check for specific mount patterns first - PL must be standalone
        mount_patterns = {
            'lpl': ['lpl', 'lpl mount', '(lpl)'],
            'pl': [' pl ', '(pl)', ' pl,', ' pl.', ' pl)'],  # PL must be surrounded by spaces, parentheses, or punctuation
            'ef': ['ef', 'ef mount'],
            'rf': ['rf', 'rf mount'],
            'e-mount': ['e-mount', 'e mount', 'sony e'],
            'z-mount': ['z-mount', 'z mount', 'nikon z'],
            'f-mount': ['f-mount', 'f mount', 'nikon f'],
            'bayonet': ['bayonet'],
            'm42': ['m42', 'm42 mount'],
            'm39': ['m39', 'm39 mount'],
            'eos': ['eos']
        }
        
        for mount_name, patterns in mount_patterns.items():
            for pattern in patterns:
                if pattern in text.lower():
                    return mount_name.upper(), 0.9
        
        # Only check for 'e' if it's clearly a mount reference
        if 'e mount' in text.lower() or 'sony e' in text.lower():
            return "E-MOUNT", 0.9
        
        return "", 0.0

    def identify_anamorphic_spherical(self, text: str) -> Tuple[str, float]:
        """Identify if lens is anamorphic or spherical"""
        for type_name, patterns in self.anamorphic_patterns.items():
            for pattern in patterns:
                if pattern in text:
                    return type_name.title(), 0.9
        
        # Default to spherical if no indication
        return "Spherical", 0.5

    def extract_squeeze_factor(self, text: str) -> Tuple[str, float]:
        """Extract anamorphic squeeze factor - only for anamorphic lenses with valid range"""
        # Only extract squeeze factor if the lens is anamorphic
        if 'anamorphic' not in text.lower():
            return "", 0.0
        
        # Squeeze factor patterns - must end with 'x' and be a number between 1-2
        squeeze_patterns = [
            r'(\d+(?:\.\d+)?)x',  # 1.8x, 2x, 1.33x
            r'x(\d+(?:\.\d+)?)',  # x1.8, x2, x1.33
        ]
        
        for pattern in squeeze_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                squeeze_value = float(match.group(1))
                # Only accept squeeze factors between 1 and 2
                if 1.0 <= squeeze_value <= 2.0:
                    squeeze_factor = match.group(1) + 'x'
                    return squeeze_factor, 0.9
        
        return "", 0.0

    def extract_anamorphic_location(self, text: str) -> Tuple[str, float]:
        """Extract anamorphic location - DISABLED as per user request"""
        # This field should be ignored as it's not in lens names
        return "", 0.0

    def extract_housing(self, text: str) -> Tuple[str, float]:
        """Extract housing information"""
        housing_indicators = [
            'original housing', 'rehoused', 'ancient optics', 'zero optik', 
            'tls', 'works cameras', 'whitepoint optics', 'gl optics'
        ]
        
        for indicator in housing_indicators:
            if indicator in text.lower():
                return indicator.title(), 0.8
        
        return "", 0.0

    def extract_use_case(self, text: str) -> Tuple[str, float]:
        """Extract use case information"""
        if 'macro' in text.lower():
            return "Macro", 0.9
        return "", 0.0

    def extract_look(self, text: str) -> Tuple[str, float]:
        """Extract look information"""
        if 'vintage' in text.lower():
            return "Vintage", 0.9
        return "", 0.0

    def calculate_confidence_score(self, parsed: 'ParsedLens') -> float:
        """Calculate confidence score based on how much of the original name is captured"""
        if not parsed.original_name:
            return 0.0
        
        # Get the original name and convert to lowercase for comparison
        original_name = parsed.original_name.lower()
        
        # Collect all the parsed values that should be found in the original name
        parsed_values = []
        
        # Core fields
        if parsed.manufacturer:
            parsed_values.append(parsed.manufacturer.lower())
        if parsed.series:
            parsed_values.append(parsed.series.lower())
        if parsed.focal_length:
            parsed_values.append(parsed.focal_length.lower())  # will handle mm length later
        if parsed.t_stop:
            parsed_values.append(parsed.t_stop.lower())
        if parsed.lens_type:
            parsed_values.append(parsed.lens_type.lower())
        
        # Also include notes if available
        if parsed.notes:
            parsed_values.append(parsed.notes.lower())
        
        if not parsed_values:
            return 0.0
        
        # Count how many characters from the original name are matched by parsed values
        total_original_chars = len(original_name.replace(' ', ''))  # Remove spaces for fair comparison
        matched_chars = 0
        
        clean_original = original_name.replace(' ', '').replace('-', '').replace('/', '').replace('(', '').replace(')', '')

        for value in parsed_values:
            clean_value = value.replace(' ', '').replace('-', '').replace('/', '').replace('(', '').replace(')', '')

            # Special handling for focal length where 'mm' may be omitted in parsed value
            if value == parsed.focal_length.lower():
                focal_with_mm = f"{clean_value}mm"
                if focal_with_mm in clean_original:
                    matched_chars += len(focal_with_mm)
                    continue  # skip normal add
            if clean_value in clean_original:
                matched_chars += len(clean_value)
        
        # Calculate percentage of original name that is captured
        if total_original_chars == 0:
            return 0.0
        
        confidence = min(matched_chars / total_original_chars, 1.0)
        return confidence

    def parse_lens_name(self, lens_name: str) -> ParsedLens:
        """Parse a single lens name"""
        if not lens_name:
            return ParsedLens(original_name=lens_name)
        
        text = self.preprocess_text(lens_name)
        scores = []
        
        # Extract manufacturer
        manufacturer, mfg_score = self.identify_manufacturer(text)
        scores.append(mfg_score)
        
        # Extract series
        series, series_score = self.identify_series(text)
        scores.append(series_score)
        
        # Extract focal length
        focal_length, fl_score = self.extract_focal_length(text)
        scores.append(fl_score)
        
        # Extract T-stop
        t_stop, t_score = self.extract_t_stop(text)
        scores.append(t_score)
        
        # Determine lens type
        lens_type, type_score = self.determine_lens_type(text, focal_length)
        scores.append(type_score)
        
        # Identify format
        format_info, format_score = self.identify_format(text)
        scores.append(format_score)
        
        # Identify mount
        mount, mount_score = self.identify_mount(text)
        scores.append(mount_score)
        
        # Identify anamorphic/spherical
        anamorphic_spherical, ana_score = self.identify_anamorphic_spherical(text)
        scores.append(ana_score)
        
        # Extract squeeze factor
        squeeze_factor, squeeze_score = self.extract_squeeze_factor(text)
        scores.append(squeeze_score)
        
        # Extract anamorphic location
        anamorphic_location, loc_score = self.extract_anamorphic_location(text)
        scores.append(loc_score)
        
        # Extract housing
        housing, housing_score = self.extract_housing(text)
        scores.append(housing_score)
        
        # Extract use case and look
        use_case, use_case_score = self.extract_use_case(text)
        look, look_score = self.extract_look(text)
        scores.append(use_case_score)
        scores.append(look_score)
        
        # Extract notes from parentheses
        notes = ""
        notes_match = re.search(r'\((.*?)\)', lens_name)
        if notes_match:
            notes = notes_match.group(1).strip()
        
        # Special logic for Master Anamorphic series
        if series == "Master Anamorphic":
            manufacturer = "Arri"
            anamorphic_spherical = "Anamorphic"
            format_info = "S35"
            mfg_score = 0.9
            ana_score = 0.9
            format_score = 0.9
        
        # Special logic for Fujinon cases
        if manufacturer == "Fujinon" and any(pattern in text.lower() for pattern in [
            'ha25x16.5', 'ha42x9.7', 'ha13x4.5', 'ha18x7.6', 'ha22x7.8', 
            'za12x4.5', 'za17x7.6', 'za22x7.6'
        ]):
            # Extract everything after "Fujinon" as the series, preserving original case
            series_match = re.search(r'fujinon\s+(.+)', lens_name, re.IGNORECASE)
            if series_match:
                series = series_match.group(1).strip()
                series_score = 0.9
            lens_type = "Zoom"
            anamorphic_spherical = "Spherical"
            type_score = 0.9
            ana_score = 0.9
        
        # Special logic for missing Super Speed/Standard Speed manufacturers
        if (series == "Super Speed" or series == "Standard Speed") and not manufacturer:
            manufacturer = "Zeiss"
            mfg_score = 0.8
        
        # Special logic for K-35 series (Canon)
        if series == "K-35" and not manufacturer:
            manufacturer = "Canon"
            mfg_score = 0.9
        
        # Special logic for Super Speed/Standard Speed (Zeiss)
        if (series == "Super Speed" or series == "Standard Speed") and not manufacturer:
            manufacturer = "Zeiss"
            mfg_score = 0.9
        
        # Special logic for Master Anamorphic (Arri)
        if series == "Master Anamorphic" and not manufacturer:
            manufacturer = "Arri"
            mfg_score = 0.9
        
        # Special logic for Leitz (Leica)
        if 'leitz' in text.lower() and not manufacturer:
            manufacturer = "Leica"
            mfg_score = 0.9
        
        # Special handling for individual special lenses
        special_lens_names = [
            'snorricam', 'peephole lens', 'kish kaleidoscope lens', 'rifle scope',
            'squishy lens', 'image shaker', 'astroscope night vision module',
            'keslow flow motion lens system', 'kes-low angle mirror',
            'p+s technik skater scope', 't-rex lens', 'century super wide low angle prism',
            'leica telephoto front module', 'leica telephoto rear module',
            'optex excellence probe', 'infiniprobe', 'distortion lens',
            'swift 960 series microscope lens', 'sim ethereal', 'ethereal',
            'rodenstock münchen doppel anastigmat eurynar'
        ]
        
        if any(special_name in lens_name.lower() for special_name in special_lens_names):
            lens_type = "Special"
            type_score = 0.9
        
        # Special logic for Sony GM series
        if manufacturer == "Sony" and ('gm' in text.lower() or 'fe gm' in text.lower()):
            series = "GM"
            series_score = 0.9
        
        # Special logic for SIM Ethereal series
        if 'sim ethereal' in text.lower() or 'ethereal' in text.lower():
            manufacturer = "SIM"
            series = "Ethereal"
            mfg_score = 0.9
            series_score = 0.9
            # Ensure it's not detected as Special
            if lens_type == "Special":
                lens_type = "Prime"
                type_score = 0.7
        
        # Special logic for Swift 960 series
        if 'swift 960' in text.lower():
            manufacturer = "Swift 960"
            series = ""
            mfg_score = 0.9
        
        # Special logic for EZ-2 series (Angenieux)
        if series == "EZ-2":
            manufacturer = "Angenieux"
            mfg_score = 0.9
        
        # Special logic for EZ-1 series (Angenieux)
        if series == "EZ-1":
            manufacturer = "Angenieux"
            mfg_score = 0.9
        
        # Special logic for Cooke SF (Special Flare) series
        if manufacturer == "Cooke" and series == "SF":
            series = "Special Flare"
            series_score = 0.9
        
        # Special logic for Cooke Anamorphic SF (Special Flare) series
        if manufacturer == "Cooke" and 'anamorphic sf' in text.lower():
            series = "Special Flare"
            series_score = 0.9
        
        # Extract flare color from CINE FLARE series
        flare_color = ""
        if any(flare_series in series.lower() for flare_series in ['cine orange flare', 'cine blue flare', 'cine gold flare']):
            if 'orange' in series.lower():
                flare_color = "Orange"
            elif 'blue' in series.lower():
                flare_color = "Blue"
            elif 'gold' in series.lower():
                flare_color = "Gold"
        
        # Calculate overall confidence
        confidence = self.calculate_confidence_score(ParsedLens(
            manufacturer=manufacturer,
            series=series,
            focal_length=focal_length,
            t_stop=t_stop,
            lens_type=lens_type,
            format=format_info,
            mount=mount,
            anamorphic_spherical=anamorphic_spherical,
            anamorphic_squeeze=squeeze_factor,
            anamorphic_location=anamorphic_location,
            housing=housing,
            notes=notes,
            original_name=lens_name,
            needs_review=False, # This will be set by calculate_confidence_score
            confidence_score=sum(scores) / len(scores) # This will be set by calculate_confidence_score
        ))
        
        return ParsedLens(
            manufacturer=manufacturer,
            series=series,
            focal_length=focal_length,
            t_stop=t_stop,
            lens_type=lens_type,
            format=format_info,
            mount=mount,
            anamorphic_spherical=anamorphic_spherical,
            anamorphic_squeeze=squeeze_factor,
            anamorphic_location=anamorphic_location,
            housing=housing,
            use_case=use_case,
            look=look,
            notes=notes,
            flare=flare_color,
            original_name=lens_name,
            needs_review=confidence < 0.6,
            confidence_score=confidence
        )

    def parse_csv(self, input_file: str, output_file: str) -> None:
        """Parse lens names from CSV file"""
        try:
            # Read input CSV
            df = pd.read_csv(input_file)
            
            if 'Lens Name' not in df.columns:
                print(f"Error: 'Lens Name' column not found in {input_file}")
                return
            
            # Parse each lens name
            parsed_lenses = []
            for _, row in df.iterrows():
                lens_name = str(row['Lens Name'])
                parsed = self.parse_lens_name(lens_name)
                parsed_lenses.append(parsed)
            
            # Convert to DataFrame
            result_df = pd.DataFrame([vars(lens) for lens in parsed_lenses])
            
            # Save to CSV
            result_df.to_csv(output_file, index=False)
            print(f"Parsed {len(parsed_lenses)} lenses and saved to {output_file}")
            
        except Exception as e:
            print(f"Error processing CSV: {e}")

def main():
    """Main function for testing"""
    parser = SimpleLensParser()
    
    # Test with some sample lens names
    test_lenses = [
        "Canon 6.6-66mm T2.5 Zoom",
        "Cooke S4/i 18mm T2.0",
        "Angenieux Optimo Ultra 24-290mm T2.8",
        "Zeiss Master Prime 50mm T1.3",
        "Laowa Nanomorph 1.5xS35 27mm",
        "Canon Rangefinder - TLS - LPL Mount",
        "100mm/150mm Caldwell Chameleon SC/XC - Rear Expander",
        "30-90mm Angenieux EZ-1 T2 - Rear (S35)"
    ]
    
    print("Testing lens parser:")
    print("=" * 50)
    
    for lens_name in test_lenses:
        parsed = parser.parse_lens_name(lens_name)
        print(f"\nOriginal: {lens_name}")
        print(f"Manufacturer: {parsed.manufacturer}")
        print(f"Series: {parsed.series}")
        print(f"Focal Length: {parsed.focal_length}")
        print(f"T-Stop: {parsed.t_stop}")
        print(f"Type: {parsed.lens_type}")
        print(f"Format: {parsed.format}")
        print(f"Mount: {parsed.mount}")
        print(f"Confidence: {parsed.confidence_score:.3f}")
        print(f"Needs Review: {parsed.needs_review}")

if __name__ == "__main__":
    main() 