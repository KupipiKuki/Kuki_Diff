# -*- coding: utf-8 -*-
"""
Excel Theme Color translation utilities.

Extracts Office XML theme palette and applies MS Excel tint calculations to convert
theme color indices and tint values to standard 6-character hex RGB strings.
Compatible with Python 3.9+ and openpyxl.
"""

from colorsys import rgb_to_hls, hls_to_rgb
import xml.etree.ElementTree as ET

RGBMAX = 255
HLSMAX = 240  # Excel's tint function operates on base-240 HLS


def rgb_to_ms_hls(red, green=None, blue=None):
    """Convert RGB (0-255 or hex string '[#aa]rrggbb') to Excel base-240 HLS tuple."""
    if green is None:
        if isinstance(red, str):
            s = red.lstrip('#')
            if len(s) > 6:
                s = s[-6:]
            blue = int(s[4:6], 16) / RGBMAX
            green = int(s[2:4], 16) / RGBMAX
            red = int(s[0:2], 16) / RGBMAX
        elif isinstance(red, (tuple, list)) and len(red) == 3:
            red, green, blue = [v / RGBMAX if v > 1.0 else v for v in red]
    else:
        red, green, blue = red / RGBMAX, green / RGBMAX, blue / RGBMAX

    h, l, s = rgb_to_hls(red, green, blue)
    return (int(round(h * HLSMAX)), int(round(l * HLSMAX)), int(round(s * HLSMAX)))


def ms_hls_to_rgb(hue, lightness=None, saturation=None):
    """Convert Excel base-240 HLS to (0.0-1.0) RGB float tuple."""
    if lightness is None and isinstance(hue, (tuple, list)):
        hue, lightness, saturation = hue
    return hls_to_rgb(hue / HLSMAX, lightness / HLSMAX, saturation / HLSMAX)


def rgb_to_hex(red, green=None, blue=None):
    """Convert (0.0-1.0) RGB float values to 6-character uppercase hex string 'RRGGBB'."""
    if green is None and isinstance(red, (tuple, list)):
        red, green, blue = red
    return ('%02x%02x%02x' % (
        int(round(red * RGBMAX)),
        int(round(green * RGBMAX)),
        int(round(blue * RGBMAX))
    )).upper()


def tint_luminance(tint, lum):
    """Apply Excel tint factor (-1.0 to +1.0) to base-240 luminance."""
    if tint is None or tint == 0.0:
        return lum
    if tint < 0:
        return int(round(lum * (1.0 + tint)))
    return int(round(lum * (1.0 - tint) + (HLSMAX - HLSMAX * (1.0 - tint))))


def get_theme_colors(xml_content):
    """
    Extract the 10 core theme palette hex colors from an openpyxl loaded theme XML.
    Returns list of 10 '#RRGGBB' strings [lt1, dk1, lt2, dk2, accent1..accent6].
    """
    if xml_content is None:
        return None
    try:
        if isinstance(xml_content, bytes):
            root = ET.fromstring(xml_content)
        else:
            root = ET.fromstring(str(xml_content))
    except Exception:
        return None

    clr_scheme = None
    for elem in root.iter():
        if elem.tag.endswith('clrScheme'):
            clr_scheme = elem
            break

    if clr_scheme is None:
        return None

    color_keys = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']
    colors = []
    children_by_tag = {child.tag.split('}')[-1]: child for child in clr_scheme}

    for c in color_keys:
        accent = children_by_tag.get(c)
        val = None
        if accent is not None:
            children = list(accent)
            if children:
                first = children[0]
                val = first.attrib.get('lastClr') or first.attrib.get('val')
        if val:
            colors.append(f"#{val}" if not val.startswith('#') else val)
        else:
            colors.append('#000000')
    return colors


def theme_and_tint_to_rgb(xml_content, theme_color_number, tint=0.0):
    """
    Given theme XML, theme color slot (0-9), and tint factor, return 6-character hex RGB.
    Falls back to '000000' on invalid slot or unparseable XML.
    """
    palette = get_theme_colors(xml_content)
    if not palette or theme_color_number < 0 or theme_color_number >= len(palette):
        return "000000"
    base_hex = palette[theme_color_number]
    if tint is None or tint == 0.0:
        return base_hex.lstrip('#').upper()
    h, l, s = rgb_to_ms_hls(base_hex)
    new_l = tint_luminance(tint, l)
    return rgb_to_hex(ms_hls_to_rgb(h, new_l, s))

