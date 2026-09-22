# -*- coding: utf-8 -*-
"""
Rich-text and manual markup inspector for Excel spreadsheets.

Detects font colors (RGB, indexed, or theme) and font strikethrough in cells.
Provides methods to extract effective text (stripping struck-through deletions)
and to detect pre-existing manual reviewer markup in submittals.
"""

import re
import pandas as pd
import openpyxl
from openpyxl.styles.colors import Color
from openpyxl.cell.rich_text import CellRichText, TextBlock
import theme_color as tc

INDEX_COLORS = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF',
    'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF',
    'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF',
    '00FFFF', '800000', '008000', '000080', '808000',
    '800080', '008080', 'C0C0C0', '808080', '9999FF',
    '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080',
    '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00',
    '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF',
    'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC',
    '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699',
    '969696', '003366', '339966', '003300', '333300',
    '993300', '993366', '333399', '333333'
]


class StringChecker:
    """Inspects cell-level rich text, font styles, and strikethroughs."""

    def __init__(self, workbook=None, file_path=None):
        if workbook is not None:
            self.wb = workbook
        elif file_path is not None:
            self.wb = openpyxl.load_workbook(file_path, rich_text=True)
        else:
            self.wb = None
        self.theme_xml = getattr(self.wb, 'loaded_theme', None) if self.wb else None

    def get_font_color_hex(self, font_or_cell):
        """Extract a 6-character hex RGB color string from a Font or Cell object."""
        color_obj = getattr(font_or_cell, 'color', None)
        if color_obj is None or not isinstance(color_obj, Color):
            return "000000"

        if color_obj.type == 'rgb':
            rgb = str(color_obj.rgb or '')
            return rgb[-6:].upper() if len(rgb) >= 6 else "000000"
        elif color_obj.type == 'indexed':
            idx = color_obj.indexed
            if idx is not None and 0 <= idx < len(INDEX_COLORS):
                return INDEX_COLORS[idx]
        elif color_obj.type == 'theme':
            if self.theme_xml is not None and color_obj.theme is not None:
                tint = getattr(color_obj, 'tint', 0.0) or 0.0
                return tc.theme_and_tint_to_rgb(self.theme_xml, color_obj.theme, tint)
        return "000000"

    def classify_color(self, hex_color):
        """Classify hex color as 'red', 'green', 'black', or 'other'."""
        s = hex_color.lstrip('#')
        if len(s) < 6:
            return 'black'
        try:
            r = int(s[0:2], 16)
            g = int(s[2:4], 16)
            b = int(s[4:6], 16)
        except ValueError:
            return 'black'

        if r > 160 and g < 100 and b < 100:
            return 'red'
        if g > 130 and r < 100 and b < 100:
            return 'green'
        if r < 80 and g < 80 and b < 80:
            return 'black'
        return 'other'

    def extract_effective_text(self, cell_value, cell_font=None, ignore_strikethrough=True):
        """
        Extract effective text from a cell value.

        If cell_value is a CellRichText and ignore_strikethrough is True,
        blocks that have font strikethrough are omitted.
        """
        if cell_value is None:
            return ""

        if not isinstance(cell_value, CellRichText):
            if ignore_strikethrough and cell_font and getattr(cell_font, 'strike', False):
                return ""
            return str(cell_value)

        parts = []
        for block in cell_value:
            if isinstance(block, TextBlock):
                font = getattr(block, 'font', None)
                if ignore_strikethrough and font and getattr(font, 'strike', False):
                    continue
                parts.append(block.text)
            else:
                if ignore_strikethrough and cell_font and getattr(cell_font, 'strike', False):
                    continue
                parts.append(str(block))

        raw = "".join(parts)
        # Normalize multiple internal spaces to a single space
        return re.sub(r'[ \t]+', ' ', raw).strip()

    def has_markup(self, cell):
        """Check if a cell contains manual markup (strikethrough or non-black font)."""
        val = cell.value
        font = cell.font
        if font and getattr(font, 'strike', False):
            return True
        if font:
            c = self.classify_color(self.get_font_color_hex(font))
            if c in ('red', 'green'):
                return True

        if isinstance(val, CellRichText):
            for block in val:
                b_font = getattr(block, 'font', None)
                if b_font:
                    if getattr(b_font, 'strike', False):
                        return True
                    if self.classify_color(self.get_font_color_hex(b_font)) in ('red', 'green'):
                        return True
        return False

