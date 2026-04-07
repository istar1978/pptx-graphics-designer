#!/usr/bin/env python3
"""
Layout Engine for PPTX Graphics Designer
Handles automatic layout, positioning, and composition of diagrams
"""

from pptx.util import Inches
import math

class LayoutEngine:
    """Engine for automatic layout and positioning of diagram elements"""

    def __init__(self, slide_width=13.333, slide_height=7.5):
        self.slide_width = slide_width
        self.slide_height = slide_height
        self.margin = 1.0  # inches

    def calculate_positions(self, items, layout_type='auto', max_width=None, max_height=None):
        """
        Calculate positions for items based on layout type

        Args:
            items: List of items with width/height properties
            layout_type: 'auto', 'vertical', 'horizontal', 'grid', 'composite'
            max_width: Maximum width constraint
            max_height: Maximum height constraint

        Returns:
            List of (x, y) positions
        """
        if not items:
            return []

        if layout_type == 'auto':
            layout_type = self._choose_best_layout(items)

        if layout_type == 'vertical':
            return self._layout_vertical(items, max_width, max_height)
        elif layout_type == 'horizontal':
            return self._layout_horizontal(items, max_width, max_height)
        elif layout_type == 'grid':
            return self._layout_grid(items, max_width, max_height)
        elif layout_type == 'composite':
            return self._layout_composite(items, max_width, max_height)
        else:
            return self._layout_vertical(items, max_width, max_height)

    def _choose_best_layout(self, items):
        """Choose the best layout based on item characteristics"""
        total_area = sum(item.get('width', 1) * item.get('height', 1) for item in items)
        avg_aspect = sum(item.get('width', 1) / max(item.get('height', 1), 0.1) for item in items) / len(items)

        slide_aspect = self.slide_width / self.slide_height

        if avg_aspect > slide_aspect * 1.5:
            return 'horizontal'
        elif avg_aspect < slide_aspect * 0.5:
            return 'vertical'
        else:
            return 'grid'

    def _layout_vertical(self, items, max_width=None, max_height=None):
        """Layout items vertically"""
        positions = []
        current_y = self.margin

        max_width = max_width or (self.slide_width - 2 * self.margin)
        max_height = max_height or (self.slide_height - 2 * self.margin)

        for item in items:
            item_width = min(item.get('width', 2), max_width)
            item_height = item.get('height', 1)

            x = (self.slide_width - item_width) / 2  # Center horizontally
            y = current_y

            positions.append((x, y))
            current_y += item_height + 0.2  # Add spacing

            if current_y + item_height > max_height:
                break  # Prevent overflow

        return positions

    def _layout_horizontal(self, items, max_width=None, max_height=None):
        """Layout items horizontally"""
        positions = []
        current_x = self.margin

        max_width = max_width or (self.slide_width - 2 * self.margin)
        max_height = max_height or (self.slide_height - 2 * self.margin)

        total_width = sum(item.get('width', 2) + 0.2 for item in items) - 0.2
        if total_width > max_width:
            # Scale down
            scale = max_width / total_width
            for item in items:
                item['width'] = item.get('width', 2) * scale

        for item in items:
            item_width = item.get('width', 2)
            item_height = item.get('height', 1)

            x = current_x
            y = (self.slide_height - item_height) / 2  # Center vertically

            positions.append((x, y))
            current_x += item_width + 0.2  # Add spacing

        return positions

    def _layout_grid(self, items, max_width=None, max_height=None):
        """Layout items in a grid"""
        positions = []

        max_width = max_width or (self.slide_width - 2 * self.margin)
        max_height = max_height or (self.slide_height - 2 * self.margin)

        num_items = len(items)
        cols = math.ceil(math.sqrt(num_items))
        rows = math.ceil(num_items / cols)

        item_width = (max_width - (cols - 1) * 0.2) / cols
        item_height = (max_height - (rows - 1) * 0.2) / rows

        for i, item in enumerate(items):
            row = i // cols
            col = i % cols

            x = self.margin + col * (item_width + 0.2)
            y = self.margin + row * (item_height + 0.2)

            positions.append((x, y))

        return positions

    def _layout_composite(self, items, max_width=None, max_height=None):
        """Composite layout mixing different arrangements"""
        if len(items) <= 2:
            return self._layout_horizontal(items, max_width, max_height)
        elif len(items) <= 4:
            return self._layout_grid(items, max_width, max_height)
        else:
            # Split into header + grid
            positions = []
            if items:
                # First item as header
                positions.append((self.margin, self.margin))
                # Rest in grid
                grid_items = items[1:]
                grid_positions = self._layout_grid(grid_items,
                                                 max_width,
                                                 max_height - items[0].get('height', 1) - 0.5)
                # Offset grid positions
                offset_y = items[0].get('height', 1) + 0.5
                for x, y in grid_positions:
                    positions.append((x, y + offset_y))
            return positions

    def optimize_spacing(self, positions, items):
        """Optimize spacing to prevent overlaps and improve aesthetics"""
        if not positions or not items:
            return positions

        # Ensure minimum spacing
        min_spacing = 0.2
        optimized = []

        for i, (x, y) in enumerate(positions):
            item = items[i]
            item_width = item.get('width', 2)
            item_height = item.get('height', 1)

            # Check for overlaps with previous items
            for j, (prev_x, prev_y) in enumerate(optimized):
                prev_item = items[j]
                prev_width = prev_item.get('width', 2)
                prev_height = prev_item.get('height', 1)

                # Simple overlap detection
                if not (x + item_width + min_spacing <= prev_x or
                       prev_x + prev_width + min_spacing <= x or
                       y + item_height + min_spacing <= prev_y or
                       prev_y + prev_height + min_spacing <= y):
                    # Overlap detected, adjust position
                    y = prev_y + prev_height + min_spacing

            optimized.append((x, y))

        return optimized

    def get_available_space(self):
        """Get available space for content"""
        return {
            'width': self.slide_width - 2 * self.margin,
            'height': self.slide_height - 2 * self.margin,
            'x': self.margin,
            'y': self.margin
        }