#!/usr/bin/env python3
"""
Style Presets Library for PPTX Graphics Designer
Contains color schemes, themes, and styling configurations
"""

from pptx.dml.color import RGBColor

# ============================================================================
# Enhanced Color Schemes (参考图配色)
# ============================================================================

COLOR_SCHEMES = {
    # 参考图配色方案
    'enterprise': {
        'access': RGBColor(66, 133, 244),      # 蓝色 - 接入层
        'service': RGBColor(234, 67, 53),       # 红色 - 服务层
        'voice': RGBColor(251, 188, 5),         # 黄色 - 语音层
        'nlp': RGBColor(142, 68, 173),          # 紫色 - NLP 层
        'dialog': RGBColor(52, 152, 219),       # 青色 - 对话层
        'business': RGBColor(231, 76, 60),      # 红色 - 业务层
        'data': RGBColor(149, 165, 166),        # 灰色 - 数据层
        'infra': RGBColor(155, 89, 182),        # 紫色 - 基础设施
        'bg_access': RGBColor(232, 240, 254),   # 浅蓝背景
        'bg_service': RGBColor(253, 237, 236),  # 浅红背景
        'bg_voice': RGBColor(254, 249, 227),    # 浅黄背景
        'bg_nlp': RGBColor(245, 232, 255),      # 浅紫背景
        'bg_dialog': RGBColor(225, 245, 254),   # 浅青背景
        'bg_business': RGBColor(253, 235, 234), # 浅红背景
        'bg_data': RGBColor(240, 240, 240),     # 浅灰背景
        'bg_infra': RGBColor(243, 236, 255),    # 浅紫背景
        'text_dark': RGBColor(51, 51, 51),
        'text_light': RGBColor(102, 102, 102),
        'border': RGBColor(150, 150, 150),
        'white': RGBColor(255, 255, 255),
    },
    'default': {
        'primary': RGBColor(41, 128, 185),
        'secondary': RGBColor(39, 174, 96),
        'accent': RGBColor(142, 68, 173),
        'warning': RGBColor(230, 126, 34),
        'dark': RGBColor(44, 62, 80),
        'light': RGBColor(236, 240, 241),
        'white': RGBColor(255, 255, 255),
        'border': RGBColor(52, 73, 94),
        'text_dark': RGBColor(51, 51, 51),
        'text_light': RGBColor(102, 102, 102),
    },
    'professional': {
        'primary': RGBColor(26, 35, 126),
        'secondary': RGBColor(13, 71, 161),
        'accent': RGBColor(57, 73, 171),
        'warning': RGBColor(192, 57, 43),
        'dark': RGBColor(33, 33, 33),
        'light': RGBColor(245, 245, 245),
        'white': RGBColor(255, 255, 255),
        'border': RGBColor(0, 0, 0),
        'text_dark': RGBColor(51, 51, 51),
        'text_light': RGBColor(102, 102, 102),
    },
    'colorful': {
        'primary': RGBColor(231, 76, 60),
        'secondary': RGBColor(52, 152, 219),
        'accent': RGBColor(243, 156, 18),
        'warning': RGBColor(230, 126, 34),
        'dark': RGBColor(44, 62, 80),
        'light': RGBColor(236, 240, 241),
        'white': RGBColor(255, 255, 255),
        'border': RGBColor(149, 165, 166),
        'text_dark': RGBColor(51, 51, 51),
        'text_light': RGBColor(102, 102, 102),
    },
    'minimal': {
        'primary': RGBColor(44, 62, 80),
        'secondary': RGBColor(149, 165, 166),
        'accent': RGBColor(189, 195, 199),
        'warning': RGBColor(230, 126, 34),
        'dark': RGBColor(44, 62, 80),
        'light': RGBColor(236, 240, 241),
        'white': RGBColor(255, 255, 255),
        'border': RGBColor(189, 195, 199),
        'text_dark': RGBColor(51, 51, 51),
        'text_light': RGBColor(102, 102, 102),
    },
}

# ============================================================================
# Font Configurations
# ============================================================================

FONT_PRESETS = {
    'default': {
        'family': 'Arial',
        'size': 12,
        'bold': False,
        'italic': False,
    },
    'title': {
        'family': 'Arial',
        'size': 24,
        'bold': True,
        'italic': False,
    },
    'subtitle': {
        'family': 'Arial',
        'size': 14,
        'bold': False,
        'italic': False,
    },
    'caption': {
        'family': 'Arial',
        'size': 10,
        'bold': False,
        'italic': False,
    },
}

# ============================================================================
# Layout Presets
# ============================================================================

LAYOUT_PRESETS = {
    'single': {
        'margin_left': 1.0,
        'margin_right': 1.0,
        'margin_top': 1.0,
        'margin_bottom': 1.0,
        'columns': 1,
        'rows': 1,
    },
    'two_column': {
        'margin_left': 0.5,
        'margin_right': 0.5,
        'margin_top': 1.0,
        'margin_bottom': 1.0,
        'columns': 2,
        'rows': 1,
        'gutter': 0.5,
    },
    'three_column': {
        'margin_left': 0.3,
        'margin_right': 0.3,
        'margin_top': 1.0,
        'margin_bottom': 1.0,
        'columns': 3,
        'rows': 1,
        'gutter': 0.3,
    },
    'grid_2x2': {
        'margin_left': 0.5,
        'margin_right': 0.5,
        'margin_top': 1.0,
        'margin_bottom': 1.0,
        'columns': 2,
        'rows': 2,
        'gutter': 0.3,
    },
}

def get_color_scheme(style_name):
    """Get color scheme by name"""
    return COLOR_SCHEMES.get(style_name, COLOR_SCHEMES['default'])

def get_font_preset(preset_name):
    """Get font preset by name"""
    return FONT_PRESETS.get(preset_name, FONT_PRESETS['default'])

def get_layout_preset(layout_name):
    """Get layout preset by name"""
    return LAYOUT_PRESETS.get(layout_name, LAYOUT_PRESETS['single'])