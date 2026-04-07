# Enterprise Architecture Examples

参考图风格的企业级架构图示例。

## JSON 格式

```json
[
  {
    "name": "接入层",
    "color": "access",
    "bg_color": "bg_access",
    "items": [
      {"text": "400 热线"},
      {"text": "官网客服"},
      {"text": "意图识别", "subtitle": "BERT · Transformer"}
    ]
  },
  {
    "name": "语音处理层",
    "color": "voice",
    "bg_color": "bg_voice",
    "items": [
      {"text": "语音识别 (ASR)", "subtitle": "阿里 · 腾讯 · 讯飞 · 百度"},
      {"text": "语音合成 (TTS)", "subtitle": "阿里 · 腾讯 · 讯飞 · 百度 · 美团"}
    ]
  }
]
```

## 配色方案

### 标准企业配色

| 层级 | 颜色 | 背景色 | 用途 |
|------|------|--------|------|
| access | #4285F4 | #E8F0FE | 接入层 |
| service | #EA4335 | #FDEDEC | 服务层 |
| voice | #FBBC05 | #FEF9E3 | 语音处理 |
| nlp | #8E44AD | #F5E8FF | NLP/意图 |
| dialog | #3498DB | #E1F5FE | 对话引擎 |
| business | #E74C3C | #FDEBEA | 业务处理 |
| data | #95A5A6 | #F0F0F0 | 数据层 |
| infra | #9B59B6 | #F3ECFF | 基础设施 |

## 使用示例

### 命令行

```bash
python scripts/pptx_graphics.py \
  --input '{"name":"接入层","items":[{"text":"400 热线"},{"text":"官网客服"}]}' \
  --output arch.pptx \
  --type enterprise \
  --style enterprise \
  --title "技术架构"
```

### Python API

```python
from pptx_graphics import generate_enterprise_architecture

layers = [
    {
        'name': '接入层',
        'color': 'access',
        'bg_color': 'bg_access',
        'items': [
            {'text': '400 热线'},
            {'text': '官网客服', 'subtitle': 'Web'},
        ]
    },
    {
        'name': '语音处理层',
        'color': 'voice',
        'bg_color': 'bg_voice',
        'items': [
            {'text': 'ASR', 'subtitle': '阿里 · 腾讯'},
            {'text': 'TTS', 'subtitle': '讯飞 · 百度'},
        ]
    }
]

side_panel = [
    {
        'title': '1. 服务接入',
        'items': ['说明文字 1', '说明文字 2']
    }
]

generate_enterprise_architecture(
    'output.pptx',
    layers,
    side_panel,
    title='技术架构',
    style='enterprise'
)
```

## 布局参数

### 默认尺寸

- 幻灯片：13.333" x 7.5" (16:9)
- 内容宽度：10.0" (留 3.3" 给侧边面板)
- 层高度：0.95"
- 层间距：0.12"
- 项目高度：0.65"
- 项目间距：0.15"

### 自动计算

项目宽度根据内容自动计算：
```
item_width = (layer_width - margin - (num_items - 1) * gap) / num_items
```

## 高级功能

### 多行布局

业务处理层等特殊场景支持多行：

```python
# 2 行 4 列布局
biz_items = [
    [item1, item2, item3, item4],  # 第一行
    [item5, item6, item7, item8],  # 第二行
]
```

### 虚线边框

逻辑分组使用虚线边框：

```json
{
  "name": "智能对话引擎",
  "dashed": true,
  "items": [...]
}
```

### 自定义配色

```python
colors = {
    'custom': RGBColor(100, 150, 200),
    'bg_custom': RGBColor(240, 245, 250),
}
```
