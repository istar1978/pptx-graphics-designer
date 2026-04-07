---
name: pptx-graphics-designer
description: Professional PowerPoint graphics designer for creating Visio-style and enterprise architecture diagrams. Supports complex multi-layer architectures, network topology diagrams with standard device icons, flowcharts, and deployment maps with professional styling. Now includes rich layout designs, PPT template loading, and Mermaid diagram generation.
---

# PPTX Graphics Designer

Professional diagram and chart creation for PowerPoint presentations with Visio-style and enterprise architecture quality.

## When to Use

Use this skill when:
- Creating enterprise-style multi-layer architecture diagrams (参考图风格)
- Building **network topology diagrams** with standard device icons (routers, switches, firewalls, servers, PCs)
- Creating flowcharts, system diagrams, or deployment maps in PPTX
- Generating Gantt charts, sequence diagrams, or timeline visualizations
- Creating diagrams with subtitles, color-coded layers, dashed borders, side panels
- Converting text descriptions into professional diagram objects
- Redrawing existing diagrams with improved layout and styling
- Designing rich layouts with multiple parallel arrangements (up-down, left-right, etc.) for beautiful graphic combinations
- Loading PPT templates to generate pages based on template layouts
- Generating various PPT graphics from Mermaid syntax

## Core Capabilities

### Supported Diagram Types

1. **企业级架构图 (Enterprise Architecture)** - 多层架构、带副标题、彩色编码、侧边说明面板
2. **网络拓扑图 (Network Topology)** - 标准网络设备图标、区域分组、连接关系、物理部署图
3. **流程图 (Flowcharts)** - 标准流程图、泳道图、决策树
4. **部署图 (Deployment Maps)** - 物理环境部署、云架构、基础设施图
5. **甘特图 (Gantt Charts)** - 项目计划、时间线、里程碑
6. **时序图 (Sequence Diagrams)** - 交互流程、消息序列
7. **散点图 (Scatter Plots)** - 数据分布、相关性分析
8. **Mermaid 图表 (Mermaid Diagrams)** - 支持所有 Mermaid 图表类型（流程图、时序图、甘特图、类图等），自动转换为 PPT 图形

### Network Topology Features (新增)

- **标准网络设备图标** - Router(路由器), Switch(交换机), Firewall(防火墙), Server(服务器), Database(数据库), PC(终端), Cloud(云), Internet(互联网)
- **网络区域容器** - DMZ 区、办公区、数据中心区、互联网接入区、安全管理中心等
- **连接线** - 支持实线/虚线、带标签、自定义颜色和宽度
- **Cisco 风格图标** - 专业网络设备图标样式，带箭头/插槽等细节
- **区域分组** - 支持多层嵌套区域，带标题和彩色边框
- **物理部署图** - 完整参考图风格，支持复杂网络架构

### Enterprise Architecture Features (参考图风格)

- **多层架构容器** - 每层独立背景色、边框、标题栏
- **双行文字** - 主标题 + 副标题（如"语音识别 (ASR)" + "阿里 · 腾讯 · 讯飞"）
- **彩色编码** - 每层独立配色方案（接入层蓝、服务层红、语音层黄等）
- **虚线边框** - 支持虚线边框区分逻辑组
- **侧边说明面板** - 右侧添加详细说明文字
- **多行布局** - 支持单层内多行多列项目（如业务层 2 行 4 列）
- **自动间距** - 智能计算项目宽度和间距

### Design Principles

- **Visio/企业风格** - 使用标准图形对象，专业配色方案
- **自动适应** - 图形和文字大小根据内容自动调整
- **布局优化** - 智能排列，避免重叠，保持对齐
- **主题一致** - 匹配 PPT 模板的配色和字体

### Layout and Composition Features (新增)

- **多并列排版布局** - 支持上下、左右等多个方向的并列排版，根据页面内容自动选择最美观的布局
- **多种图形组合** - 智能组合不同类型的图形（流程图+架构图、拓扑图+时序图等），创建复合页面
- **美观排版效果** - 自动计算间距、对齐和比例，确保视觉平衡和专业外观
- **响应式布局** - 根据幻灯片尺寸和内容密度调整布局，适应不同屏幕比例

## Usage

### Basic Command

```bash
python scripts/pptx_graphics.py --input <description> --output <file.pptx> --type <diagram_type>

# Network topology (新增)
python scripts/network_topology.py --input <config.json> --output <file.pptx> --title <title>

# Mermaid diagrams (新增)
python scripts/pptx_graphics.py --input <mermaid_code> --output <file.pptx> --type mermaid
```

### Parameters

| Parameter | Description |
|-----------|-------------|
| `--input` | 图形描述文本、JSON 结构（架构图/网络拓扑）或 Mermaid 代码 |
| `--output` | 输出 PPTX 文件路径 |
| `--type` | 图表类型：flowchart/architecture/network/gantt/sequence/scatter/mermaid |
| `--slide` | 目标幻灯片索引（默认新建） |
| `--template` | PPT 模板文件路径，或预设：`reference`（参考图）、`custom`（自定义）、`default`（默认） |
| `--layout` | 排版布局：auto/vertical/horizontal/grid/composite（自动选择最美观的布局） |
| `--style` | 样式主题：default/professional/colorful/minimal/enterprise |
| `--title` | 图表标题 |

### Example Usage

```bash
# 创建流程图
python scripts/pptx_graphics.py --input "用户→登录→验证→主页" --output deck.pptx --type flowchart

# 创建架构图
python scripts/pptx_graphics.py --input '{"layers":["接入层","核心层","数据层"]}' --output deck.pptx --type architecture

# 创建甘特图
python scripts/pptx_graphics.py --input tasks.json --output deck.pptx --type gantt

# 创建网络拓扑图（新增）
python scripts/network_topology.py --input topology.json --output network.pptx --title "网络拓扑图"

# 创建参考图风格的物理部署图
python scripts/network_topology.py --input physical_deployment.json --output deploy.pptx --template reference

# 从 Mermaid 代码生成图表（新增）
python scripts/pptx_graphics.py --input "graph TD; A-->B; A-->C; B-->D; C-->D;" --output mermaid.pptx --type mermaid

# 载入 PPT 模板并生成页面（新增）
python scripts/pptx_graphics.py --input "用户流程" --output template_deck.pptx --template my_template.pptx --layout composite

# 自动选择美观排版布局（新增）
python scripts/pptx_graphics.py --input '{"diagrams": ["flowchart", "architecture"]}' --output multi.pptx --layout auto
```

## Scripts

### Main Scripts

**`scripts/pptx_graphics.py`** - 主程序，支持架构图、流程图、甘特图、时序图、散点图、Mermaid 图表等

**`scripts/network_topology.py`** (新增) - 网络拓扑图生成器，支持：
- 标准网络设备图标（Router, Switch, Firewall, Server, PC, Cloud, Internet）
- 网络区域容器（DMZ, Office, Data Center, Security 等）
- 连接线和标签
- 参考图风格物理部署图

```bash
# 创建网络拓扑图
python scripts/network_topology.py --input topology.json --output network.pptx --title "网络拓扑图"

# 创建参考图风格物理部署图
python scripts/network_topology.py --input physical_deployment.json --output deploy.pptx --template reference
```

### Helper Scripts

- **`scripts/layout_engine.py`** - 自动布局引擎，支持多并列排版和复合布局
- **`scripts/style_presets.py`** - 样式预设库
- **`scripts/shape_factory.py`** - 图形对象工厂
- **`scripts/mermaid_parser.py`** (新增) - Mermaid 语法解析器，转换为 PPT 图形
- **`scripts/template_loader.py`** (新增) - PPT 模板载入器，支持基于模板布局生成页面

## Style Presets

### Color Schemes

| Style | Primary | Secondary | Accent |
|-------|---------|-----------|--------|
| default | #2980B9 | #27AE60 | #8E44AD |
| professional | #1A237E | #0D47A1 | #3949AB |
| colorful | #E74C3C | #3498DB | #F39C12 |
| minimal | #2C3E50 | #95A5A6 | #BDC3C7 |

### Shape Standards

- 流程图节点：圆角矩形 (Rounded Rectangle)
- 决策点：菱形 (Diamond)
- 数据存储：圆柱体 (Cylinder)
- 外部系统：矩形加阴影
- 箭头：带箭头直线，3pt 宽度

## References

Load these files for detailed specifications:

- **`references/shape_specs.md`** - 图形对象规格和 OOXML 细节
- **`references/layout_algorithms.md`** - 布局算法和自动排列逻辑，包括多并列排版
- **`references/color_theory.md`** - 配色理论和可访问性标准
- **`references/mermaid_integration.md`** (新增) - Mermaid 图表集成和转换规则
- **`references/template_system.md`** (新增) - PPT 模板系统和布局设计

## Common Patterns

### Flowchart Pattern

```
[开始] → [处理] → <决策> → [结束]
              ↓
         [异常处理]
```

### Layered Architecture Pattern

```
┌─────────────────────────┐
│      接入层 (4 项)        │
├─────────────────────────┤
│      核心层 (3 项)        │
├─────────────────────────┤
│      数据层 (2 项)        │
└─────────────────────────┘
```

### Network Topology Pattern

```
    [Internet]
         │
    ┌────┴────┐
    │  Firewall │
    └────┬────┘
    ────┴────┐
    │  Switch  │
    └─┬─┬─┬──┘
    │ │ │
   [S1][S2][S3]
```

### Physical Deployment Pattern (参考图)

```
┌─────────────────┐      ┌──────────────────┐
│  互联网接入区    │      │     DMZ 区        │
│   Internet      │      │ ┌──────────────┐ │
│      ☁️         │      │ │应用服务器区  │ │
│   ┌──┴──┐       │      │ │  [S][S][S]   │ │
│  FW     FW      │      │ ├──────────────┤ │
─────┬─┬─────────┘      │ │数据库储存区  │ │
      │ │                │ │  [D][D][D]   │ │
┌─────┴─┴─────────┐      │ ├──────────────┤ │
│   核心交换区     │──┼───│ │公共服务器区  │ │
│  [SW]   [SW]    │      │ │  [P][P][P]   │ │
└─────┬─┬─────────┘      │ └──────────────┘ │
      │ │                └──────────────────┘
┌─────┴─┴─────────
│   办公运维区     │
│     [SW]        │
│  ┌─┼─┼─┼─┐      │
│ [PC][PC][PC]... │
└─────────────────┘
```

### Mermaid Diagram Pattern (新增)

```
graph TD
    A[开始] --> B{判断}
    B -->|是| C[处理]
    B -->|否| D[结束]
    C --> D
```

### Composite Layout Pattern (新增)

```
┌─────────────────┬─────────────────┐
│   流程图         │   架构图         │
├─────────────────┴─────────────────┤
│           时序图                   │
└───────────────────────────────────┘
```

## Quality Checklist

Before delivering:
- [ ] 所有图形对齐且间距一致
- [ ] 文字大小适合阅读（最小 10pt）
- [ ] 颜色对比度符合可访问性标准
- [ ] 箭头方向正确且无交叉混乱
- [ ] 图形无溢出幻灯片边界
- [ ] 配色与 PPT 主题一致

## Troubleshooting

### Text Overflow

如果文字溢出形状：
1. 自动增大形状尺寸
2. 或减小字体大小（不低于 8pt）
3. 或启用文字换行

### Layout Overlap

如果图形重叠：
1. 运行布局引擎重新排列
2. 增加节点间距参数
3. 调整幻灯片尺寸或方向

### Color Mismatch

如果配色不协调：
1. 使用 `--style` 指定预设主题
2. 或从模板提取配色方案
3. 参考 `references/color_theory.md`

### Mermaid Syntax Errors (新增)

如果 Mermaid 代码解析失败：
1. 验证 Mermaid 语法正确性
2. 检查图表类型是否支持
3. 使用 `--type mermaid` 参数

### Template Loading Issues (新增)

如果模板载入失败：
1. 确保模板文件路径正确
2. 检查模板文件是否为有效 PPTX 格式
3. 使用预设模板如 `reference` 或 `default`

### Composite Layout Problems (新增)

如果复合布局不美观：
1. 使用 `--layout auto` 让系统自动选择
2. 或手动指定 `vertical`、`horizontal`、`grid`
3. 调整幻灯片比例以适应内容

## Related Skills

- **powerpoint-pptx** - 基础 PPTX 操作技能
- **design** - 视觉设计原则
- **documents** - 文档工作流
- **mermaid** - Mermaid 图表语法和渲染
- **layout-design** - 高级排版和布局设计
