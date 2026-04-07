# pptx-graphics-designer 技能更新日志

## v1.1.0 (2026-03-30) - 网络拓扑图功能增强

### 新增功能

#### 1. 网络拓扑图生成器 (`scripts/network_topology.py`)

全新的网络拓扑图生成脚本，支持创建 Visio 风格的企业网络架构图：

**设备图标库** (13 种标准网络设备)：
- `router` - 路由器（椭圆形带⇄符号，蓝色）
- `switch` - 交换机（椭圆形带⇅符号，蓝色）
- `firewall` - 防火墙（立方体，橙色）
- `load_balancer` - 负载均衡器（六边形，紫色）
- `server` - 服务器（立方体带≡符号，灰色）
- `database` - 数据库（圆柱体，绿色）
- `storage` - 存储设备（圆柱体，蓝色）
- `pc` - 台式机（矩形，浅蓝色）
- `laptop` - 笔记本电脑（矩形，浅蓝色）
- `phone` - 移动设备（圆角矩形，紫色）
- `cloud` - 云服务（云朵形，蓝色）
- `internet` - 互联网（椭圆形带🌐符号，蓝色）
- `datacenter` - 数据中心（矩形，灰色）

**区域容器** (7 种预定义类型)：
- `internet` - 互联网接入区（浅蓝背景，蓝色边框）
- `dmz` - DMZ 隔离区（浅橙背景，橙色边框）
- `office` - 办公网络区（浅绿背景，绿色边框）
- `datacenter` - 数据中心区（浅紫背景，紫色边框）
- `security` - 安全管理中心（浅红背景，红色边框）
- `core` - 核心交换区（浅蓝背景，蓝色边框）
- `custom` - 自定义区域（浅灰背景，灰色边框）

**连接功能**：
- 实线/虚线连接
- 连接标签（如"MPLS"、"专线"）
- 自定义颜色
- 自定义线宽

#### 2. 示例配置文件

- `examples/physical_deployment.json` - 参考图风格物理环境部署图
  - 互联网接入区（Internet + 双边界防火墙）
  - 核心交换区（双核心交换机）
  - 办公运维区（接入交换机 + 5 个终端）
  - DMZ 区（应用/数据库/公共服务器各 3 台）
  - 安全管理中心（5 套审计/管理系统）
  - 完整连接关系（30+ 条连接线）

- `examples/enterprise_network.json` - 企业网络拓扑图
  - 云端服务（负载均衡 + 云服务）
  - 总部数据中心（路由器 + 双防火墙 + 核心交换 + 应用/数据库）
  - 分支机构 1（路由器 + 交换机 + 2PC）
  - 分支机构 2（路由器 + 交换机 + 2PC）
  - MPLS 专线连接

#### 3. 文档

- `examples/README.md` - 示例说明和使用指南
- `references/network_topology_guide.md` - 完整的网络拓扑图生成指南
  - 设备图标说明
  - 区域类型说明
  - JSON 配置结构
  - 坐标系统
  - 布局建议
  - 最佳实践
  - 故障排除

### 改进功能

#### 1. SKILL.md 更新

- 更新技能描述，突出网络拓扑图功能
- 添加"网络拓扑图"到支持的图表类型
- 新增"Network Topology Features"章节
- 更新使用示例，添加网络拓扑图命令
- 更新 Scripts 章节，说明两个主脚本的用途
- 添加"Physical Deployment Pattern"示意图

### 技术细节

#### 图标绘制

- 使用 python-pptx 标准形状（MSO_SHAPE）
- 为设备图标添加符号细节：
  - Router: ⇄（双向箭头）
  - Switch: ⇅（上下箭头）
  - Server: ≡（服务器插槽）
  - Internet: 🌐（地球图标）
- 支持自定义颜色和尺寸

#### 区域容器

- 使用矩形背景 + 右上角标题
- 预定义配色方案确保视觉一致性
- 支持自定义区域类型

#### 连接线

- 使用 MSO_CONNECTOR.STRAIGHT 直线连接
- 自动计算设备中心点
- 支持中点标签（带背景框）
- 支持虚线（dash_style = 4）

### 使用示例

```bash
# 生成参考图风格物理部署图
python scripts/network_topology.py \
  --input examples/physical_deployment.json \
  --output examples/physical_deployment.pptx \
  --title "物理环境部署图"

# 生成企业网络拓扑图
python scripts/network_topology.py \
  --input examples/enterprise_network.json \
  --output examples/enterprise_network.pptx \
  --title "企业网络拓扑图"

# 使用自定义配置
python scripts/network_topology.py \
  --input my_topology.json \
  --output my_topology.pptx \
  --title "我的网络拓扑"
```

### 输出示例

- `examples/physical_deployment.pptx` (31 KB) - 物理环境部署图
- `examples/enterprise_network.pptx` (31 KB) - 企业网络拓扑图

### 兼容性

- Python 3.7+
- python-pptx 0.6.18+
- 输出格式：PowerPoint 2007+ (.pptx)

### 已知限制

1. 图标形状受限于 python-pptx 支持的 MSO_SHAPE
2. 复杂曲线连接需要手动调整
3. 不支持图标导入（如 PNG/SVG）
4. 连接线自动避让功能有限

### 未来计划

- [ ] 支持更多设备图标类型
- [ ] 添加图标导入功能（PNG/SVG）
- [ ] 智能连接线避让算法
- [ ] 自动布局优化
- [ ] 支持更多区域模板
- [ ] 添加动画效果支持
- [ ] 支持分组折叠/展开

---

## v1.0.0 (2026-03-28) - 初始版本

### 功能

- 企业级架构图生成
- 多层架构容器
- 双行文字（主标题 + 副标题）
- 彩色编码
- 虚线边框
- 侧边说明面板
- 流程图生成
- 甘特图生成
- 样式预设

### 脚本

- `scripts/pptx_graphics.py` - 主程序

### 文档

- `SKILL.md` - 技能说明
- `references/shape_specs.md` - 图形规格
- `references/enterprise_examples.md` - 企业示例

---

## 版本规范

- 主版本号：重大功能更新（如新增网络拓扑图）
- 次版本号：功能增强和改进
- 修订号：bug 修复和文档更新
