# Bump 工具箱

整合 Bump 版本比较 和 Bump DXF 生成 两个功能模块的统一工具。

## 功能模块

### Tab 1：Bump 版本比较
- 选择含两个 Sheet 的 Excel 文件（Sheet1=新版，Sheet2=旧版）
- 比较两个版本 Bump 差异（相同/不同/删除/新增）
- 生成 Excel 报告文件

### Tab 2：Bump DXF 生成
- 读取 Excel 中 Bump 坐标数据（A/B/C 列 = 名称/X/Y 坐标）
- 自动分类：VSS → GND，VDD 开头 → Power，其他 → Signal
- 配置 Bump 形状（正方形/圆形）和尺寸
- 自定义各类别显示颜色（ACI 颜色索引）
- 生成彩色 DXF 文件（含图例），可用 AutoCAD、DraftSight 等打开

## 环境要求

| 依赖 | 版本 |
|------|------|
| Python | ≥ 3.10 |
| pandas | ≥ 2.0.0 |
| openpyxl | ≥ 3.1.0 |
| ezdxf | ≥ 1.0.0 |

## 安装依赖

```bash
pip install -r requirements.txt
```

## 启动方式

```bash
python main.py
```

## Excel 数据格式要求

### 版本比较模块（Sheet1 + Sheet2）
| 列 | 内容 | 说明 |
|----|------|------|
| **A 列** | Bump 名称 | 字符串 |
| **B 列** | X 坐标 | 数值 |
| **C 列** | Y 坐标 | 数值 |

### DXF 生成模块（Sheet1）
| 列 | 内容 | 说明 |
|----|------|------|
| **A 列** | Bump 名称 | 字符串，如 `VSS`、`VDD1`、`IO1` |
| **B 列** | X 坐标 | 数值，单位 μm |
| **C 列** | Y 坐标 | 数值，单位 μm |

## 输出文件

- 版本比较：Excel 报告（含 Summary / Same_Bumps / Deleted_Bumps / Added_Bumps 四个 Sheet）
- DXF 生成：DXF 文件（含 GND / POWER / SIGNAL / CUSTOM 四个图层）
