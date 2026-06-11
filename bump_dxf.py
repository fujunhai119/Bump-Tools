"""Bump DXF 导出核心逻辑模块。

提供 Bump 分类、Excel 解析、DXF 生成功能（不含 GUI）。
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Tuple

import ezdxf
import pandas as pd
from ezdxf import units
from ezdxf.document import Drawing
from ezdxf.enums import TextEntityAlignment
from ezdxf.layouts import Modelspace


# ==================== 常量定义 ====================


class BumpCategory(Enum):
    """Bump 分类枚举。"""

    GND = "GND"
    POWER = "Power"
    SIGNAL = "Signal"
    CUSTOM = "Custom"


class BumpShape(Enum):
    """Bump 形状枚举。"""

    SQUARE = "正方形"
    CIRCLE = "圆形"


# DXF ACI 颜色索引默认映射
DEFAULT_COLOR_MAP: Dict[BumpCategory, int] = {
    BumpCategory.GND: 7,       # 白色
    BumpCategory.POWER: 1,     # 红色
    BumpCategory.SIGNAL: 3,    # 绿色
    BumpCategory.CUSTOM: 6,     # 洋红
}

# 可选颜色列表：(ACI索引, 显示名称, 颜色预览)
ACI_COLORS: List[Tuple[int, str, str]] = [
    (1, "红色", "#FF0000"),
    (2, "黄色", "#FFFF00"),
    (3, "绿色", "#00FF00"),
    (4, "青色", "#00FFFF"),
    (5, "蓝色", "#0000FF"),
    (6, "洋红", "#FF00FF"),
    (7, "白色", "#FFFFFF"),
    (8, "深灰", "#808080"),
    (9, "浅灰", "#C0C0C0"),
    (10, "深红", "#800000"),
    (11, "深黄", "#808000"),
    (12, "深绿", "#008000"),
    (13, "深青", "#008080"),
    (14, "深蓝", "#000080"),
    (15, "深洋红", "#800080"),
    (30, "橙色", "#FF8000"),
    (31, "棕色", "#A52A2A"),
    (140, "天蓝", "#00BFFF"),
    (200, "紫罗兰", "#8B00FF"),
    (210, "粉红", "#FF1493"),
    (250, "深棕", "#8B4513"),
]

# 内置颜色查找表（优化二次查询）
_ACI_HEX_MAP: Dict[int, str] = {aci: hex_ for aci, _, hex_ in ACI_COLORS}
_ACI_NAME_MAP: Dict[int, str] = {aci: name for aci, name, _ in ACI_COLORS}

# 中文类别名映射
CATEGORY_LABELS: Dict[BumpCategory, str] = {
    BumpCategory.GND: "GND",
    BumpCategory.POWER: "Power",
    BumpCategory.SIGNAL: "Signal",
    BumpCategory.CUSTOM: "自定义",
}


# ==================== 数据模型 ====================


@dataclass(frozen=True)
class BumpDxf:
    """Bump DXF 数据模型。

    Attributes:
        name: Bump 名称。
        x: X 坐标（μm）。
        y: Y 坐标（μm）。
        category: Bump 分类。
    """

    name: str
    x: float
    y: float
    category: BumpCategory


@dataclass
class DxfExportConfig:
    """DXF 导出配置。

    Attributes:
        shape: Bump 形状（正方形或圆形）。
        size: 正方形边长或圆直径（μm）。
        layer_settings: 分类→图层名称映射。
        color_map: 分类→ACI 颜色索引映射。
    """

    shape: BumpShape = BumpShape.SQUARE
    size: float = 100.0
    layer_settings: Dict[BumpCategory, str] = field(default_factory=dict)
    color_map: Dict[BumpCategory, int] = field(
        default_factory=lambda: dict(DEFAULT_COLOR_MAP)
    )

    def __post_init__(self) -> None:
        """初始化后补全默认图层设置。"""
        if not self.layer_settings:
            self.layer_settings = {
                BumpCategory.GND: "GND",
                BumpCategory.POWER: "POWER",
                BumpCategory.SIGNAL: "SIGNAL",
                BumpCategory.CUSTOM: "CUSTOM",
            }
        if not self.color_map:
            self.color_map = dict(DEFAULT_COLOR_MAP)

    @property
    def half_size(self) -> float:
        """获取尺寸的一半（半径或半边长）。"""
        return self.size / 2.0


# ==================== 工具函数 ====================


def get_aci_hex(aci: int) -> str:
    """根据 ACI 索引获取对应的十六进制颜色值。

    Args:
        aci: ACI 颜色索引。

    Returns:
        十六进制颜色字符串，若未找到则返回默认值 "#FFFFFF"。
    """
    return _ACI_HEX_MAP.get(aci, "#FFFFFF")


def get_aci_name(aci: int) -> str:
    """根据 ACI 索引获取对应的颜色名称。

    Args:
        aci: ACI 颜色索引。

    Returns:
        颜色名称，若未找到则返回 "自定义"。
    """
    return _ACI_NAME_MAP.get(aci, "自定义")


def get_aci_display(aci: int) -> str:
    """生成 ACI 颜色的可读显示字符串。

    Args:
        aci: ACI 颜色索引。

    Returns:
        格式为 "ACI:{idx} {name}" 的字符串。
    """
    name = get_aci_name(aci)
    return f"ACI:{aci} {name}"


# ==================== Bump 分类模块 ====================


def classify_bump(name: str) -> BumpCategory:
    """对 bump 名称进行分类。

    规则：
        - "VSS"（不区分大小写）→ GND
        - 以 "VDD" 开头（不区分大小写）→ Power
        - 其他 → Signal

    Args:
        name: Bump 名称。

    Returns:
        分类结果。
    """
    if not name:
        return BumpCategory.SIGNAL

    name_upper = name.strip().upper()

    match name_upper:
        case "VSS":
            return BumpCategory.GND
        case _ if len(name_upper) >= 3 and name_upper[:3] == "VDD":
            return BumpCategory.POWER
        case _:
            return BumpCategory.SIGNAL


# ==================== Excel 解析模块 ====================


def _detect_header_row(dataframe: pd.DataFrame) -> int:
    """检测 DataFrame 的表头行位置。

    Args:
        dataframe: 待检测的 DataFrame。

    Returns:
        数据起始行索引（0 表示无表头）。
    """
    if len(dataframe) == 0:
        return 0
    try:
        float(dataframe.iloc[0, 1])
        float(dataframe.iloc[0, 2])
        return 0
    except (ValueError, TypeError):
        return 1


def _parse_bump_cell(raw_value: object) -> str:
    """解析 Excel 单元格为字符串名称。

    Args:
        raw_value: 单元格原始值。

    Returns:
        去除首尾空格后的字符串。
    """
    return str(raw_value).strip()


def _parse_coordinate(
    raw_value: object,
    bump_name: str,
    row_number: int,
    axis: str,
) -> float:
    """解析并验证坐标单元格。

    Args:
        raw_value: 单元格原始值。
        bump_name: Bump 名称（用于错误提示）。
        row_number: 当前行号。
        axis: 坐标轴标识（"X" 或 "Y"）。

    Returns:
        解析后的浮点坐标值。

    Raises:
        ValueError: 当坐标缺失或格式错误时。
    """
    if pd.isna(raw_value):
        raise ValueError(
            f"第{row_number}行：'{bump_name}' 缺少{axis}坐标，跳过"
        )
    try:
        return float(raw_value)
    except (ValueError, TypeError) as exc:
        raise ValueError(
            f"第{row_number}行：'{bump_name}' {axis}坐标格式错误，跳过"
        ) from exc


def _parse_single_row(
    row: pd.Series,
    row_number: int,
) -> BumpDxf | None:
    """解析单行数据为 BumpDxf 对象。

    Args:
        row: DataFrame 的一行。
        row_number: 原始行号（用于错误提示）。

    Returns:
        BumpDxf 对象，若为空行则返回 None。

    Raises:
        ValueError: 数据不合法时抛出。
    """
    # 检查是否为空行
    if pd.isna(row.iloc[0]) and pd.isna(row.iloc[1]) and pd.isna(row.iloc[2]):
        return None

    # 名称 (A列)
    if pd.isna(row.iloc[0]):
        raise ValueError(f"第{row_number}行：缺少Bump名称，跳过")

    name = _parse_bump_cell(row.iloc[0])
    if not name:
        return None

    # 坐标 (B列、C列)
    x = _parse_coordinate(row.iloc[1], name, row_number, "X")
    y = _parse_coordinate(row.iloc[2], name, row_number, "Y")

    category = classify_bump(name)
    return BumpDxf(name=name, x=x, y=y, category=category)


def parse_bump_excel(file_path: str) -> List[BumpDxf]:
    """解析 Bump Excel 文件。

    Sheet1 结构：
        A列：bump name（字符串）
        B列：bump X 坐标（μm）
        C列：bump Y 坐标（μm）

    Args:
        file_path: Excel 文件路径。

    Returns:
        BumpDxf 对象列表。

    Raises:
        ValueError: 当文件无法读取或无有效数据时。
    """
    try:
        dataframe = pd.read_excel(file_path, sheet_name=0, header=None)
    except Exception as exc:
        raise ValueError(f"无法读取Excel文件: {exc}") from exc

    start_row = _detect_header_row(dataframe)
    bumps: List[BumpDxf] = []
    errors: List[str] = []

    for idx in range(start_row, len(dataframe)):
        row_number = idx + 1
        try:
            bump = _parse_single_row(dataframe.iloc[idx], row_number)
            if bump is not None:
                bumps.append(bump)
        except ValueError as exc:
            errors.append(str(exc))

    _report_parse_errors(errors, bumps)

    if not bumps:
        raise ValueError("未能从Excel文件中读取到任何有效Bump数据！")

    return bumps


def _report_parse_errors(
    errors: List[str],
    bumps: List[BumpDxf],
) -> None:
    """报告解析过程中的跳行警告。

    Args:
        errors: 错误信息列表。
        bumps: 成功解析的 Bump 列表。
    """
    if not errors or not bumps:
        return
    print(f"警告：解析过程中有 {len(errors)} 行跳过：")
    for err in errors[:10]:
        print(f"  - {err}")
    if len(errors) > 10:
        print(f"  ... 还有 {len(errors) - 10} 个错误")


# ==================== DXF 生成模块 ====================


def _create_dxf_document() -> Drawing:
    """创建并初始化 DXF 文档。

    Returns:
        已设置单位的 DXF 文档对象。
    """
    doc = ezdxf.new(dxfversion="R2010")
    doc.units = units.MM
    doc.header["$DWGCODEPAGE"] = "UTF-8"
    try:
        doc.styles.add("ZH_CN", font="simsun.ttc")
    except (OSError, ezdxf.DXFError):
        # 非 Windows 平台或字体不存在时静默跳过
        pass
    return doc


def _setup_dxf_layers(
    doc: Drawing,
    config: DxfExportConfig,
) -> Dict[BumpCategory, Dict[str, str | int]]:
    """创建 DXF 图层并返回图层配置。

    Args:
        doc: DXF 文档对象。
        config: 导出配置。

    Returns:
        分类 → 图层属性映射。
    """
    layer_defs = {
        BumpCategory.GND: {
            "name": "GND",
            "color": config.color_map[BumpCategory.GND],
        },
        BumpCategory.POWER: {
            "name": "POWER",
            "color": config.color_map[BumpCategory.POWER],
        },
        BumpCategory.SIGNAL: {
            "name": "SIGNAL",
            "color": config.color_map[BumpCategory.SIGNAL],
        },
        BumpCategory.CUSTOM: {
            "name": "CUSTOM",
            "color": config.color_map[BumpCategory.CUSTOM],
        },
    }
    for cat, defn in layer_defs.items():
        layer = doc.layers.add(name=str(defn["name"]), color=int(defn["color"]))
        layer.dxf.lineweight = 18  # 0.18mm
    return layer_defs


def _compute_square_points(
    center_x: float,
    center_y: float,
    half: float,
) -> List[Tuple[float, float]]:
    """计算正方形的四个顶点坐标。

    Args:
        center_x: 中心 X 坐标。
        center_y: 中心 Y 坐标。
        half: 半边长。

    Returns:
        按左下→右下→右上→左上顺序排列的四个点。
    """
    return [
        (center_x - half, center_y - half),
        (center_x + half, center_y - half),
        (center_x + half, center_y + half),
        (center_x - half, center_y + half),
    ]


def _draw_square_bump(
    msp: Modelspace,
    bump: BumpDxf,
    half: float,
    layer_name: str,
    color_idx: int,
) -> None:
    """在模型空间绘制正方形 Bump（含填充）。

    Args:
        msp: DXF 模型空间。
        bump: Bump 数据。
        half: 半边长。
        layer_name: 图层名称。
        color_idx: ACI 颜色索引。
    """
    points = _compute_square_points(bump.x, bump.y, half)
    attribs = {"layer": layer_name, "color": color_idx}
    msp.add_lwpolyline(points, close=True, dxfattribs=attribs)
    hatch = msp.add_hatch(color=color_idx, dxfattribs={"layer": layer_name})
    hatch.paths.add_polyline_path(points, is_closed=True)


def _draw_circle_bump(
    msp: Modelspace,
    bump: BumpDxf,
    half: float,
    layer_name: str,
    color_idx: int,
) -> None:
    """在模型空间绘制圆形 Bump（含填充）。

    Args:
        msp: DXF 模型空间。
        bump: Bump 数据。
        half: 半径。
        layer_name: 图层名称。
        color_idx: ACI 颜色索引。
    """
    attribs = {"layer": layer_name, "color": color_idx}
    msp.add_circle(center=(bump.x, bump.y), radius=half, dxfattribs=attribs)
    hatch = msp.add_hatch(color=color_idx, dxfattribs={"layer": layer_name})
    edge = hatch.paths.add_edge_path()
    edge.add_ellipse(
        center=(bump.x, bump.y),
        major_axis=(half, 0),
        ratio=1.0,
        start_angle=0.0,
        end_angle=2.0 * math.pi,
    )


def _draw_single_bump(
    msp: Modelspace,
    bump: BumpDxf,
    config: DxfExportConfig,
    layer_config: Dict[BumpCategory, Dict[str, str | int]],
) -> None:
    """根据配置绘制单个 Bump。

    Args:
        msp: DXF 模型空间。
        bump: Bump 数据。
        config: 导出配置。
        layer_config: 图层配置映射。
    """
    layer_name = str(layer_config[bump.category]["name"])
    color_idx = int(layer_config[bump.category]["color"])
    half = config.half_size

    match config.shape:
        case BumpShape.SQUARE:
            _draw_square_bump(msp, bump, half, layer_name, color_idx)
        case BumpShape.CIRCLE:
            _draw_circle_bump(msp, bump, half, layer_name, color_idx)


# ==================== 图例渲染模块 ====================


def _compute_bump_bounds(
    bumps: List[BumpDxf],
) -> Tuple[float, float] | None:
    """计算 Bump 列表的坐标边界。

    Args:
        bumps: Bump 数据列表。

    Returns:
        (max_x, min_y) 元组，若列表为空则返回 None。
    """
    if not bumps:
        return None
    all_x = [b.x for b in bumps]
    all_y = [b.y for b in bumps]
    return max(all_x), min(all_y)


def _draw_legend_title(
    msp: Modelspace,
    origin_x: float,
    origin_y: float,
    text_height: float,
) -> None:
    """绘制图例标题。

    Args:
        msp: DXF 模型空间。
        origin_x: 左下角 X 坐标。
        origin_y: 图例标题 Y 坐标。
        text_height: 文字高度。
    """
    msp.add_text(
        "Bump DXF 图例",
        dxfattribs={
            "layer": "0", "height": text_height * 1.2,
            "style": "ZH_CN",
        },
    ).set_placement((origin_x, origin_y), align=TextEntityAlignment.LEFT)


def _draw_legend_color_block(
    msp: Modelspace,
    block_cx: float,
    block_cy: float,
    block_half: float,
    color_idx: int,
) -> None:
    """绘制图例色块（带填充的正方形）。

    Args:
        msp: DXF 模型空间。
        block_cx: 色块中心 X。
        block_cy: 色块中心 Y。
        block_half: 色块半边长。
        color_idx: ACI 颜色索引。
    """
    points = [
        (block_cx - block_half, block_cy - block_half),
        (block_cx + block_half, block_cy - block_half),
        (block_cx + block_half, block_cy + block_half),
        (block_cx - block_half, block_cy + block_half),
    ]
    msp.add_lwpolyline(
        points, close=True, dxfattribs={"layer": "0", "color": color_idx}
    )
    hatch = msp.add_hatch(color=color_idx, dxfattribs={"layer": "0"})
    hatch.paths.add_polyline_path(points, is_closed=True)


def _draw_legend_item(
    msp: Modelspace,
    category: BumpCategory,
    color_idx: int,
    count: int,
    legend_x: float,
    label_y: float,
    config: DxfExportConfig,
    text_height: float,
) -> None:
    """绘制单个分类的图例条目（色块+文字）。

    Args:
        msp: DXF 模型空间。
        category: Bump 分类。
        color_idx: 颜色索引。
        count: 该分类的 Bump 数量。
        legend_x: 文字起始 X 坐标。
        label_y: 条目 Y 坐标。
        config: 导出配置。
        text_height: 文字高度。
    """
    block_size = config.size * 0.5
    block_cx = legend_x - config.size
    block_half = block_size / 2

    _draw_legend_color_block(msp, block_cx, label_y, block_half, color_idx)

    label_text = f"{CATEGORY_LABELS[category]} (n={count})"
    msp.add_text(
        label_text,
        dxfattribs={"layer": "0", "height": text_height, "style": "ZH_CN"},
    ).set_placement(
        (legend_x, label_y - text_height * 0.3),
        align=TextEntityAlignment.LEFT,
    )


def _draw_legend_info(
    msp: Modelspace,
    origin_x: float,
    info_y: float,
    config: DxfExportConfig,
    total_count: int,
    text_height: float,
    line_spacing: float,
) -> None:
    """绘制图例底部的参数说明。

    Args:
        msp: DXF 模型空间。
        origin_x: 文字起始 X。
        info_y: 起始 Y。
        config: 导出配置。
        total_count: Bump 总数。
        text_height: 文字高度。
        line_spacing: 图例行间距。
    """
    shape_text = "正方形" if config.shape == BumpShape.SQUARE else "圆形"
    size_text = (
        f"边长 {config.size}um"
        if config.shape == BumpShape.SQUARE
        else f"直径 {config.size}um"
    )
    info_lines = [
        f"形状: {shape_text}",
        f"尺寸: {size_text}",
        f"Bump总计: {total_count}个",
    ]
    for j, line in enumerate(info_lines):
        msp.add_text(
            line,
            dxfattribs={
                "layer": "0", "height": text_height * 0.8, "style": "ZH_CN",
            },
        ).set_placement(
            (origin_x, info_y - j * line_spacing * 0.7),
            align=TextEntityAlignment.LEFT,
        )


def _add_legend(
    msp: Modelspace,
    bumps: List[BumpDxf],
    config: DxfExportConfig,
) -> None:
    """在 DXF 中添加图例说明（标题+色块+统计+参数）。

    Args:
        msp: DXF 模型空间。
        bumps: Bump 数据列表。
        config: 导出配置。
    """
    bounds = _compute_bump_bounds(bumps)
    if bounds is None:
        return

    max_x, min_y = bounds
    legend_x = max_x + config.size * 3
    line_spacing = config.size * 1.5
    text_height = config.size * 0.6

    title_y = min_y + line_spacing * 5
    _draw_legend_title(msp, legend_x, title_y, text_height)

    legend_items = [
        (BumpCategory.GND, config.color_map[BumpCategory.GND]),
        (BumpCategory.POWER, config.color_map[BumpCategory.POWER]),
        (BumpCategory.SIGNAL, config.color_map[BumpCategory.SIGNAL]),
        (BumpCategory.CUSTOM, config.color_map[BumpCategory.CUSTOM]),
    ]
    for i, (cat, color_idx) in enumerate(legend_items):
        count = sum(1 for b in bumps if b.category == cat)
        cat_y = title_y - line_spacing * (i + 1)
        _draw_legend_item(
            msp, cat, color_idx, count, legend_x, cat_y, config, text_height
        )

    info_y = title_y - line_spacing * len(legend_items) - line_spacing
    _add_legend_info(
        msp, legend_x, info_y, config, len(bumps), text_height, line_spacing
    )


# ==================== 完整 DXF 导出 ====================


def create_bump_dxf(
    bumps: List[BumpDxf],
    config: DxfExportConfig,
    output_path: str,
) -> str:
    """生成 Bump DXF 文件。

    Args:
        bumps: Bump 数据列表。
        config: 导出配置（形状、尺寸、颜色）。
        output_path: 输出文件路径。

    Returns:
        输出文件路径。
    """
    doc = _create_dxf_document()
    layer_config = _setup_dxf_layers(doc, config)
    msp = doc.modelspace()

    for bump in bumps:
        _draw_single_bump(msp, bump, config, layer_config)

    _add_legend(msp, bumps, config)
    doc.saveas(output_path)
    return output_path
