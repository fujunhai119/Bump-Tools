"""Bump List 分析核心模块。

提供 bump 数据读取、分类、差分信号识别、最小距离计算(KD-Tree)和报告导出功能。
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Tuple, TypedDict

import numpy as np
import pandas as pd
from scipy.spatial import cKDTree

# ---------------------------------------------------------------------------
# 常量
# ---------------------------------------------------------------------------

DIFF_CHAR_PAIRS: frozenset[Tuple[str, str]] = frozenset(
    {
        ("P", "N"),
        ("N", "P"),
        ("L", "H"),
        ("H", "L"),
    }
)
"""差分信号字符配对集合。两个 bump 名中唯一不同的字符属于此集合时，判定为差分信号。"""


# ---------------------------------------------------------------------------
# 数据类
# ---------------------------------------------------------------------------


@dataclass
class BumpInfo:
    """单个 bump 数据。

    Attributes:
        name: bump 名称。
        x: x 坐标。
        y: y 坐标。
        category: bump 分类，取值为 power / power_sense / vss / signal_diff / signal_single。
    """

    name: str
    x: float
    y: float
    category: str = ""


class BumpAnalysisResult(TypedDict):
    """Bump 分析结果汇总。"""

    total_count: int
    power_count: int
    power_sense_count: int
    vss_count: int
    signal_total_count: int
    diff_pair_count: int  # 差分对数量（一对计为 1）
    single_end_count: int  # 单端信号 bump 数量
    min_distance: float
    bumps: List[BumpInfo]
    diff_pairs: List[Tuple[BumpInfo, BumpInfo]]
    power_by_name: Dict[str, int]  # 每种 power bump 名称的数量
    power_sense_by_name: Dict[str, int]  # 每种 power sense bump 名称的数量


# ---------------------------------------------------------------------------
# Excel 读取
# ---------------------------------------------------------------------------


def read_bump_list(filepath: str) -> List[BumpInfo]:
    """读取 Excel Sheet1 中的 bump 列表（A=名称，B=x，C=y）。

    Args:
        filepath: Excel 文件路径。

    Returns:
        按行顺序返回的 BumpInfo 列表。

    Raises:
        FileNotFoundError: 文件不存在时抛出。
        ValueError: Sheet 为空或坐标列包含非数值时抛出。
    """
    try:
        df = pd.read_excel(
            filepath,
            sheet_name="Sheet1",
            usecols=[0, 1, 2],
            header=None,
        )
    except FileNotFoundError:
        raise FileNotFoundError(f"文件不存在: {filepath}")
    except ValueError as e:
        if "Worksheet" in str(e) or "sheet" in str(e).lower():
            raise ValueError("Sheet1 不存在，请检查 Excel 文件。")
        raise

    if df.empty:
        raise ValueError("Sheet1 为空，请检查文件内容。")

    bumps: List[BumpInfo] = []
    for idx, row in df.iterrows():
        name = str(row[0]).strip()
        try:
            x = float(row[1])
            y = float(row[2])
        except (ValueError, TypeError):
            raise ValueError(
                f"Sheet1 第 {idx + 1} 行坐标非数值：x={row[1]}, y={row[2]}"
            )
        bumps.append(BumpInfo(name=name, x=x, y=y))

    return bumps


# ---------------------------------------------------------------------------
# Bump 分类
# ---------------------------------------------------------------------------


def _classify_single(name: str) -> str:
    """对单个 bump 名称进行分类。

    Args:
        name: bump 名称（已 strip）。

    Returns:
        分类字符串：power / power_sense / vss / signal。
    """
    name_upper = name.upper()

    if name_upper.startswith("VDD"):
        if "SENSE" in name_upper:
            return "power_sense"
        return "power"

    if name_upper in ("VSS", "GND"):
        return "vss"

    return "signal"


def classify_bumps(bumps: List[BumpInfo]) -> None:
    """对 bumps 列表中的每个 bump 进行分类，直接修改原对象。

    Args:
        bumps: BumpInfo 列表（会被就地修改 category 字段）。
    """
    for bump in bumps:
        bump.category = _classify_single(bump.name)


# ---------------------------------------------------------------------------
# 差分信号识别
# ---------------------------------------------------------------------------


def _is_diff_pair_name(name1: str, name2: str) -> bool:
    """判断两个 bump 名是否构成差分信号对。

    条件：长度相同，且恰好只有一个位置的字符不同，
    且不同的两个字符属于 DIFF_CHAR_PAIRS。

    Args:
        name1: 第一个 bump 名。
        name2: 第二个 bump 名。

    Returns:
        是否构成差分对。
    """
    if len(name1) != len(name2):
        return False

    diff_count = 0
    diff_chars: Tuple[str, str] | None = None

    for c1, c2 in zip(name1, name2):
        if c1 != c2:
            diff_count += 1
            diff_chars = (c1, c2)
            if diff_count > 1:
                return False

    if diff_count != 1 or diff_chars is None:
        return False

    return diff_chars in DIFF_CHAR_PAIRS


def find_diff_pairs(
    bumps: List[BumpInfo],
) -> List[Tuple[BumpInfo, BumpInfo]]:
    """从 signal bump 中识别差分信号对。

    使用分组优化：按名字长度分组，再用通配符 key 匹配潜在配对，
    最后精确校验字符对是否属于 DIFF_CHAR_PAIRS。

    Args:
        bumps: 已分类的 BumpInfo 列表。

    Returns:
        差分对列表，每对为 (bump1, bump2)。每个 bump 只出现在一个对中。
    """
    signal_bumps = [b for b in bumps if b.category == "signal"]

    # 按名字长度分组
    length_groups: Dict[int, List[BumpInfo]] = {}
    for b in signal_bumps:
        length_groups.setdefault(len(b.name), []).append(b)

    paired: set[int] = set()  # 已配对的 bump 索引（在 signal_bumps 中的索引）
    diff_pairs: List[Tuple[BumpInfo, BumpInfo]] = []

    for group in length_groups.values():
        # 构建通配符 key -> list of (bump, diff_position, diff_char)
        key_map: Dict[str, List[Tuple[BumpInfo, int, str]]] = {}
        for b in group:
            for i, c in enumerate(b.name):
                # 将第 i 位替换为 * 作为通配符 key
                key = b.name[:i] + "*" + b.name[i + 1 :]
                key_map.setdefault(key, []).append((b, i, c))

        # 在同一 key 下寻找合法差分对
        for key, candidates in key_map.items():
            if len(candidates) < 2:
                continue
            for i in range(len(candidates)):
                if id(candidates[i][0]) in paired:
                    continue
                for j in range(i + 1, len(candidates)):
                    if id(candidates[j][0]) in paired:
                        continue
                    b1, pos1, char1 = candidates[i]
                    b2, pos2, char2 = candidates[j]
                    if pos1 != pos2:
                        continue
                    if (char1, char2) in DIFF_CHAR_PAIRS:
                        diff_pairs.append((b1, b2))
                        paired.add(id(b1))
                        paired.add(id(b2))
                        break  # b1 已配对，跳出内层循环

    # 将配对的 signal bump 分类更新为 signal_diff
    paired_bumps: set[str] = set()
    for b1, b2 in diff_pairs:
        paired_bumps.add(id(b1))
        paired_bumps.add(id(b2))
        b1.category = "signal_diff"
        b2.category = "signal_diff"

    # 未配对的 signal bump 分类为 signal_single
    for b in signal_bumps:
        if id(b) not in paired_bumps:
            b.category = "signal_single"

    return diff_pairs


# ---------------------------------------------------------------------------
# 最小距离计算（KD-Tree）
# ---------------------------------------------------------------------------


def calc_min_distance(bumps: List[BumpInfo]) -> float:
    """使用 cKDTree 计算所有 bump 之间的最小距离。

    Args:
        bumps: BumpInfo 列表。

    Returns:
        最小距离（浮点数）。若 bump 数量 < 2，返回 0.0。
    """
    if len(bumps) < 2:
        return 0.0

    coords = np.array([[b.x, b.y] for b in bumps], dtype=np.float64)
    tree = cKDTree(coords)

    # query 找每个点的第 2 近邻（第 1 近邻是自身，distance=0）
    distances, _ = tree.query(coords, k=2)
    min_dist = float(distances[:, 1].min())
    return min_dist


# ---------------------------------------------------------------------------
# 统计汇总
# ---------------------------------------------------------------------------


def _count_by_name(
    bumps: List[BumpInfo], category: str
) -> Dict[str, int]:
    """按名称统计指定分类的 bump 数量，按数量降序排列。

    Args:
        bumps: BumpInfo 列表。
        category: 分类名。

    Returns:
        {bump名称: 数量} 字典。
    """
    counter: Dict[str, int] = {}
    for b in bumps:
        if b.category == category:
            counter[b.name] = counter.get(b.name, 0) + 1
    return dict(
        sorted(counter.items(), key=lambda item: item[1], reverse=True)
    )


def analyze_bumps(bumps: List[BumpInfo]) -> BumpAnalysisResult:
    """完整分析流程：分类 + 差分识别 + 最小距离计算。

    Args:
        bumps: 从 Excel 读取的 BumpInfo 列表。

    Returns:
        完整分析结果 BumpAnalysisResult。
    """
    classify_bumps(bumps)
    diff_pairs = find_diff_pairs(bumps)
    min_dist = calc_min_distance(bumps)

    power_count = sum(1 for b in bumps if b.category == "power")
    power_sense_count = sum(1 for b in bumps if b.category == "power_sense")
    vss_count = sum(1 for b in bumps if b.category == "vss")
    diff_count = sum(1 for b in bumps if b.category == "signal_diff")
    single_count = sum(1 for b in bumps if b.category == "signal_single")

    return {
        "total_count": len(bumps),
        "power_count": power_count,
        "power_sense_count": power_sense_count,
        "vss_count": vss_count,
        "signal_total_count": diff_count + single_count,
        "diff_pair_count": diff_count // 2,
        "single_end_count": single_count,
        "min_distance": min_dist,
        "bumps": bumps,
        "diff_pairs": diff_pairs,
        "power_by_name": _count_by_name(bumps, "power"),
        "power_sense_by_name": _count_by_name(bumps, "power_sense"),
    }


# ---------------------------------------------------------------------------
# 报告导出
# ---------------------------------------------------------------------------


def export_analysis_report(result: BumpAnalysisResult, output_path: str) -> None:
    """将分析结果导出为 Excel 文件（多 Sheet）。

    Args:
        result: analyze_bumps 返回的分析结果。
        output_path: 输出 Excel 文件路径。

    Raises:
        OSError: 文件写入失败时抛出。
    """
    bumps = result["bumps"]
    diff_pairs = result["diff_pairs"]

    summary_df = pd.DataFrame(
        {
            "分类": [
                "总 bump 数量",
                "Power bump 数量",
                "Power Sense bump 数量",
                "VSS bump 数量",
                "Signal bump 总数",
                "差分信号对数量",
                "单端信号 bump 数量",
                "最小 bump 间距",
            ],
            "值": [
                result["total_count"],
                result["power_count"],
                result["power_sense_count"],
                result["vss_count"],
                result["signal_total_count"],
                result["diff_pair_count"],
                result["single_end_count"],
                round(result["min_distance"], 4),
            ],
        }
    )

    def _to_df(items: List[BumpInfo]) -> pd.DataFrame:
        return pd.DataFrame(
            [{"Bump名": b.name, "X坐标": b.x, "Y坐标": b.y} for b in items]
        )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        # Power bump 按名称统计
        power_name_df = pd.DataFrame(
            [
                {"Power名称": name, "数量": cnt}
                for name, cnt in result["power_by_name"].items()
            ]
        )
        if not power_name_df.empty:
            power_name_df.to_excel(
                writer, sheet_name="Power_By_Name", index=False
            )

        # Power Sense bump 按名称统计
        ps_name_df = pd.DataFrame(
            [
                {"PowerSense名称": name, "数量": cnt}
                for name, cnt in result["power_sense_by_name"].items()
            ]
        )
        if not ps_name_df.empty:
            ps_name_df.to_excel(
                writer, sheet_name="Power_Sense_By_Name", index=False
            )

        _to_df([b for b in bumps if b.category == "power"]).to_excel(
            writer, sheet_name="Power_Bumps", index=False
        )
        _to_df([b for b in bumps if b.category == "power_sense"]).to_excel(
            writer, sheet_name="Power_Sense_Bumps", index=False
        )
        _to_df([b for b in bumps if b.category == "vss"]).to_excel(
            writer, sheet_name="VSS_Bumps", index=False
        )
        _to_df([b for b in bumps if b.category == "signal_diff"]).to_excel(
            writer, sheet_name="Diff_Signal_Bumps", index=False
        )
        _to_df([b for b in bumps if b.category == "signal_single"]).to_excel(
            writer, sheet_name="Single_End_Bumps", index=False
        )

        # 差分对明细
        if diff_pairs:
            pair_df = pd.DataFrame(
                [
                    {
                        "Bump1": b1.name,
                        "X1": b1.x,
                        "Y1": b1.y,
                        "Bump2": b2.name,
                        "X2": b2.x,
                        "Y2": b2.y,
                    }
                    for b1, b2 in diff_pairs
                ]
            )
            pair_df.to_excel(writer, sheet_name="Diff_Pairs_Detail", index=False)


def print_analysis_report(result: BumpAnalysisResult) -> None:
    """在控制台打印分析结果。

    Args:
        result: analyze_bumps 返回的分析结果。
    """
    print("=" * 45)
    print("Bump List 分析 - 结果")
    print("=" * 45)
    print(f"  总 bump 数量:         {result['total_count']}")
    print(f"  Power bump 数量:       {result['power_count']}")
    if result["power_by_name"]:
        print("    -- 明细 --")
        for name, cnt in result["power_by_name"].items():
            print(f"      {name}: {cnt}")
    print(f"  Power Sense bump 数量: {result['power_sense_count']}")
    if result["power_sense_by_name"]:
        print("    -- 明细 --")
        for name, cnt in result["power_sense_by_name"].items():
            print(f"      {name}: {cnt}")
    print(f"  VSS bump 数量:         {result['vss_count']}")
    print(f"  Signal bump 总数:      {result['signal_total_count']}")
    print(f"  差分信号对数量:         {result['diff_pair_count']}")
    print(f"  单端信号 bump 数量:     {result['single_end_count']}")
    print(f"  最小 bump 间距:         {result['min_distance']:.4f}")
    print("=" * 45)
