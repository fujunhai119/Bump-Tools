"""Bump Version Compare - 核心比较逻辑模块。

提供 Bump 数据类、Excel 读取、bump 比较、Excel 报告生成功能。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Set, TypedDict

import pandas as pd


class CompareStats(TypedDict):
    """比较结果统计字典类型。"""

    same_count: int
    different_count: int
    deleted_count: int
    added_count: int
    same_bumps: Set[Bump]
    deleted_bumps: Set[Bump]
    added_bumps: Set[Bump]


@dataclass(frozen=True)
class Bump:
    """表示单个 bump，包含名称和坐标。

    Attributes:
        name: bump 名称。
        x: x 坐标（完全匹配，不允许浮点误差）。
        y: y 坐标（完全匹配，不允许浮点误差）。
    """

    name: str
    x: float
    y: float


def read_bump_sheet(filepath: str, sheet_name: str) -> Set[Bump]:
    """读取 Excel 指定 Sheet 中的 bump 列表。

    Args:
        filepath: Excel 文件路径。
        sheet_name: Sheet 名称（如 'Sheet1'）。

    Returns:
        包含该 Sheet 所有 bump 的集合。

    Raises:
        FileNotFoundError: 文件不存在时抛出。
        ValueError: Sheet 不存在或缺少必要列（A/B/C 列）时抛出。
        ValueError: 坐标列（B/C 列）包含非数值时抛出。
    """
    try:
        df = pd.read_excel(
            filepath,
            sheet_name=sheet_name,
            usecols=[0, 1, 2],
            header=None,
        )
    except FileNotFoundError:
        raise FileNotFoundError(f"文件不存在: {filepath}")
    except ValueError as e:
        if "Worksheet" in str(e) or "sheet" in str(e).lower():
            raise ValueError(f"Sheet '{sheet_name}' 不存在，请检查 Excel 文件。")
        raise

    if df.empty:
        raise ValueError(f"Sheet '{sheet_name}' 为空，请检查文件内容。")

    bumps: Set[Bump] = set()
    for idx, row in df.iterrows():
        name = str(row[0]).strip()
        try:
            x = float(row[1])
            y = float(row[2])
        except (ValueError, TypeError):
            raise ValueError(
                f"Sheet '{sheet_name}' 第 {idx + 1} 行坐标非数值："
                f"x={row[1]}, y={row[2]}"
            )
        bumps.add(Bump(name=name, x=x, y=y))

    return bumps


def compare_bumps(
    old_bumps: Set[Bump],
    new_bumps: Set[Bump],
) -> CompareStats:
    """比较两个 bump 集合，返回统计结果。

    Args:
        old_bumps: 旧版本 bump 集合。
        new_bumps: 新版本 bump 集合。

    Returns:
        包含统计信息的 CompareStats 字典。
    """
    same_bumps: Set[Bump] = old_bumps & new_bumps
    deleted_bumps: Set[Bump] = old_bumps - new_bumps
    added_bumps: Set[Bump] = new_bumps - old_bumps
    different_bumps: Set[Bump] = deleted_bumps | added_bumps

    return {
        "same_count": len(same_bumps),
        "different_count": len(different_bumps),
        "deleted_count": len(deleted_bumps),
        "added_count": len(added_bumps),
        "same_bumps": same_bumps,
        "deleted_bumps": deleted_bumps,
        "added_bumps": added_bumps,
    }


def _bumps_to_dataframe(bumps: Set[Bump]) -> pd.DataFrame:
    """将 bump 集合转换为 DataFrame。"""
    data: List[Dict[str, str | float]] = [
        {"Bump名": b.name, "X坐标": b.x, "Y坐标": b.y} for b in sorted(bumps, key=lambda b: b.name)
    ]
    return pd.DataFrame(data)


def write_excel_report(
    stats: CompareStats,
    output_path: str,
) -> None:
    """将比较结果写入 Excel 报告文件。

    Args:
        stats: compare_bumps 返回的统计结果。
        output_path: 输出 Excel 文件路径。

    Raises:
        OSError: 文件写入失败时抛出。
    """
    summary_df = pd.DataFrame(
        {
            "统计项": [
                "相同的 bump 数量",
                "不同的 bump 数量",
                "删除的 bump 数量（旧版有，新版无）",
                "新增的 bump 数量（新版有，旧版无）",
            ],
            "数量": [
                stats["same_count"],
                stats["different_count"],
                stats["deleted_count"],
                stats["added_count"],
            ],
        }
    )

    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        _bumps_to_dataframe(stats["same_bumps"]).to_excel(
            writer, sheet_name="Same_Bumps", index=False
        )
        _bumps_to_dataframe(stats["deleted_bumps"]).to_excel(
            writer, sheet_name="Deleted_Bumps", index=False
        )
        _bumps_to_dataframe(stats["added_bumps"]).to_excel(
            writer, sheet_name="Added_Bumps", index=False
        )


def print_report(stats: CompareStats) -> None:
    """在控制台打印比较统计结果。

    Args:
        stats: compare_bumps 返回的统计结果。
    """
    print("=" * 40)
    print("Bump Version Compare - 比较结果")
    print("=" * 40)
    print(f"相同的 bump 数量:   {stats['same_count']}")
    print(f"不同的 bump 数量:   {stats['different_count']}")
    print(f"删除的 bump 数量:   {stats['deleted_count']}")
    print(f"新增的 bump 数量:   {stats['added_count']}")
    print("=" * 40)
