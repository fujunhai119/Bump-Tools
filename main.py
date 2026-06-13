"""Bump 工具箱 - tkinter GUI 主程序。

整合 Bump 版本比较 和 Bump DXF 生成 两个功能模块。
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from bump_compare import (
    CompareStats,
    Bump,
    compare_bumps,
    print_report,
    read_bump_sheet,
    write_excel_report,
)
from dxf_gui import DxfExportPanel

from bump_analyzer import (
    BumpAnalysisResult,
    BumpInfo,
    analyze_bumps,
    export_analysis_report,
    read_bump_list,
)


class BumpCompareApp:
    """Bump Version Compare 主窗口应用（嵌入型）。"""

    def __init__(self, parent: tk.Widget) -> None:
        """初始化主窗口。

        Args:
            parent: 父容器（tk.Widget），如 tk.Frame。
        """
        self.parent = parent
        self.parent.configure(bg="#F0F2F5")

        self.input_path: tk.StringVar = tk.StringVar()
        self.output_path: tk.StringVar = tk.StringVar()
        self.status_text: tk.StringVar = tk.StringVar(value="就绪")

        self._build_layout()

    def _build_layout(self) -> None:
        """构建界面布局。"""
        # 顶部：文件选择区
        top_frame = tk.LabelFrame(self.parent, text="文件选择", padx=10, pady=10)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        # Excel 文件选择行
        input_row = tk.Frame(top_frame)
        input_row.pack(fill="x", pady=(0, 8))
        tk.Label(input_row, text="Excel 文件：", width=10, anchor="w").pack(side="left")
        tk.Entry(
            input_row, textvariable=self.input_path, width=55, state="readonly"
        ).pack(side="left", padx=(0, 5), fill="x", expand=True)
        tk.Button(
            input_row, text="浏览...", width=8, command=self._on_browse_input
        ).pack(side="left")

        # 输出路径选择行
        output_row = tk.Frame(top_frame)
        output_row.pack(fill="x")
        tk.Label(output_row, text="输出路径：", width=10, anchor="w").pack(side="left")
        tk.Entry(
            output_row, textvariable=self.output_path, width=55, state="readonly"
        ).pack(side="left", padx=(0, 5), fill="x", expand=True)
        tk.Button(
            output_row, text="选择...", width=8, command=self._on_browse_output
        ).pack(side="left")

        # 中部：结果展示区
        mid_frame = tk.LabelFrame(self.parent, text="比较结果", padx=10, pady=10)
        mid_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(
            mid_frame, height=18, width=90, state="disabled", wrap="none"
        )
        scrollbar = tk.Scrollbar(mid_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 底部：操作区
        bot_frame = tk.Frame(self.parent)
        bot_frame.pack(fill="x", padx=10, pady=(5, 10))

        self.compare_btn = tk.Button(
            bot_frame, text="执行比较", width=15, command=self._on_compare
        )
        self.compare_btn.pack(side="left", padx=(0, 5))

        tk.Button(bot_frame, text="退出", width=10, command=self.parent.quit).pack(
            side="right"
        )

        # 状态栏
        status_bar = tk.Label(
            self.parent,
            textvariable=self.status_text,
            bd=1,
            relief="sunken",
            anchor="w",
        )
        status_bar.pack(side="bottom", fill="x")

    def _on_browse_input(self) -> None:
        """打开文件选择对话框，选择输入 Excel 文件。"""
        filepath = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not filepath:
            return

        self.input_path.set(filepath)
        # 自动生成默认输出路径
        input_p = Path(filepath)
        default_output = input_p.parent / f"{input_p.stem}_compare_report.xlsx"
        self.output_path.set(str(default_output))

    def _on_browse_output(self) -> None:
        """打开保存对话框，选择输出 Excel 文件路径。"""
        filepath = filedialog.asksaveasfilename(
            title="保存报告为",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if filepath:
            self.output_path.set(filepath)

    def _on_compare(self) -> None:
        """点击"执行比较"按钮的回调，启动后台线程执行比较。"""
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showwarning("提示", "请先选择 Excel 文件。")
            return
        if not output_file:
            messagebox.showwarning("提示", "请先选择输出路径。")
            return

        self.compare_btn.config(state="disabled")
        self.status_text.set("正在比较...")
        self._set_result_text("正在比较，请稍候...\n")

        thread = threading.Thread(
            target=self._do_compare, args=(input_file, output_file), daemon=True
        )
        thread.start()

    def _do_compare(self, input_file: str, output_file: str) -> None:
        """后台线程执行比较逻辑。

        Args:
            input_file: 输入 Excel 文件路径。
            output_file: 输出报告文件路径。
        """
        try:
            old_bumps: set[Bump] = read_bump_sheet(input_file, sheet_name="Sheet2")
            new_bumps: set[Bump] = read_bump_sheet(input_file, sheet_name="Sheet1")
        except (FileNotFoundError, ValueError, Exception) as e:
            self.parent.after(0, self._on_compare_error, str(e))
            return

        stats: CompareStats = compare_bumps(old_bumps, new_bumps)

        try:
            write_excel_report(stats, output_file)
        except (OSError, Exception) as e:
            self.parent.after(0, self._on_compare_error, f"写入报告失败: {e}")
            return

        print_report(stats)
        self.parent.after(0, self._on_compare_success, stats, output_file)

    def _on_compare_success(self, stats: CompareStats, output_file: str) -> None:
        """比较成功，更新界面并显示结果。

        Args:
            stats: 比较统计结果。
            output_file: 输出报告文件路径。
        """
        self.compare_btn.config(state="normal")
        self.status_text.set(f"完成 | 报告已保存: {output_file}")

        report = (
            f"{'=' * 45}\n"
            f" Bump Version Compare - 比较结果\n"
            f"{'=' * 45}\n"
            f"  输入文件: {self.input_path.get()}\n"
            f"{'=' * 45}\n"
            f"  相同的 bump 数量:   {stats['same_count']}\n"
            f"  不同的 bump 数量:   {stats['different_count']}\n"
            f"  删除的 bump 数量:   {stats['deleted_count']}\n"
            f"  新增的 bump 数量:   {stats['added_count']}\n"
            f"{'=' * 45}\n"
            f"  报告已保存至:\n  {output_file}\n"
            f"{'=' * 45}\n"
        )
        self._set_result_text(report)
        messagebox.showinfo("完成", f"比较完成！\n报告已保存至:\n{output_file}")

    def _on_compare_error(self, error_msg: str) -> None:
        """比较失败，显示错误信息。

        Args:
            error_msg: 错误描述。
        """
        self.compare_btn.config(state="normal")
        self.status_text.set("出错")
        self._set_result_text(f"比较失败:\n{error_msg}\n")
        messagebox.showerror("错误", error_msg)

    def _set_result_text(self, text: str) -> None:
        """设置结果文本框内容。

        Args:
            text: 要显示的文本内容。
        """
        self.result_text.config(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", text)
        self.result_text.config(state="disabled")


class BumpListAnalyzerPanel(tk.Frame):
    """Bump List 分析面板（Tab 3）。"""

    def __init__(self, parent: tk.Widget) -> None:
        """初始化分析面板。

        Args:
            parent: 父容器。
        """
        super().__init__(parent)
        self.parent = parent
        self.parent.configure(bg="#F0F2F5")

        self.file_path: tk.StringVar = tk.StringVar()
        self.status_text: tk.StringVar = tk.StringVar(value="就绪")
        self._analysis_result: BumpAnalysisResult | None = None

        self._build_layout()

    def _build_layout(self) -> None:
        """构建界面布局。"""
        # 顶部：文件选择区
        top_frame = tk.LabelFrame(self.parent, text="文件选择", padx=10, pady=10)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        file_row = tk.Frame(top_frame)
        file_row.pack(fill="x")
        tk.Label(file_row, text="Excel 文件：", width=10, anchor="w").pack(side="left")
        tk.Entry(
            file_row, textvariable=self.file_path, width=55, state="readonly"
        ).pack(side="left", padx=(0, 5), fill="x", expand=True)
        tk.Button(
            file_row, text="浏览...", width=8, command=self._on_browse_file
        ).pack(side="left")

        # 中部：统计结果区
        mid_frame = tk.LabelFrame(self.parent, text="统计分析结果", padx=10, pady=10)
        mid_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # 左右分栏容器
        stats_columns = tk.Frame(mid_frame)
        stats_columns.pack(fill="both", expand=True)

        # 左侧：汇总统计表格
        left_frame = tk.LabelFrame(stats_columns, text="汇总", padx=5, pady=5)
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 3))

        columns = ("分类", "数量")
        self.stats_tree = ttk.Treeview(
            left_frame, columns=columns, show="headings", height=8
        )
        for col in columns:
            self.stats_tree.heading(col, text=col)
            self.stats_tree.column(col, width=180, anchor="center")
        stats_scroll = tk.Scrollbar(left_frame, command=self.stats_tree.yview)
        self.stats_tree.configure(yscrollcommand=stats_scroll.set)
        self.stats_tree.pack(side="left", fill="both", expand=True)
        stats_scroll.pack(side="right", fill="y")

        # 右侧：Power bump 名称明细
        right_frame = tk.LabelFrame(stats_columns, text="Power Bump 明细", padx=5, pady=5)
        right_frame.pack(side="right", fill="both", expand=True, padx=(3, 0))

        power_columns = ("Power 名称", "数量")
        self.power_tree = ttk.Treeview(
            right_frame, columns=power_columns, show="headings", height=8
        )
        for col in power_columns:
            self.power_tree.heading(col, text=col)
            self.power_tree.column(col, width=180, anchor="center")
        power_scroll = tk.Scrollbar(right_frame, command=self.power_tree.yview)
        self.power_tree.configure(yscrollcommand=power_scroll.set)
        self.power_tree.pack(side="left", fill="both", expand=True)
        power_scroll.pack(side="right", fill="y")

        # 最小距离显示
        self.min_dist_var: tk.StringVar = tk.StringVar(value="最小 bump 间距: --")
        dist_label = tk.Label(
            mid_frame, textvariable=self.min_dist_var, font=("Microsoft YaHei", 10)
        )
        dist_label.pack(side="top", anchor="w", padx=10, pady=(5, 0))

        # 差分信号展示区
        diff_frame = tk.LabelFrame(self.parent, text="差分信号对", padx=10, pady=10)
        diff_frame.pack(fill="x", padx=10, pady=5)

        self.diff_text = tk.Text(diff_frame, height=6, width=90, state="disabled", wrap="none")
        diff_scroll = tk.Scrollbar(diff_frame, command=self.diff_text.yview)
        self.diff_text.configure(yscrollcommand=diff_scroll.set)
        self.diff_text.pack(side="left", fill="x", expand=True)
        diff_scroll.pack(side="right", fill="y")

        # 底部：操作区
        bot_frame = tk.Frame(self.parent)
        bot_frame.pack(fill="x", padx=10, pady=(5, 10))

        self.analyze_btn = tk.Button(
            bot_frame, text="开始分析", width=15, command=self._on_analyze
        )
        self.analyze_btn.pack(side="left", padx=(0, 5))

        self.export_btn = tk.Button(
            bot_frame, text="导出报告", width=15, command=self._on_export, state="disabled"
        )
        self.export_btn.pack(side="left")

        tk.Button(bot_frame, text="退出", width=10, command=self.parent.quit).pack(
            side="right"
        )

        # 状态栏
        status_bar = tk.Label(
            self.parent,
            textvariable=self.status_text,
            bd=1,
            relief="sunken",
            anchor="w",
        )
        status_bar.pack(side="bottom", fill="x")

    def _on_browse_file(self) -> None:
        """打开文件选择对话框，选择 Excel 文件。"""
        filepath = filedialog.askopenfilename(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not filepath:
            return
        self.file_path.set(filepath)
        self.status_text.set(f"已选择文件: {Path(filepath).name}")

    def _on_analyze(self) -> None:
        """点击"开始分析"按钮的回调，启动后台线程执行分析。"""
        input_file = self.file_path.get()
        if not input_file:
            messagebox.showwarning("提示", "请先选择 Excel 文件。")
            return

        self.analyze_btn.config(state="disabled")
        self.export_btn.config(state="disabled")
        self.status_text.set("正在分析...")
        self._set_diff_text("正在分析，请稍候...\n")

        thread = threading.Thread(
            target=self._do_analyze, args=(input_file,), daemon=True
        )
        thread.start()

    def _do_analyze(self, input_file: str) -> None:
        """后台线程执行分析逻辑。

        Args:
            input_file: 输入 Excel 文件路径。
        """
        try:
            bumps: list[BumpInfo] = read_bump_list(input_file)
        except (FileNotFoundError, ValueError, Exception) as e:
            self.parent.after(0, self._on_analyze_error, str(e))
            return

        result: BumpAnalysisResult = analyze_bumps(bumps)
        self._analysis_result = result

        self.parent.after(0, self._on_analyze_success, result, input_file)

    def _on_analyze_success(self, result: BumpAnalysisResult, input_file: str) -> None:
        """分析成功，更新界面并显示结果。

        Args:
            result: 分析结果。
            input_file: 输入文件路径。
        """
        self.analyze_btn.config(state="normal")
        self.export_btn.config(state="normal")
        self.status_text.set("分析完成")

        # 更新统计表格
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        stats_data = [
            ("总 bump 数量", result["total_count"]),
            ("Power bump 数量", result["power_count"]),
            ("Power Sense bump 数量", result["power_sense_count"]),
            ("VSS bump 数量", result["vss_count"]),
            ("Signal bump 总数", result["signal_total_count"]),
            ("差分信号对数量", result["diff_pair_count"]),
            ("单端信号 bump 数量", result["single_end_count"]),
        ]
        for category, count in stats_data:
            self.stats_tree.insert("", "end", values=(category, count))

        # 更新 Power bump 名称明细
        for item in self.power_tree.get_children():
            self.power_tree.delete(item)
        for name, cnt in result["power_by_name"].items():
            self.power_tree.insert("", "end", values=(name, cnt))

        # 更新最小距离
        self.min_dist_var.set(f"最小 bump 间距: {result['min_distance']:.4f}")

        # 更新差分信号展示
        diff_pairs = result["diff_pairs"]
        if diff_pairs:
            diff_lines = [f"{b1.name}  <-->  {b2.name}" for b1, b2 in diff_pairs]
            self._set_diff_text("\n".join(diff_lines))
        else:
            self._set_diff_text("未识别到差分信号对。")

        messagebox.showinfo("完成", f"分析完成！\n共分析 {result['total_count']} 个 bump。")

    def _on_analyze_error(self, error_msg: str) -> None:
        """分析失败，显示错误信息。

        Args:
            error_msg: 错误描述。
        """
        self.analyze_btn.config(state="normal")
        self.status_text.set("出错")
        self._set_diff_text(f"分析失败:\n{error_msg}\n")
        messagebox.showerror("错误", error_msg)

    def _on_export(self) -> None:
        """点击"导出报告"按钮的回调。"""
        if self._analysis_result is None:
            messagebox.showwarning("提示", "请先完成分析。")
            return

        output_file = filedialog.asksaveasfilename(
            title="保存分析报告为",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if not output_file:
            return

        try:
            export_analysis_report(self._analysis_result, output_file)
        except (OSError, Exception) as e:
            messagebox.showerror("导出失败", f"写入报告失败: {e}")
            return

        self.status_text.set(f"报告已保存: {Path(output_file).name}")
        messagebox.showinfo("完成", f"报告已保存至:\n{output_file}")

    def _set_diff_text(self, text: str) -> None:
        """设置差分信号文本框内容。

        Args:
            text: 要显示的文本内容。
        """
        self.diff_text.config(state="normal")
        self.diff_text.delete("1.0", "end")
        self.diff_text.insert("1.0", text)
        self.diff_text.config(state="disabled")


def main() -> None:
    """程序入口，创建主窗口并启动事件循环。"""
    root = tk.Tk()
    root.title("Bump 工具箱")
    root.geometry("1050x750")
    root.resizable(True, True)
    root.minsize(900, 600)

    # 高 DPI 支持（Windows）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except (ImportError, AttributeError):
        pass

    notebook = ttk.Notebook(root)
    notebook.pack(fill="both", expand=True)

    # Tab 1: Bump 版本比较
    tab_compare = tk.Frame(notebook)
    notebook.add(tab_compare, text="Bump 版本比较")
    BumpCompareApp(tab_compare)

    # Tab 2: Bump DXF 生成
    tab_dxf = tk.Frame(notebook)
    notebook.add(tab_dxf, text="Bump DXF 生成")
    DxfExportPanel(tab_dxf).pack(fill="both", expand=True)

    # Tab 3: Bump List 分析
    tab_analyzer = tk.Frame(notebook)
    notebook.add(tab_analyzer, text="Bump List 分析")
    BumpListAnalyzerPanel(tab_analyzer)

    root.mainloop()


if __name__ == "__main__":
    main()
