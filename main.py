"""Bump Version Compare - tkinter GUI 主程序。

通过图形界面选择 Excel 文件，比较两个版本的 bump list，
显示统计结果，并生成 Excel 报告文件。
"""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

from bump_compare import (
    CompareStats,
    Bump,
    compare_bumps,
    print_report,
    read_bump_sheet,
    write_excel_report,
)


class BumpCompareApp:
    """Bump Version Compare 主窗口应用。"""

    def __init__(self, root: tk.Tk) -> None:
        """初始化主窗口。

        Args:
            root: tkinter 根窗口。
        """
        self.root = root
        self.root.title("Bump Version Compare")
        self.root.geometry("800x620")
        self.root.resizable(True, True)
        self.root.minsize(600, 400)

        self.input_path: tk.StringVar = tk.StringVar()
        self.output_path: tk.StringVar = tk.StringVar()
        self.status_text: tk.StringVar = tk.StringVar(value="就绪")

        self._build_layout()

    def _build_layout(self) -> None:
        """构建界面布局。"""
        # 顶部：文件选择区
        top_frame = tk.LabelFrame(self.root, text="文件选择", padx=10, pady=10)
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
        mid_frame = tk.LabelFrame(self.root, text="比较结果", padx=10, pady=10)
        mid_frame.pack(fill="both", expand=True, padx=10, pady=5)

        self.result_text = tk.Text(
            mid_frame, height=18, width=90, state="disabled", wrap="none"
        )
        scrollbar = tk.Scrollbar(mid_frame, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)
        self.result_text.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # 底部：操作区
        bot_frame = tk.Frame(self.root)
        bot_frame.pack(fill="x", padx=10, pady=(5, 10))

        self.compare_btn = tk.Button(
            bot_frame, text="执行比较", width=15, command=self._on_compare
        )
        self.compare_btn.pack(side="left", padx=(0, 5))

        tk.Button(bot_frame, text="退出", width=10, command=self.root.quit).pack(
            side="right"
        )

        # 状态栏
        status_bar = tk.Label(
            self.root,
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
            self.root.after(0, self._on_compare_error, str(e))
            return

        stats: CompareStats = compare_bumps(old_bumps, new_bumps)

        try:
            write_excel_report(stats, output_file)
        except (OSError, Exception) as e:
            self.root.after(0, self._on_compare_error, f"写入报告失败: {e}")
            return

        print_report(stats)
        self.root.after(0, self._on_compare_success, stats, output_file)

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


def main() -> None:
    """程序入口，创建主窗口并启动事件循环。"""
    root = tk.Tk()
    BumpCompareApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
