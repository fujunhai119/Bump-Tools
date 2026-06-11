"""Bump DXF 导出 GUI 面板。

提供 DxfExportPanel(tk.Frame) 类，可嵌入 tkinter Notebook Tab 页。
"""

from __future__ import annotations

import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from bump_dxf import (
    BumpCategory,
    BumpDxf,
    BumpShape,
    DxfExportConfig,
    CATEGORY_LABELS,
    DEFAULT_COLOR_MAP,
    ACI_COLORS,
    get_aci_hex,
    get_aci_display,
    classify_bump,
    parse_bump_excel,
    create_bump_dxf,
)


# ==================== 常量定义 ====================

_UI_BG = "#F0F2F5"
_UI_CARD_BG = "#FFFFFF"
_UI_ACCENT = "#3B82F6"
_UI_SUCCESS = "#10B981"
_UI_WARNING = "#F59E0B"
_UI_ERROR = "#DC2626"
_UI_TEXT = "#1E293B"
_UI_TEXT_SECONDARY = "#475569"
_UI_TEXT_MUTED = "#94A3B8"
_UI_BORDER = "#E2E8F0"

# 颜色文本映射
_COLOR_OPTIONS: list[str] = [
    f"ACI:{aci} {name}" for aci, name, _ in ACI_COLORS
]
_COLOR_TEXT_TO_ACI: dict[str, int] = {
    f"ACI:{aci} {name}": aci for aci, name, _ in ACI_COLORS
}


# ==================== GUI 面板类 ====================


class DxfExportPanel(tk.Frame):
    """Bump DXF 导出面板（嵌入型 Frame）。

    用于嵌入 tkinter Notebook Tab 页，不创建独立窗口。
    """

    def __init__(self, parent: tk.Widget) -> None:
        """初始化面板。

        Args:
            parent: 父容器（Notebook Tab Frame）。
        """
        super().__init__(parent, bg=_UI_BG)
        self.pack(fill=tk.BOTH, expand=True)

        # 数据状态
        self.bumps: list[BumpDxf] = []
        self.excel_file_path: str | None = None
        self.output_dir: str | None = None

        # 配置
        self.config = DxfExportConfig()

        # 颜色预览画布引用
        self._color_preview_canvases: dict[BumpCategory, tk.Canvas] = {}

        self._setup_ui()
        self._update_color_previews()
        self._update_state()

    # ---- UI 布局 ----

    def _setup_ui(self) -> None:
        """初始化 UI 布局（标题栏 + 左右分栏）。"""
        self._create_title_bar()
        main_container = tk.Frame(self, bg=_UI_BG)
        main_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=(0, 15))
        self._create_left_panel(main_container)
        self._create_right_panel(main_container)

    def _create_title_bar(self) -> None:
        """创建顶部标题栏。"""
        frame = tk.Frame(self, bg=_UI_TEXT, height=55)
        frame.pack(fill=tk.X)
        frame.pack_propagate(False)
        tk.Label(
            frame,
            text="Bump DXF 生成",
            font=("Microsoft YaHei", 16, "bold"),
            bg=_UI_TEXT,
            fg="#F8FAFC",
        ).pack(side=tk.LEFT, padx=20, pady=12)
        tk.Label(
            frame,
            text="v1.0",
            font=("Microsoft YaHei", 10),
            bg=_UI_TEXT,
            fg=_UI_TEXT_MUTED,
        ).pack(side=tk.RIGHT, padx=20, pady=12)

    def _create_left_panel(self, parent: tk.Frame) -> None:
        """创建左侧配置面板（带滚动条）。"""
        panel = tk.Frame(parent, bg=_UI_CARD_BG, width=380)
        panel.pack(side=tk.LEFT, fill=tk.BOTH, padx=(0, 8))
        panel.pack_propagate(False)
        self._build_scrollable_config(panel)

    def _create_right_panel(self, parent: tk.Frame) -> None:
        """创建右侧数据预览面板。"""
        panel = tk.Frame(parent, bg=_UI_CARD_BG)
        panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._create_preview_panel(panel)

    # ---- 滚动容器 ----

    def _build_scrollable_config(self, parent: tk.Frame) -> None:
        """在面板中构建可滚动的配置区域。

        Args:
            parent: 父容器 Frame。
        """
        canvas = tk.Canvas(parent, bg=_UI_CARD_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg=_UI_CARD_BG)
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window(
            (0, 0), window=scroll_frame, anchor="nw", width=360
        )
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.bind_all(
            "<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )
        self._populate_config_sections(scroll_frame)

    def _populate_config_sections(self, parent: tk.Frame) -> None:
        """填充配置区域的所有子面板。

        Args:
            parent: 滚动区域的 Frame。
        """
        self._create_file_section(parent)
        self._create_shape_section(parent)
        self._create_color_section(parent)
        self._create_stats_section(parent)
        self._create_export_section(parent)
        # 底部间距
        tk.Frame(parent, height=20, bg=_UI_CARD_BG).pack(fill=tk.X)

    # ---- 文件导入区域 ----

    def _create_file_section(self, parent: tk.Frame) -> None:
        """创建文件导入区域。

        Args:
            parent: 父容器。
        """
        section = tk.LabelFrame(
            parent,
            text=" 文件导入",
            font=("Microsoft YaHei", 11, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
            padx=10,
            pady=10,
        )
        section.pack(fill=tk.X, padx=10, pady=(10, 5))

        self.file_btn = tk.Button(
            section,
            text="选择Excel文件",
            font=("Microsoft YaHei", 11),
            bg=_UI_ACCENT,
            fg="white",
            activebackground="#2563EB",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            padx=12,
            pady=6,
            command=self._select_file,
        )
        self.file_btn.pack(fill=tk.X, pady=(0, 8))

        self.file_label = tk.Label(
            section,
            text="尚未选择文件",
            font=("Microsoft YaHei", 9),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_MUTED,
            wraplength=320,
            justify=tk.LEFT,
        )
        self.file_label.pack(fill=tk.X)

    # ---- 形状配置区域 ----

    def _create_shape_section(self, parent: tk.Frame) -> None:
        """创建形状配置区域。

        Args:
            parent: 父容器。
        """
        section = tk.LabelFrame(
            parent,
            text=" 形状配置",
            font=("Microsoft YaHei", 11, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
            padx=10,
            pady=10,
        )
        section.pack(fill=tk.X, padx=10, pady=5)
        self._build_shape_radio(section)
        self._build_size_input(section)
        self._build_preset_buttons(section)

    def _build_shape_radio(self, parent: tk.Frame) -> None:
        """构建形状单选按钮。

        Args:
            parent: 父容器。
        """
        tk.Label(
            parent,
            text="Bump形状：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(anchor="w", pady=(0, 5))

        radio_frame = tk.Frame(parent, bg=_UI_CARD_BG)
        radio_frame.pack(fill=tk.X, pady=(0, 8))

        self.shape_var = tk.StringVar(value="square")
        tk.Radiobutton(
            radio_frame,
            text="■ 正方形",
            variable=self.shape_var,
            value="square",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            activebackground=_UI_BG,
            command=self._on_shape_changed,
        ).pack(side=tk.LEFT, padx=(0, 20))
        tk.Radiobutton(
            radio_frame,
            text="● 圆形",
            variable=self.shape_var,
            value="circle",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            activebackground=_UI_BG,
            command=self._on_shape_changed,
        ).pack(side=tk.LEFT)

    def _build_size_input(self, parent: tk.Frame) -> None:
        """构建尺寸输入控件。

        Args:
            parent: 父容器。
        """
        tk.Label(
            parent,
            text="边长（正方形）/ 直径（圆形）：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(anchor="w", pady=(5, 5))

        input_frame = tk.Frame(parent, bg=_UI_CARD_BG)
        input_frame.pack(fill=tk.X)
        self.size_entry = tk.Entry(
            input_frame,
            font=("Microsoft YaHei", 11),
            bg="#F8FAFC",
            fg=_UI_TEXT,
            relief=tk.SOLID,
            bd=1,
            width=12,
        )
        self.size_entry.insert(0, "100")
        self.size_entry.pack(side=tk.LEFT, padx=(0, 5))
        tk.Label(
            input_frame,
            text="μm",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg="#64748B",
        ).pack(side=tk.LEFT)

    def _build_preset_buttons(self, parent: tk.Frame) -> None:
        """构建快捷尺寸按钮。

        Args:
            parent: 父容器。
        """
        tk.Label(
            parent,
            text="快捷尺寸：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(anchor="w", pady=(10, 5))
        preset_frame = tk.Frame(parent, bg=_UI_CARD_BG)
        preset_frame.pack(fill=tk.X)
        for val in [50, 80, 100, 130, 150, 200]:
            tk.Button(
                preset_frame,
                text=str(val),
                font=("Microsoft YaHei", 9),
                bg=_UI_BORDER,
                fg="#334155",
                activebackground="#CBD5E1",
                relief=tk.FLAT,
                cursor="hand2",
                padx=6,
                pady=2,
                command=lambda v=val: self._set_size(v),
            ).pack(side=tk.LEFT, padx=2)

    # ---- 颜色配置区域 ----

    def _create_color_section(self, parent: tk.Frame) -> None:
        """创建颜色选择区域。

        Args:
            parent: 父容器。
        """
        section = tk.LabelFrame(
            parent,
            text=" 颜色选择",
            font=("Microsoft YaHei", 11, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
            padx=10,
            pady=10,
        )
        section.pack(fill=tk.X, padx=10, pady=5)

        self.color_vars: dict[BumpCategory, tk.StringVar] = {}
        self._build_color_rows(section)
        self._build_reset_button(section)

    def _build_color_rows(self, parent: tk.Frame) -> None:
        """构建各分类颜色选择行。

        Args:
            parent: 父容器。
        """
        entries = [
            (BumpCategory.GND, "GND 颜色："),
            (BumpCategory.POWER, "Power 颜色："),
            (BumpCategory.SIGNAL, "Signal 颜色："),
            (BumpCategory.CUSTOM, "自定义 颜色："),
        ]
        for cat, label_text in entries:
            row = tk.Frame(parent, bg=_UI_CARD_BG)
            row.pack(fill=tk.X, pady=2)
            tk.Label(
                row,
                text=label_text,
                font=("Microsoft YaHei", 10),
                bg=_UI_CARD_BG,
                fg=_UI_TEXT_SECONDARY,
                width=12,
                anchor="w",
            ).pack(side=tk.LEFT)

            default_aci = self.config.color_map[cat]
            self.color_vars[cat] = tk.StringVar(
                value=get_aci_display(default_aci)
            )
            combo = ttk.Combobox(
                row,
                textvariable=self.color_vars[cat],
                values=_COLOR_OPTIONS,
                state="readonly",
                width=22,
            )
            combo.pack(side=tk.LEFT, padx=(0, 5))
            combo.bind(
                "<<ComboboxSelected>>",
                lambda e, c=cat: self._on_color_changed(c),
            )
            preview = tk.Canvas(
                row,
                width=20,
                height=20,
                bg=_UI_CARD_BG,
                highlightthickness=1,
                highlightbackground="#CBD5E1",
            )
            preview.pack(side=tk.LEFT)
            self._color_preview_canvases[cat] = preview

    def _build_reset_button(self, parent: tk.Frame) -> None:
        """构建恢复默认颜色按钮。

        Args:
            parent: 父容器。
        """
        tk.Button(
            parent,
            text="恢复默认颜色",
            font=("Microsoft YaHei", 9),
            bg=_UI_BORDER,
            fg="#334155",
            activebackground="#CBD5E1",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=3,
            command=self._reset_colors,
        ).pack(anchor="e", pady=(8, 0))

    # ---- 分类统计区域 ----

    def _create_stats_section(self, parent: tk.Frame) -> None:
        """创建分类统计区域。

        Args:
            parent: 父容器。
        """
        section = tk.LabelFrame(
            parent,
            text=" 分类统计",
            font=("Microsoft YaHei", 11, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
            padx=10,
            pady=10,
        )
        section.pack(fill=tk.X, padx=10, pady=5)
        self._build_stats_tree(section)

    def _build_stats_tree(self, parent: tk.Frame) -> None:
        """构建统计表格和总计标签。

        Args:
            parent: 父容器。
        """
        columns = ("category", "count", "color_display")
        self.stats_tree = ttk.Treeview(
            parent, columns=columns, show="headings", height=5
        )
        self.stats_tree.heading("category", text="分类")
        self.stats_tree.heading("count", text="数量")
        self.stats_tree.heading("color_display", text="显示颜色")
        self.stats_tree.column("category", width=80, anchor="center")
        self.stats_tree.column("count", width=70, anchor="center")
        self.stats_tree.column("color_display", width=100, anchor="center")
        self.stats_tree.pack(fill=tk.X)

        self.total_label = tk.Label(
            parent,
            text="总计：0 个Bump",
            font=("Microsoft YaHei", 10, "bold"),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT,
        )
        self.total_label.pack(anchor="e", pady=(8, 0))

    # ---- 导出操作区域 ----

    def _create_export_section(self, parent: tk.Frame) -> None:
        """创建导出操作区域。

        Args:
            parent: 父容器。
        """
        section = tk.LabelFrame(
            parent,
            text=" 导出操作",
            font=("Microsoft YaHei", 11, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
            padx=10,
            pady=10,
        )
        section.pack(fill=tk.X, padx=10, pady=5)
        self._build_output_path_row(section)
        self._build_export_button(section)

    def _build_output_path_row(self, parent: tk.Frame) -> None:
        """构建输出路径选择行。

        Args:
            parent: 父容器。
        """
        tk.Label(
            parent,
            text="输出目录：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(anchor="w")

        row = tk.Frame(parent, bg=_UI_CARD_BG)
        row.pack(fill=tk.X, pady=(2, 8))

        self.output_label = tk.Label(
            row,
            text="（与输入文件同目录）",
            font=("Microsoft YaHei", 9),
            bg="#F8FAFC",
            fg=_UI_TEXT_MUTED,
            anchor="w",
        )
        self.output_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.output_btn = tk.Button(
            row,
            text="选择目录",
            font=("Microsoft YaHei", 9),
            bg=_UI_BORDER,
            fg="#334155",
            activebackground="#CBD5E1",
            relief=tk.FLAT,
            cursor="hand2",
            padx=8,
            pady=2,
            command=self._select_output_dir,
        )
        self.output_btn.pack(side=tk.RIGHT)

    def _build_export_button(self, parent: tk.Frame) -> None:
        """构建导出按钮和状态标签。

        Args:
            parent: 父容器。
        """
        self.export_btn = tk.Button(
            parent,
            text="导出DXF文件",
            font=("Microsoft YaHei", 12, "bold"),
            bg=_UI_SUCCESS,
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            relief=tk.FLAT,
            cursor="hand2",
            pady=8,
            state=tk.DISABLED,
            command=self._export_dxf,
        )
        self.export_btn.pack(fill=tk.X, pady=(0, 5))

        self.status_label = tk.Label(
            parent,
            text="请先导入Excel文件",
            font=("Microsoft YaHei", 9),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_MUTED,
        )
        self.status_label.pack(anchor="w")

    # ---- 预览面板 ----

    def _create_preview_panel(self, parent: tk.Frame) -> None:
        """创建右侧数据预览面板。

        Args:
            parent: 父容器。
        """
        tk.Label(
            parent,
            text="Bump数据预览",
            font=("Microsoft YaHei", 12, "bold"),
            bg=_UI_CARD_BG,
            fg="#334155",
        ).pack(anchor="w", padx=15, pady=(15, 5))
        self._build_search_bar(parent)
        self._build_data_tree(parent)

    def _build_search_bar(self, parent: tk.Frame) -> None:
        """构建搜索和筛选栏。

        Args:
            parent: 父容器。
        """
        frame = tk.Frame(parent, bg=_UI_CARD_BG)
        frame.pack(fill=tk.X, padx=15, pady=(0, 8))

        tk.Label(
            frame,
            text="搜索：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(side=tk.LEFT)

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self._filter_preview())
        tk.Entry(
            frame,
            textvariable=self.search_var,
            font=("Microsoft YaHei", 10),
            bg="#F8FAFC",
            fg=_UI_TEXT,
            relief=tk.SOLID,
            bd=1,
            width=25,
        ).pack(side=tk.LEFT, padx=5)

        tk.Label(
            frame,
            text="分类：",
            font=("Microsoft YaHei", 10),
            bg=_UI_CARD_BG,
            fg=_UI_TEXT_SECONDARY,
        ).pack(side=tk.LEFT, padx=(15, 0))

        self.filter_var = tk.StringVar(value="全部")
        filter_combo = ttk.Combobox(
            frame,
            textvariable=self.filter_var,
            values=["全部", "GND", "Power", "Signal", "自定义"],
            state="readonly",
            width=8,
        )
        filter_combo.pack(side=tk.LEFT, padx=5)
        filter_combo.bind(
            "<<ComboboxSelected>>", lambda e: self._filter_preview()
        )

    def _build_data_tree(self, parent: tk.Frame) -> None:
        """构建数据表格（含滚动条）。

        Args:
            parent: 父容器。
        """
        tree_frame = tk.Frame(parent, bg=_UI_CARD_BG)
        tree_frame.pack(
            fill=tk.BOTH, expand=True, padx=15, pady=(0, 15)
        )

        columns = ("index", "name", "x", "y", "category")
        self.data_tree = ttk.Treeview(
            tree_frame, columns=columns, show="headings", height=20,
            selectmode="extended",
        )
        self.data_tree.heading("index", text="#")
        self.data_tree.heading("name", text="Bump名称")
        self.data_tree.heading("x", text="X坐标 (μm)")
        self.data_tree.heading("y", text="Y坐标 (μm)")
        self.data_tree.heading("category", text="分类")
        self.data_tree.column("index", width=50, anchor="center")
        self.data_tree.column("name", width=180, anchor="w")
        self.data_tree.column("x", width=120, anchor="center")
        self.data_tree.column("y", width=120, anchor="center")
        self.data_tree.column("category", width=100, anchor="center")

        v_scroll = ttk.Scrollbar(
            tree_frame, orient=tk.VERTICAL, command=self.data_tree.yview
        )
        h_scroll = ttk.Scrollbar(
            tree_frame, orient=tk.HORIZONTAL, command=self.data_tree.xview
        )
        self.data_tree.configure(
            yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set
        )
        self.data_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # 分类颜色标记
        self.data_tree.tag_configure("gnd", background="#E8E8E8")
        self.data_tree.tag_configure("power", background="#FEE2E2")
        self.data_tree.tag_configure("signal", background="#DCFCE7")
        self.data_tree.tag_configure("custom", background="#F3E8FF")

        # 右键菜单绑定
        self.data_tree.bind("<Button-3>", self._on_tree_right_click)
        if os.name == "nt":
            self.data_tree.bind("<Button-2>", self._on_tree_right_click)

    # ---- 事件回调 ----

    def _select_file(self) -> None:
        """选择并解析 Excel 文件。"""
        file_path = filedialog.askopenfilename(
            title="选择Bump Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if not file_path:
            return None
        self._load_excel_file(file_path)

    def _load_excel_file(self, file_path: str) -> None:
        """加载并解析 Excel 文件到内存。

        Args:
            file_path: Excel 文件路径。
        """
        self.bumps = parse_bump_excel(file_path)
        self.excel_file_path = file_path

        self.file_label.config(
            text=(
                f"  {os.path.basename(file_path)}\n"
                f"  共 {len(self.bumps)} 个Bump"
            ),
            fg="#16A34A",
        )

        if not self.output_dir:
            self.output_dir = os.path.dirname(file_path)
            self.output_label.config(
                text=self.output_dir, fg=_UI_TEXT_SECONDARY
            )

        self._update_preview()
        self._update_stats()
        self._update_state()

        self.status_label.config(
            text=f"已加载 {len(self.bumps)} 个Bump，准备就绪",
            fg="#16A34A",
        )

    def _select_output_dir(self) -> None:
        """选择输出目录。"""
        dir_path = filedialog.askdirectory(
            title="选择DXF输出目录",
            initialdir=self.output_dir or os.path.expanduser("~"),
        )
        if dir_path:
            self.output_dir = dir_path
            self.output_label.config(text=dir_path, fg=_UI_TEXT_SECONDARY)

    def _on_shape_changed(self) -> None:
        """形状变更回调。"""
        shape_val = self.shape_var.get()
        self.config.shape = (
            BumpShape.SQUARE if shape_val == "square" else BumpShape.CIRCLE
        )

    def _set_size(self, val: int) -> None:
        """设置尺寸输入框的值。

        Args:
            val: 预设尺寸值。
        """
        self.size_entry.delete(0, tk.END)
        self.size_entry.insert(0, str(val))

    def _on_color_changed(self, category: BumpCategory) -> None:
        """颜色选择变更回调。

        Args:
            category: 变更的分类。
        """
        text = self.color_vars[category].get()
        aci = _COLOR_TEXT_TO_ACI.get(text, 7)
        self.config.color_map[category] = aci
        self._update_single_color_preview(category)
        self._update_stats()

    def _reset_colors(self) -> None:
        """恢复所有分类为默认颜色。"""
        for cat in (
            BumpCategory.GND,
            BumpCategory.POWER,
            BumpCategory.SIGNAL,
            BumpCategory.CUSTOM,
        ):
            aci = DEFAULT_COLOR_MAP[cat]
            self.config.color_map[cat] = aci
            self.color_vars[cat].set(get_aci_display(aci))
            self._update_single_color_preview(cat)
        self._update_stats()

    def _update_single_color_preview(self, category: BumpCategory) -> None:
        """更新单个分类的颜色预览色块。

        Args:
            category: Bump 分类。
        """
        if category not in self._color_preview_canvases:
            return
        aci = self.config.color_map[category]
        self._color_preview_canvases[category].configure(
            bg=get_aci_hex(aci)
        )

    def _update_color_previews(self) -> None:
        """初始化时刷新所有颜色预览色块。"""
        for cat in (
            BumpCategory.GND,
            BumpCategory.POWER,
            BumpCategory.SIGNAL,
            BumpCategory.CUSTOM,
        ):
            self._update_single_color_preview(cat)

    # ---- 配置获取 ----

    def _get_config(self) -> DxfExportConfig | None:
        """从 UI 控件获取当前导出配置。

        Returns:
            DxfExportConfig 对象；尺寸无效时返回 None。
        """
        shape = (
            BumpShape.SQUARE
            if self.shape_var.get() == "square"
            else BumpShape.CIRCLE
        )
        try:
            size = float(self.size_entry.get())
            if size <= 0:
                messagebox.showwarning("警告", "尺寸必须大于0！")
                return None
        except ValueError:
            messagebox.showwarning("警告", "请输入有效的尺寸数值！")
            return None
        return DxfExportConfig(
            shape=shape, size=size, color_map=dict(self.config.color_map)
        )

    # ---- DXF 导出 ----

    def _export_dxf(self) -> None:
        """导出 DXF 文件（含路径选择与结果反馈）。"""
        if not self.bumps:
            messagebox.showwarning("警告", "没有Bump数据可以导出！")
            return

        config = self._get_config()
        if config is None:
            return

        file_path = self._prompt_save_path(config)
        if not file_path:
            return

        self._do_export(config, file_path)

    def _prompt_save_path(self, config: DxfExportConfig) -> str | None:
        """弹出保存文件对话框，获取输出路径。

        Args:
            config:

        Returns:
            选中的文件路径，取消返回 None。
        """
        if not self.output_dir:
            self.output_dir = os.path.dirname(self.excel_file_path or "")
        if not self.output_dir:
            self.output_dir = os.path.expanduser("~")

        base = (
            os.path.splitext(os.path.basename(self.excel_file_path or ""))[0]
            or "bump"
        )
        suffix = "sq" if config.shape == BumpShape.SQUARE else "cir"
        default_name = f"{base}_dxf_{suffix}{int(config.size)}.dxf"

        return filedialog.asksaveasfilename(
            title="保存DXF文件",
            initialdir=self.output_dir,
            initialfile=default_name,
            defaultextension=".dxf",
            filetypes=[("DXF文件", "*.dxf"), ("所有文件", "*.*")],
        )

    def _do_export(self, config: DxfExportConfig, file_path: str) -> None:
        """执行 DXF 导出并显示结果。

        Args:
            config: 导出配置。
            file_path: 输出文件路径。
        """
        try:
            self.status_label.config(
                text="正在生成DXF文件...", fg=_UI_WARNING
            )
            self.update()

            output_path = create_bump_dxf(self.bumps, config, file_path)

            self.status_label.config(
                text=f"  DXF文件已导出：{os.path.basename(output_path)}",
                fg="#16A34A",
            )
            messagebox.showinfo(
                "导出成功", self._build_success_message(config, output_path)
            )
        except Exception as exc:
            self.status_label.config(text="导出失败", fg=_UI_ERROR)
            messagebox.showerror("导出失败", f"生成DXF文件时出错：{exc}")

    def _build_success_message(
        self, config: DxfExportConfig, output_path: str
    ) -> str:
        """构建导出成功的提示消息。

        Args:
            config: 导出配置。
            output_path: 输出文件路径。

        Returns:
            格式化的成功消息字符串。
        """
        shape_desc = (
            f"正方形 边长{config.size}μm"
            if config.shape == BumpShape.SQUARE
            else f"圆形 直径{config.size}μm"
        )

        def _desc(cat: BumpCategory) -> str:
            aci = config.color_map.get(cat, DEFAULT_COLOR_MAP[cat])
            return get_aci_display(aci)

        return (
            f"DXF文件已成功导出！\n\n"
            f"文件路径：{output_path}\n"
            f"Bump数量：{len(self.bumps)}\n"
            f"形状：{shape_desc}\n\n"
            f"图层颜色：\n"
            f"  GND层 — {_desc(BumpCategory.GND)}\n"
            f"  POWER层 — {_desc(BumpCategory.POWER)}\n"
            f"  SIGNAL层 — {_desc(BumpCategory.SIGNAL)}\n"
            f"  自定义层 — {_desc(BumpCategory.CUSTOM)}\n\n"
            f"可用AutoCAD、DraftSight等软件打开查看。"
        )

    # ---- 预览数据刷新 ----

    def _update_preview(self) -> None:
        """更新数据预览表格（显示全部数据）。"""
        self._clear_tree(self.data_tree)
        if not self.bumps:
            return
        tag_map = {
            BumpCategory.GND: "gnd",
            BumpCategory.POWER: "power",
            BumpCategory.SIGNAL: "signal",
            BumpCategory.CUSTOM: "custom",
        }
        for i, bump in enumerate(self.bumps, start=1):
            self.data_tree.insert(
                "",
                tk.END,
                values=(
                    i,
                    bump.name,
                    f"{bump.x:.3f}",
                    f"{bump.y:.3f}",
                    CATEGORY_LABELS[bump.category],
                ),
                tags=(tag_map.get(bump.category, ""),),
            )

    def _filter_preview(self) -> None:
        """根据搜索文本和分类筛选刷新预览表格。"""
        search_text = self.search_var.get().strip().lower()
        filter_cat = self.filter_var.get()
        self._clear_tree(self.data_tree)
        if not self.bumps:
            return

        tag_map = {
            BumpCategory.GND: "gnd",
            BumpCategory.POWER: "power",
            BumpCategory.SIGNAL: "signal",
            BumpCategory.CUSTOM: "custom",
        }
        display_index = 0
        for bump in self.bumps:
            if not self._matches_filter(bump.category, filter_cat):
                continue
            if search_text and search_text not in bump.name.lower():
                continue
            display_index += 1
            self.data_tree.insert(
                "",
                tk.END,
                values=(
                    display_index,
                    bump.name,
                    f"{bump.x:.3f}",
                    f"{bump.y:.3f}",
                    CATEGORY_LABELS[bump.category],
                ),
                tags=(tag_map.get(bump.category, ""),),
            )

    @staticmethod
    def _matches_filter(category: BumpCategory, filter_val: str) -> bool:
        """检查分类是否匹配筛选条件。

        Args:
            category: Bump 分类。
            filter_val: 筛选值（"全部"/"GND"/"Power"/"Signal"）。

        Returns:
            True 表示匹配。
        """
        match filter_val:
            case "全部":
                return True
            case "GND":
                return category == BumpCategory.GND
            case "Power":
                return category == BumpCategory.POWER
            case "Signal":
                return category == BumpCategory.SIGNAL
            case "自定义":
                return category == BumpCategory.CUSTOM
            case _:
                return True

    @staticmethod
    def _clear_tree(tree: ttk.Treeview) -> None:
        """清空 Treeview 中所有数据行。

        Args:
            tree: 目标 Treeview 控件。
        """
        for item in tree.get_children():
            tree.delete(item)

    # ---- 右键菜单（自定义分类） ----

    def _on_tree_right_click(self, event: tk.Event) -> None:
        """数据表格右键菜单：设置选中 bump 为自定义分类或恢复。

        Args:
            event: Tkinter 鼠标事件。
        """
        selection = self.data_tree.selection()
        if not selection:
            return

        # 移除已有选择高亮，选中当前右键目标行
        self.data_tree.selection_set(
            self.data_tree.identify_row(event.y)
        )
        # 确保原有多选也保留
        if self.data_tree.identify_row(event.y) not in selection:
            selection = self.data_tree.selection()

        # 检查选中项中是否有自定义分类
        all_custom = all(
            self._is_tree_item_custom(item) for item in selection
        )

        menu = tk.Menu(self.data_tree, tearoff=0)
        if all_custom:
            menu.add_command(
                label="恢复为原始分类",
                command=lambda: self._set_selected_bumps_original(
                    selection
                ),
            )
        else:
            menu.add_command(
                label="设为自定义分类",
                command=lambda: self._set_selected_bumps_category(
                    selection, BumpCategory.CUSTOM
                ),
            )
        if any(self._is_tree_item_custom(item) for item in selection):
            menu.add_command(
                label="全部恢复为原始分类",
                command=lambda: self._set_selected_bumps_original(
                    selection
                ),
            )
        try:
            menu.tk_popup(event.x_root, event.y_root)
        finally:
            menu.grab_release()

    def _is_tree_item_custom(self, item_id: str) -> bool:
        """判断 tree item 对应的 bump 是否为自定义分类。

        Args:
            item_id: Treeview 行 ID。

        Returns:
            True 表示该 bump 属于自定义分类。
        """
        values = self.data_tree.item(item_id, "values")
        if not values:
            return False
        bump_name = str(values[1])
        for bump in self.bumps:
            if bump.name == bump_name:
                return bump.category == BumpCategory.CUSTOM
        return False

    def _set_selected_bumps_category(
        self,
        selection: tuple[str, ...],
        category: BumpCategory,
    ) -> None:
        """将选中的 bump 设置为指定分类。

        Args:
            selection: Treeview 中选中的行 ID 元组。
            category: 目标分类。
        """
        names: set[str] = set()
        for item_id in selection:
            values = self.data_tree.item(item_id, "values")
            if values:
                names.add(str(values[1]))
        if not names:
            return
        for bump in self.bumps:
            if bump.name in names:
                bump.category = category
        self._update_preview()
        self._update_stats()

    def _set_selected_bumps_original(
        self,
        selection: tuple[str, ...],
    ) -> None:
        """将选中的自定义分类 bump 恢复为原始分类（按名称规则）。

        Args:
            selection: Treeview 中选中的行 ID 元组。
        """
        names: set[str] = set()
        for item_id in selection:
            values = self.data_tree.item(item_id, "values")
            if values:
                names.add(str(values[1]))
        if not names:
            return
        for bump in self.bumps:
            if bump.name in names:
                bump.category = classify_bump(bump.name)
        self._update_preview()
        self._update_stats()

    # ---- 统计刷新 ----

    def _update_stats(self) -> None:
        """更新分类统计面板。"""
        self._clear_tree(self.stats_tree)
        if not self.bumps:
            self.total_label.config(text="总计：0 个Bump")
            return

        counts = self._count_categories()
        for cat in (
            BumpCategory.GND,
            BumpCategory.POWER,
            BumpCategory.SIGNAL,
            BumpCategory.CUSTOM,
        ):
            self.stats_tree.insert(
                "",
                tk.END,
                values=(
                    CATEGORY_LABELS[cat],
                    counts.get(cat, 0),
                    get_aci_display(
                        self.config.color_map.get(cat, DEFAULT_COLOR_MAP[cat])
                    ),
                ),
            )
        self.total_label.config(text=f"总计：{len(self.bumps)} 个Bump")

    def _count_categories(self) -> dict[BumpCategory, int]:
        """统计各分类 Bump 数量。

        Returns:
            分类 → 数量的映射。
        """
        counts: dict[BumpCategory, int] = {
            BumpCategory.GND: 0,
            BumpCategory.POWER: 0,
            BumpCategory.SIGNAL: 0,
            BumpCategory.CUSTOM: 0,
        }
        for b in self.bumps:
            counts[b.category] += 1
        return counts

    # ---- 状态管理 ----

    def _update_state(self) -> None:
        """根据数据加载状态启用/禁用导出按钮。"""
        if self.bumps:
            self.export_btn.config(state=tk.NORMAL, text="导出DXF文件")
            self.status_label.config(
                text=f"已加载 {len(self.bumps)} 个Bump，准备就绪",
                fg="#16A34A",
            )
        else:
            self.export_btn.config(
                state=tk.DISABLED, text="导出DXF文件（需先导入数据）"
            )
            self.status_label.config(
                text="请先导入Excel文件", fg=_UI_TEXT_MUTED
            )
