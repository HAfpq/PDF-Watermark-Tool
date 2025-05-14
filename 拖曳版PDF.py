import os
import sys
from io import BytesIO
from PIL import Image
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QLabel, QPushButton, QLineEdit,
    QComboBox, QSpinBox, QSlider, QProgressBar, QFileDialog,
    QHBoxLayout, QVBoxLayout, QSplitter, QScrollArea, QMessageBox,
    QGroupBox, QFormLayout, QSizePolicy, QColorDialog
)
from PyQt5.QtGui import QPixmap, QPainter, QFontDatabase, QFont, QFontMetrics, QColor
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# 字体配置
FONT_OPTIONS = {
    "微软雅黑": ("MicrosoftYaHei", "C:/Windows/Fonts/msyh.ttc"),
    "宋体": ("SimSun", "C:/Windows/Fonts/simsun.ttc"),
    "黑体": ("SimHei", "C:/Windows/Fonts/simhei.ttf"),
    "楷体": ("KaiTi", "C:/Windows/Fonts/simkai.ttf"),
    "仿宋": ("FangSong", "C:/Windows/Fonts/simfang.ttf"),
    "Arial": ("Arial", "C:/Windows/Fonts/arial.ttf"),
    "Times New Roman": ("TimesNewRoman", "C:/Windows/Fonts/times.ttf"),
    "Courier New": ("CourierNew", "C:/Windows/Fonts/cour.ttf")
}

class WatermarkThread(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)

    def __init__(
        self, pdf_list, text, font_name, logo_path,
        alpha, h_count, v_count, text_pos, logo_pos,
        angle, text_size_pct, logo_size_pct, text_color, parent=None
    ):
        super().__init__(parent)
        self.pdf_list       = pdf_list
        self.text           = text
        self.font_name      = font_name
        self.logo_path      = logo_path
        self.alpha          = alpha
        self.h_count        = h_count
        self.v_count        = v_count
        self.text_pos       = text_pos
        self.logo_pos       = logo_pos
        self.angle          = angle
        self.text_size_pct  = text_size_pct
        self.logo_size_pct  = logo_size_pct
        self.text_color     = text_color
        self.parent         = parent

    def run(self):
        out_dir = os.path.join(os.path.expanduser('~'), 'Desktop', 'pdf_watermark_output')
        os.makedirs(out_dir, exist_ok=True)
        log_file = os.path.join(out_dir, 'error_log.txt')
        with open(log_file, 'w', encoding='utf-8') as log:
            for idx, inp in enumerate(self.pdf_list, 1):
                fn = os.path.basename(inp)
                out = os.path.join(out_dir, f"wm_{fn}")
                try:    self.parent._add_watermark(
                        inp, out,
                        self.text, self.font_name,
                        self.logo_path, self.alpha,
                        self.h_count, self.v_count,
                        self.text_pos, self.logo_pos,
                        self.angle,
                        self.text_size_pct, self.logo_size_pct, self.text_color
                    )
                except Exception as e:
                    log.write(f"{fn} failed: {e}\n")
                self.progress.emit(idx)
        self.finished.emit(out_dir)

class PDFWatermarkerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.resize(1000, 700)
        self.folder_path     = ''
        self.logo_path       = ''
        self.font_cache      = {}
        self.text_size_pct   = 100  # 文字大小百分比
        self.logo_size_pct   = 100  # Logo 大小百分比
        self.text_color      = QColor(0, 0, 0)  # 默认文字颜色：黑色
        self.worker          = None

        self._init_ui()
        self._connect_signals()
        self.update_preview()
        self.setAcceptDrops(True)

    def _init_ui(self):
        # ---- 全局容器 & 样式 ----
        central = QWidget(self)
        central.setStyleSheet("background-color: #f7f9fb; font-family: 'Microsoft YaHei';")
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(16, 16, 16, 16)
        main_layout.setSpacing(16)

        # ---- 分割器 ----
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(4)
        splitter.setStyleSheet("QSplitter::handle { background-color: #d9d9d9; }")
        main_layout.addWidget(splitter)

        # ---- 样式定义 ----
        group_style = """
        QGroupBox {
            font: 600 16px 'Microsoft YaHei';
            margin-top: 26px;
            border: 1px solid #a0aec0;
            border-radius: 10px;
            background-color: #e6f7ff;
            padding: 10px;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            subcontrol-position: top left;
            padding: 0 10px;
            color: #1f1f1f;
            background-color: #e6f7ff;
            font-weight: bold;
        }
        """
        btn_style = """
        QPushButton {
            background-color: #7cb8e6;
            color: white;
            border: none;
            border-radius: 6px;
            font: bold 15px 'Microsoft YaHei';
            padding: 6px 14px;
            box-shadow: 0px 2px 4px rgba(0,0,0,0.15);
            transition: background-color 0.3s ease;
        }
        QPushButton:hover {
            background-color: #40a9ff;
        }
        QPushButton:pressed {
            background-color: #096dd9;
        }
        """
        slider_style = """
        QSlider::groove:horizontal {
            height: 8px;
            background: #e0e0e0;
            border-radius: 4px;
        }
        QSlider::handle:horizontal {
            background: qlineargradient(
                spread:pad, x1:0, y1:0, x2:1, y2:0,
                stop:0 #40a9ff, stop:1 #1890ff
            );
            width: 16px;
            margin: -4px 0;
            border-radius: 8px;
        }
        """
        progress_style = """
        QProgressBar {
            border: 1px solid #ccc; border-radius: 6px;
            background: #f5f5f5; height: 20px;
        }
        QProgressBar::chunk {
            background-color: #1890ff; border-radius: 6px;
        }
        """

        # ---- 左侧参数面板 ----
        panel = QWidget()
        panel.setStyleSheet(
            "background-color: #e6f7ff;"
            "border: 1px solid #d9d9d9;"
            "border-radius: 8px;"
        )
        panel_layout = QVBoxLayout(panel)
        panel_layout.setContentsMargins(14, 22, 14, 14)
        panel_layout.setSpacing(14)

        # 1. 选择 PDF 文件夹
        gb_folder = QGroupBox("📂 选择文件夹")
        gb_folder.setStyleSheet(group_style)
        form_folder = QHBoxLayout(gb_folder)
        form_folder.setContentsMargins(10, 6, 10, 6)
        self.btn_folder = QPushButton("🔍 浏览文件夹")
        self.btn_folder.setFixedHeight(32)
        self.btn_folder.setStyleSheet(btn_style)
        self.lbl_folder = QLabel("⚪未选择文件夹")
        self.lbl_folder.setStyleSheet("color: #333333;")
        form_folder.addWidget(self.btn_folder)
        form_folder.addSpacing(8)
        form_folder.addWidget(self.lbl_folder)
        panel_layout.addWidget(gb_folder)

        # 📝 文本与 Logo 设置分组
        gb_text_logo = QGroupBox("📝 文本与Logo ")
        gb_text_logo.setStyleSheet(group_style)
        layout_text_logo = QVBoxLayout(gb_text_logo)
        layout_text_logo.setContentsMargins(10, 6, 10, 6)

        # 🔤 文本输入
        hlayout_text = QHBoxLayout()
        self.edit_text = QLineEdit("研汇工坊")
        self.edit_text.setFixedHeight(30)
        self.edit_text.setStyleSheet(
            "padding: 4px; border: 1px solid #d9d9d9; border-radius: 4px;"
        )
        hlayout_text.addWidget(QLabel("🔤 文本"))
        hlayout_text.addWidget(self.edit_text)
        layout_text_logo.addLayout(hlayout_text)

        # 🏷️ 字体和 Logo
        hlayout_font_logo = QHBoxLayout()
        # 字体选择
        hlayout_font_logo.addWidget(QLabel("🖋️ 字体"))
        self.combo_font = QComboBox()
        self.combo_font.addItems(FONT_OPTIONS.keys())
        self.combo_font.setFixedHeight(30)
        self.combo_font.setStyleSheet(
            "border:1px solid #d9d9d9; border-radius:4px; padding:4px;"
        )
        hlayout_font_logo.addWidget(self.combo_font)

        # Logo 选择按钮
        hlayout_font_logo.addStretch()
        self.btn_logo = QPushButton("🖼️ 选择Logo图片")
        self.btn_logo.setFixedHeight(32)
        self.btn_logo.setStyleSheet(btn_style)
        hlayout_font_logo.addWidget(self.btn_logo)

        layout_text_logo.addLayout(hlayout_font_logo)

        # 添加到主布局
        panel_layout.addWidget(gb_text_logo)

        # 💬 水印样式设置分组
        gb_style = QGroupBox("💬 水印样式")
        gb_style.setStyleSheet(group_style)
        style_layout = QVBoxLayout(gb_style)
        style_layout.setContentsMargins(10, 6, 10, 6)

        # 文字大小滑块
        hlayout_size = QHBoxLayout()
        label_size = QLabel("🖱️ 文字大小")
        self.slider_text_size = QSlider(Qt.Horizontal)
        self.slider_text_size.setRange(10, 200)
        self.slider_text_size.setValue(self.text_size_pct)
        hlayout_size.addWidget(label_size)
        hlayout_size.addWidget(self.slider_text_size)
        style_layout.addLayout(hlayout_size)

        # 文字颜色选择器
        hlayout_color = QHBoxLayout()
        self.btn_text_color = QPushButton("🎨 文字颜色")
        self.lbl_color_preview = QLabel()
        self.lbl_color_preview.setFixedSize(24, 24)
        self.lbl_color_preview.setStyleSheet(f"background-color: {self.text_color.name()};")
        hlayout_color.addWidget(self.btn_text_color)
        hlayout_color.addWidget(self.lbl_color_preview)
        style_layout.addLayout(hlayout_color)

        # Logo 大小滑块
        hlayout_logo = QHBoxLayout()
        label_logo_size = QLabel("🖱️ Logo 大小")
        self.slider_logo_size = QSlider(Qt.Horizontal)
        self.slider_logo_size.setRange(10, 200)
        self.slider_logo_size.setValue(self.logo_size_pct)
        hlayout_logo.addWidget(label_logo_size)
        hlayout_logo.addWidget(self.slider_logo_size)
        style_layout.addLayout(hlayout_logo)

        # 透明度滑块
        hlayout_opacity = QHBoxLayout()
        label_opacity = QLabel("🌫️ 透明度")
        self.slider_alpha = QSlider(Qt.Horizontal)
        self.slider_alpha.setRange(0, 100)
        self.slider_alpha.setValue(20)
        self.slider_alpha.setStyleSheet(slider_style)
        hlayout_opacity.addWidget(label_opacity)
        hlayout_opacity.addWidget(self.slider_alpha)
        style_layout.addLayout(hlayout_opacity)

        # 添加到主面板布局
        panel_layout.addWidget(gb_style)

        # 📐 布局与排列设置 分组
        gb_layout = QGroupBox("📐 布局与排列")
        gb_layout.setStyleSheet(group_style)
        layout_all = QVBoxLayout(gb_layout)
        layout_all.setContentsMargins(10, 6, 10, 6)

        # 🔢 水印数量
        layout_cnt = QHBoxLayout()
        self.spin_h = QSpinBox();
        self.spin_h.setRange(1, 10);
        self.spin_h.setValue(1);
        self.spin_h.setFixedWidth(60)
        self.spin_v = QSpinBox();
        self.spin_v.setRange(1, 10);
        self.spin_v.setValue(1);
        self.spin_v.setFixedWidth(60)
        layout_cnt.addWidget(QLabel("↔️ 横行数量"));
        layout_cnt.addWidget(self.spin_h)
        layout_cnt.addSpacing(12)
        layout_cnt.addWidget(QLabel("↕️ 纵行数量"));
        layout_cnt.addWidget(self.spin_v)
        layout_all.addLayout(layout_cnt)

        # 🚩 水印位置
        layout_pos = QFormLayout()
        self.combo_text_pos = QComboBox();
        self.combo_text_pos.addItems(["中心", "左上", "右上", "左下", "右下"])
        self.combo_logo_pos = QComboBox();
        self.combo_logo_pos.addItems(["中心", "左上", "右上", "左下", "右下"])
        self.combo_text_pos.setFixedHeight(25);
        self.combo_logo_pos.setFixedHeight(25)
        layout_pos.addRow("📝 文本位置", self.combo_text_pos)
        layout_pos.addRow("📊 Logo 位置", self.combo_logo_pos)
        layout_all.addLayout(layout_pos)

        # 🔄 旋转角度
        layout_ang = QHBoxLayout()
        self.spin_angle = QSpinBox();
        self.spin_angle.setRange(-90, 90);
        self.spin_angle.setValue(20);
        self.spin_angle.setFixedWidth(100)
        layout_ang.addWidget(QLabel("🔺 角度 (°):"))
        layout_ang.addWidget(self.spin_angle)
        layout_all.addLayout(layout_ang)

        # 添加到主布局
        panel_layout.addWidget(gb_layout)

        # 操作按钮 & 进度条
        self.btn_start = QPushButton("▶️ 开始添加"); self.btn_start.setFixedHeight(34); self.btn_start.setStyleSheet(btn_style)
        self.btn_clear = QPushButton("🔙 重置设置");     self.btn_clear.setFixedHeight(34); self.btn_clear.setStyleSheet(btn_style)
        self.progress  = QProgressBar();              self.progress.setFixedHeight(14); self.progress.setStyleSheet(progress_style)
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)  # 两按钮之间的间距
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addWidget(self.btn_start)
        panel_layout.addLayout(btn_layout)
        panel_layout.addWidget(self.progress)
        panel_layout.addStretch()
        splitter.addWidget(panel)

        # ---- 右侧预览面板 ----
        preview_panel = QWidget()
        preview_panel.setStyleSheet(
            "background-color: #ffffff;"
            "border: 1px solid #d9d9d9;"
            "border-radius: 8px;"
        )
        layout_preview = QVBoxLayout(preview_panel)
        layout_preview.setContentsMargins(14, 14, 14, 14)
        layout_preview.setSpacing(14)
        lbl_preview = QLabel("👀 实时预览区")
        lbl_preview.setAlignment(Qt.AlignCenter)
        lbl_preview.setStyleSheet("font: bold 18px 'Microsoft YaHei'; color: #096dd9;")
        layout_preview.addWidget(lbl_preview)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.preview_label = QLabel()
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.preview_label.setMinimumSize(200, 200)
        self.scroll_area.setWidget(self.preview_label)
        layout_preview.addWidget(self.scroll_area)
        splitter.addWidget(preview_panel)

        # ---- 左右面板比例 ----
        splitter.setStretchFactor(0, 2)
        splitter.setStretchFactor(1, 8)

    def _connect_signals(self):
        self.btn_folder.clicked.connect(self.browse_folder)
        self.btn_logo.clicked.connect(self.choose_logo)
        self.btn_start.clicked.connect(self.start_process)
        self.btn_clear.clicked.connect(self.clear_settings)

        widgets = [self.edit_text, self.combo_font, self.slider_alpha,
                   self.spin_h, self.spin_v, self.combo_text_pos,
                   self.combo_logo_pos, self.spin_angle]
        # 现有 widgets 列表后面添加：
        widgets += [self.slider_text_size, self.slider_logo_size]
        for w in (self.slider_text_size, self.slider_logo_size):
            w.valueChanged.connect(self.update_preview)

        # 文字颜色按钮
        self.btn_text_color.clicked.connect(self.choose_text_color)

        for w in widgets:
            if hasattr(w, 'textChanged'):
                w.textChanged.connect(self.update_preview)
            if hasattr(w, 'currentIndexChanged'):
                w.currentIndexChanged.connect(self.update_preview)
            if hasattr(w, 'valueChanged'):
                w.valueChanged.connect(self.update_preview)

    def dragEnterEvent(self, event):
        # 如果拖入的是本地文件或文件夹，则接受
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        # 只取第一个拖入项目
        url = event.mimeData().urls()[0]
        path = url.toLocalFile()
        # 如果是文件夹
        if os.path.isdir(path):
            self.dropped_file = None
            self.folder_path = path
            self.lbl_folder.setText(path)
        # 如果是单个 PDF 文件
        elif os.path.isfile(path) and path.lower().endswith('.pdf'):
            self.dropped_file = path
            # 也设置 folder_path 为所在目录，以免后续列表扫描报错
            self.folder_path = os.path.dirname(path)
            # 在标签中显示文件名
            self.lbl_folder.setText(os.path.basename(path))
        else:
            # 非 PDF 或文件夹，忽略
            return
        # 更新进度条归零、重置预览
        self.progress.setValue(0)
        self.preview_label.clear()


    def browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "📂 选择文件夹")
        if folder:
            self.folder_path = folder
            self.lbl_folder.setText(folder)

    def choose_text_color(self):
        color = QColorDialog.getColor(initial=self.text_color, parent=self, title="🎨 文字颜色")
        if color.isValid():
            self.text_color = color
            self.lbl_color_preview.setStyleSheet(f"background-color: {color.name()};")
            self.update_preview()

    def choose_logo(self):
        path, _ = QFileDialog.getOpenFileName(self, "🖼️ 选择Logo图片", filter="Images (*.png *.jpg *.jpeg)")
        if path:
            self.logo_path = path
            self.update_preview()

    def clear_settings(self):
        self.lbl_folder.setText('未选择文件夹')
        self.edit_text.setText('研汇工坊')
        self.slider_alpha.setValue(20)
        self.spin_h.setValue(1)
        self.spin_v.setValue(1)
        self.combo_text_pos.setCurrentIndex(0)
        self.combo_logo_pos.setCurrentIndex(0)
        self.spin_angle.setValue(30)
        self.progress.setValue(0)
        self.preview_label.clear()
        self.dropped_file = None
    def update_preview(self):
        # A4 尺寸 pt 转像素
        w_pt, h_pt = A4
        scale = 0.4  # 初始缩放比例
        w_px, h_px = int(w_pt * scale), int(h_pt * scale)
        # 获取预览框的实际尺寸
        w_1b1 = self.preview_label.width()
        h_1b1 = self.preview_label.height()
        # 预防不合理的尺寸
        if w_1b1 <= 0 or h_1b1 <= 0:
            return
        # 计算缩放比例，保持宽高比例
        scale = min(w_1b1 / w_px, h_1b1 / h_px)
        # 根据缩放比例调整尺寸
        w_px, h_px = int(w_px * scale), int(h_px * scale)
        # 创建 Pixmap
        pixmap = QPixmap(w_px, h_px)
        pixmap.fill(Qt.white)
        # 创建画笔对象
        painter = QPainter(pixmap)
        painter.setOpacity(self.slider_alpha.value() / 100)
        painter.setPen(self.text_color)

        # 加载字体
        font_key = self.combo_font.currentText()
        _, font_path = FONT_OPTIONS[font_key]
        if font_key not in self.font_cache:
            fid = QFontDatabase.addApplicationFont(font_path)
            fam = QFontDatabase.applicationFontFamilies(fid)
            self.font_cache[font_key] = fam[0] if fam else ''
        # 动态字体大小：字体大小 = 预览画布宽度 * 比例
        font_ratio = 0.05  # 5%，可以根据你界面自行调整
        base_size = w_px * font_ratio
        font_size = max(1, int(base_size * (self.slider_text_size.value() / 100.0)))
        font = QFont(self.font_cache[font_key], font_size)
        painter.setFont(font)

        # … 计算好 w_px, h_px、设置好 font 之后 …
        # 取字体度量
        metrics = QFontMetrics(font)
        text = self.edit_text.text()
        text_width = metrics.horizontalAdvance(text)
        text_height = metrics.height()
        # —— 拉取用户设置 ——
        h_count  = self.spin_h.value()
        v_count  = self.spin_v.value()
        angle    = self.spin_angle.value()
        text_pos = self.combo_text_pos.currentText()

        # —— 定义偏移量表 ——
        offsets = {
            '左上': (-w_px/4, -h_px/4),
            '右上': ( w_px/4, -h_px/4),
            '左下': (-w_px/4,  h_px/4),
            '右下': ( w_px/4,  h_px/4),
            '中心': (0, 0),
        }

        # 绘制文本水印
        for i in range(1, h_count + 1):
            for j in range(1, v_count + 1):
                x = i * w_px / (h_count + 1)
                y = j * h_px / (v_count + 1)
                ox, oy = offsets[text_pos]
                painter.save()
                painter.translate(int(x + ox), int(y + oy))
                painter.rotate(angle)
                # 以文字中心为原点，向左/向上偏移一半宽高
                painter.drawText(
                    int(-text_width / 2),
                    int(text_height / 2),
                    text
                )
                painter.restore()

        # 绘制 Logo 水印
        if self.logo_path:
            logo = QPixmap(self.logo_path)
            if not logo.isNull():
                logo_size = int(100 * scale * (self.slider_logo_size.value() / 100.0))
                logo = logo.scaled(logo_size, logo_size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
                pos = self.combo_logo_pos.currentText()
                coords = {
                    '左上': (0, 0), '右上': (w_px - logo.width(), 0),
                    '左下': (0, h_px - logo.height()), '右下': (w_px - logo.width(), h_px - logo.height()),
                    '中心': ((w_px - logo.width()) // 2, (h_px - logo.height()) // 2)
                }
                x_l, y_l = coords[pos]
                painter.drawPixmap(x_l, y_l, logo)

        painter.end()

        # 更新预览标签
        self.preview_label.setPixmap(pixmap.scaled(
            self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
        ))

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_preview()

    def start_process(self):
        # 校验：必须有单文件（dropped_file）或文件夹
        if not ((hasattr(self, 'dropped_file') and self.dropped_file) or self.folder_path):
            QMessageBox.warning(self, "错误", "请先选择 PDF 文件夹 或 拖入单个 PDF")
            return
        # 水印文字不能为空
        text = self.edit_text.text().strip()
        if not text:
            QMessageBox.warning(self, "错误", "请输入水印内容")
            return

        # 注册字体（保持原有逻辑）
        key = self.combo_font.currentText()
        font_name, font_path = FONT_OPTIONS[key]
        if not os.path.exists(font_path):
            QMessageBox.warning(self, "错误", f"找不到字体文件: {font_path}")
            return
        pdfmetrics.registerFont(TTFont(font_name, font_path))

        # 准备要处理的 PDF 列表
        if hasattr(self, 'dropped_file') and self.dropped_file:
            pdfs = [self.dropped_file]
        else:
            pdfs = [
                os.path.join(self.folder_path, f)
                for f in os.listdir(self.folder_path)
                if f.lower().endswith('.pdf')
            ]

        if not pdfs:
            QMessageBox.warning(self, "错误", "没有找到可处理的 PDF")
            return

        # 禁用按钮，防止重复点击
        self.btn_start.setEnabled(False)
        self.btn_clear.setEnabled(False)

        # 设置进度条
        self.progress.setMaximum(len(pdfs))
        self.progress.setValue(0)

        # 创建并启动后台线程
        self.worker = WatermarkThread(
            pdf_list=pdfs,
            text=text,
            font_name=font_name,
            logo_path=self.logo_path,
            alpha=self.slider_alpha.value() / 100,
            h_count=self.spin_h.value(),
            v_count=self.spin_v.value(),
            text_pos=self.combo_text_pos.currentText(),
            logo_pos=self.combo_logo_pos.currentText(),
            angle=self.spin_angle.value(),
            text_size_pct=self.slider_text_size.value(),
            logo_size_pct=self.slider_logo_size.value(),
            text_color=self.text_color,
            parent=self
        )
        # 信号绑定
        self.worker.progress.connect(self.progress.setValue)
        self.worker.finished.connect(self._on_finished)
        # 启动
        self.worker.start()

    def _on_finished(self, out_dir):
        QMessageBox.information(self, "完成", f"处理完成，输出目录: {out_dir}")
        self.btn_start.setEnabled(True)
        self.btn_clear.setEnabled(True)

    def _create_watermark_page(self,
                                   text, font_name, logo_path,
                                   alpha, h_count, v_count,
                                   text_pos, logo_pos, angle,
                                   text_size_pct, logo_size_pct, text_color):
        packet = BytesIO()
        can = canvas.Canvas(packet, pagesize=A4)

        # --- 文字：动态字号 & 颜色 ---
        base_pt = 40
        pt_size = base_pt * (text_size_pct / 100.0)
        can.setFont(font_name, pt_size)
        # reportlab 颜色需要 0–1 浮点
        r, g, b, _ = text_color.getRgbF()
        can.setFillColor(Color(r, g, b, alpha))
        w, h = A4
        # 文字宽度（pt 单位）
        text_width = stringWidth(text, font_name, pt_size)
        text_height = pt_size  # 近似行高就是字号

        offsets = {
            '左上': (-w / 4, h / 4),
            '右上': (w / 4, h / 4),
            '左下': (-w / 4, -h / 4),
            '右下': (w / 4, -h / 4),
            '中心': (0, 0)
        }

        for i in range(1, h_count + 1):
            for j in range(1, v_count + 1):
                cx = i * w / (h_count + 1) + offsets[text_pos][0]
                cy = j * h / (v_count + 1) + offsets[text_pos][1]
                can.saveState()
                can.translate(cx, cy)
                can.rotate(-angle)
                # 居中绘制：左移一半宽度，上移半行高
                can.drawString(-text_width / 2, -text_height / 2, text)
                can.restoreState()

        if logo_path and os.path.exists(logo_path):
            img = Image.open(logo_path)
            scale_factor = 0.2 * (logo_size_pct / 100.0)
            img_width, img_height = img.size
            img_width *= scale_factor
            img_height *= scale_factor

            coords = {
                '左上': (0, h - img_height),
                '右上': (w - img_width, h - img_height),
                '左下': (0, 0),
                '右下': (w - img_width, 0),
                '中心': ((w - img_width) / 2, (h - img_height) / 2),
            }
            x, y = coords[logo_pos]

            can.drawImage(logo_path, x, y, width=img_width, height=img_height, preserveAspectRatio=True, mask='auto')

        can.showPage()
        can.save()
        packet.seek(0)
        return packet

    def _add_watermark(self, inp_path, out_path,
                           text, font_name, logo_path,
                           alpha, h_count, v_count,
                           text_pos, logo_pos, angle,
                           text_size_pct, logo_size_pct, text_color):

        output = PdfWriter()
        reader = PdfReader(inp_path)

        watermark = self._create_watermark_page(
            text, font_name, logo_path,
            alpha, h_count, v_count,
            text_pos, logo_pos, angle,
            text_size_pct, logo_size_pct, text_color
        )
        watermark_pdf = PdfReader(watermark)
        watermark_page = watermark_pdf.pages[0]

        for page in reader.pages:
            page.merge_page(watermark_page)
            output.add_page(page)

        with open(out_path, "wb") as f:
            output.write(f)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFWatermarkerApp()
    window.show()
    sys.exit(app.exec_())