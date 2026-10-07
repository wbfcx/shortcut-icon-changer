import os
import sys
import math
import ctypes
import hashlib
import pythoncom
import win32com.client

from PIL import Image, ImageGrab, ImageDraw
from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QGridLayout, QLineEdit, QMessageBox, QFrame, QFileDialog
)
from PyQt6.QtGui import QPixmap, QImage
from PyQt6.QtCore import Qt


IMAGE_EXTS = [".png", ".jpg", ".jpeg", ".bmp", ".webp"]
ICON_EXTS = [".ico"]

SAVE_DIR = os.path.join(os.path.expanduser("~"), "Documents", "ShortcutIcons")
os.makedirs(SAVE_DIR, exist_ok=True)

PREVIEW_PNG = os.path.join(SAVE_DIR, "preview.png")


class DropArea(QLabel):
    def __init__(self, text, callback):
        super().__init__()
        self.callback = callback
        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setText(text)
        self.setStyleSheet("""
            QLabel {
                background:#F9FAFB;
                color:#6B7280;
                border:2px dashed #CBD5E1;
                border-radius:12px;
                font-size:13px;
            }
        """)

    def mousePressEvent(self, event):
        self.setFocus()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            self.callback(urls[0].toLocalFile())

    def mark_done(self, text):
        self.setText(text)
        self.setStyleSheet("""
            QLabel {
                background:#ECFDF5;
                color:#059669;
                border:2px dashed #10B981;
                border-radius:12px;
                font-size:12px;
            }
        """)


class ShortcutIconChanger(QWidget):
    def __init__(self):
        super().__init__()

        self.shortcut_path = None
        self.original_shortcut_name = None
        self.source_image = None

        self.current_shape = "原图"
        self.polygon_sides = 6
        self.polygon_started = False

        self.setWindowTitle("快捷方式图标修改器")
        self.setFixedSize(590, 690)
        self.setStyleSheet("background:#EEF2F7;")
        self.build_ui()

    def build_ui(self):
        main = QVBoxLayout(self)
        main.setContentsMargins(22, 16, 22, 16)
        main.setSpacing(10)

        title = QLabel("快捷方式图标修改器")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size:22px;font-weight:bold;color:#111827;")
        main.addWidget(title)

        subtitle = QLabel("ICO 保存到 Documents\\ShortcutIcons")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color:#6B7280;font-size:12px;")
        main.addWidget(subtitle)

        card = QFrame()
        card.setStyleSheet("QFrame {background:white;border-radius:18px;}")
        card_layout = QVBoxLayout(card)
        card_layout.setContentsMargins(16, 16, 16, 16)
        card_layout.setSpacing(12)

        icon_row = QHBoxLayout()
        self.icon_drop_area = DropArea(
            "① 拖入图片 / 图标\n点击后 Ctrl+V 粘贴",
            self.handle_icon_file
        )
        self.icon_drop_area.setFixedHeight(78)
        icon_row.addWidget(self.icon_drop_area, 3)

        self.select_icon_btn = QPushButton("选择\n图片/图标")
        self.select_icon_btn.clicked.connect(self.select_icon_file)
        self.select_icon_btn.setFixedHeight(78)
        self.select_icon_btn.setStyleSheet(self.blue_button_style())
        icon_row.addWidget(self.select_icon_btn, 1)
        card_layout.addLayout(icon_row)

        preview_shape_row = QHBoxLayout()

        self.preview_label = QLabel("暂无图标预览")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setFixedSize(260, 230)
        self.preview_label.setStyleSheet("""
            QLabel {
                color:#9CA3AF;
                background:#F3F4F6;
                font-size:13px;
                border-radius:14px;
            }
        """)
        preview_shape_row.addWidget(self.preview_label)

        shape_grid = QGridLayout()
        shape_grid.setSpacing(7)

        self.polygon_btn = None
        shapes = [
            "原图", "圆形",
            "圆角", "超圆角",
            "椭圆", "菱形",
            "正多边形", "四角同色透明"
        ]

        for i, shape in enumerate(shapes):
            btn = QPushButton(shape)
            btn.setFixedHeight(32)
            btn.clicked.connect(lambda checked=False, s=shape: self.set_shape(s))
            btn.setStyleSheet(self.small_button_style())
            shape_grid.addWidget(btn, i // 2, i % 2)

            if shape == "正多边形":
                self.polygon_btn = btn

        preview_shape_row.addLayout(shape_grid)
        card_layout.addLayout(preview_shape_row)

        shortcut_row = QHBoxLayout()
        self.shortcut_drop_area = DropArea(
            "② 拖入快捷方式 .lnk",
            self.handle_shortcut_file
        )
        self.shortcut_drop_area.setFixedHeight(70)
        shortcut_row.addWidget(self.shortcut_drop_area, 3)

        self.select_shortcut_btn = QPushButton("选择\n快捷方式")
        self.select_shortcut_btn.clicked.connect(self.select_shortcut_file)
        self.select_shortcut_btn.setFixedHeight(70)
        self.select_shortcut_btn.setStyleSheet(self.blue_button_style())
        shortcut_row.addWidget(self.select_shortcut_btn, 1)
        card_layout.addLayout(shortcut_row)

        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("新的快捷方式名称，可不填")
        self.name_input.setFixedHeight(40)
        self.name_input.setStyleSheet("""
            QLineEdit {
                background:#F3F4F6;
                border:none;
                border-radius:10px;
                padding:9px;
                font-size:14px;
                color:#111827;
            }
        """)
        card_layout.addWidget(self.name_input)

        action_row = QHBoxLayout()

        self.apply_btn = QPushButton("应用修改")
        self.apply_btn.clicked.connect(self.apply_changes)
        self.apply_btn.setFixedHeight(44)
        self.apply_btn.setStyleSheet(self.green_button_style())
        action_row.addWidget(self.apply_btn)

        self.export_btn = QPushButton("导出 ICO")
        self.export_btn.clicked.connect(self.export_ico)
        self.export_btn.setFixedHeight(44)
        self.export_btn.setStyleSheet(self.blue_button_style())
        action_row.addWidget(self.export_btn)

        card_layout.addLayout(action_row)

        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.status_label.setStyleSheet(
            "color:#059669;background:white;font-size:12px;padding:4px;"
        )
        card_layout.addWidget(self.status_label)

        main.addWidget(card)

    def keyPressEvent(self, event):
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() in (Qt.Key.Key_V, Qt.Key.Key_C):
                if self.icon_drop_area.hasFocus():
                    self.paste_icon_from_clipboard()

    def select_icon_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择图片或图标",
            "",
            "图片或图标 (*.png *.jpg *.jpeg *.bmp *.webp *.ico)"
        )
        if path:
            self.handle_icon_file(path)

    def select_shortcut_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self,
            "选择快捷方式",
            "",
            "快捷方式 (*.lnk)"
        )
        if path:
            self.handle_shortcut_file(path)

    def paste_icon_from_clipboard(self):
        try:
            data = ImageGrab.grabclipboard()

            if data is None:
                QMessageBox.warning(self, "提示", "剪切板里没有图片或图标文件")
                return

            if isinstance(data, list):
                self.handle_icon_file(data[0])
                return

            self.source_image = data.convert("RGBA")
            self.rebuild_preview()
            self.icon_drop_area.mark_done("已粘贴图片")
            self.status_label.setText("")

        except Exception as e:
            QMessageBox.critical(self, "粘贴失败", str(e))

    def handle_icon_file(self, path):
        ext = os.path.splitext(path)[1].lower()

        if ext not in IMAGE_EXTS and ext not in ICON_EXTS:
            QMessageBox.warning(self, "格式不支持", "请选择图片或 .ico 图标文件")
            return

        try:
            self.source_image = Image.open(path).convert("RGBA")
            self.rebuild_preview()

            if ext in ICON_EXTS:
                self.icon_drop_area.mark_done(
                    f"已载入图标：{os.path.basename(path)}"
                )
            else:
                self.icon_drop_area.mark_done(
                    f"已载入图片：{os.path.basename(path)}"
                )

            self.status_label.setText("")

        except Exception as e:
            QMessageBox.critical(self, "图片读取失败", str(e))

    def handle_shortcut_file(self, path):
        if not path.lower().endswith(".lnk"):
            QMessageBox.warning(self, "格式不支持", "请选择 .lnk 快捷方式")
            return

        self.shortcut_path = path
        name = os.path.splitext(os.path.basename(path))[0]
        self.original_shortcut_name = name

        self.shortcut_drop_area.mark_done(
            f"已选择快捷方式：{os.path.basename(path)}"
        )
        self.name_input.setText(name)
        self.status_label.setText("")

    def set_shape(self, shape):
        if shape == "正多边形":
            if not self.polygon_started:
                self.polygon_sides = 6
                self.polygon_started = True
            else:
                self.polygon_sides += 1
                if self.polygon_sides > 17:
                    self.polygon_sides = 3

            self.current_shape = "正多边形"
            self.polygon_btn.setText(f"正{self.polygon_sides}边形")
            self.status_label.setText(f"当前形状：正{self.polygon_sides}边形")
        else:
            self.current_shape = shape
            self.status_label.setText(f"当前形状：{shape}")

        if self.source_image is not None:
            self.rebuild_preview()

    def rebuild_preview(self):
        final_img = self.build_final_image()
        preview = self.make_checkerboard_preview(final_img, 230)
        preview.save(PREVIEW_PNG, "PNG")
        self.show_preview()

    def build_final_image(self):
        img = self.source_image.convert("RGBA")

        if self.current_shape == "四角同色透明":
            img = self.make_corner_same_color_transparent_on_original(img)

        return self.apply_shape(img)

    def make_square_canvas(self, img):
        size = max(img.width, img.height)
        canvas = Image.new("RGBA", (size, size), (0, 0, 0, 0))

        x = (size - img.width) // 2
        y = (size - img.height) // 2

        canvas.paste(img, (x, y), img)

        return canvas

    def apply_shape(self, img):
        canvas = self.make_square_canvas(img)
        shape = self.current_shape

        if shape in ["原图", "四角同色透明"]:
            return canvas

        size = canvas.size[0]
        scale = 4
        big = size * scale

        mask = Image.new("L", (big, big), 0)
        draw = ImageDraw.Draw(mask)
        m = int(big * 0.04)

        if shape == "圆形":
            draw.ellipse((m, m, big - m, big - m), fill=255)

        elif shape == "圆角":
            draw.rounded_rectangle(
                (m, m, big - m, big - m),
                radius=int(big * 0.18),
                fill=255
            )

        elif shape == "超圆角":
            draw.rounded_rectangle(
                (m, m, big - m, big - m),
                radius=int(big * 0.34),
                fill=255
            )

        elif shape == "椭圆":
            draw.ellipse(
                (
                    int(big * 0.08),
                    int(big * 0.18),
                    int(big * 0.92),
                    int(big * 0.82)
                ),
                fill=255
            )

        elif shape == "菱形":
            draw.polygon(
                [
                    (big // 2, m),
                    (big - m, big // 2),
                    (big // 2, big - m),
                    (m, big // 2)
                ],
                fill=255
            )

        elif shape == "正多边形":
            cx = big / 2
            cy = big / 2
            r = big / 2 - m
            points = []

            for i in range(self.polygon_sides):
                angle = math.radians(360 * i / self.polygon_sides - 90)
                points.append((cx + r * math.cos(angle), cy + r * math.sin(angle)))

            draw.polygon(points, fill=255)

        mask = mask.resize((size, size), Image.Resampling.LANCZOS)

        result = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        result.paste(canvas, (0, 0), mask)

        return result

    def make_corner_same_color_transparent_on_original(self, img):
        img = img.convert("RGBA")
        pixels = img.load()
        w, h = img.size

        c1 = pixels[0, 0]
        c2 = pixels[w - 1, 0]
        c3 = pixels[0, h - 1]
        c4 = pixels[w - 1, h - 1]

        if not (c1 == c2 == c3 == c4):
            self.status_label.setText("四角颜色不完全相同，未透明处理")
            return img

        target = c1
        result = Image.new("RGBA", img.size, (0, 0, 0, 0))
        rp = result.load()

        for y in range(h):
            for x in range(w):
                px = pixels[x, y]

                if px == target:
                    rp[x, y] = (px[0], px[1], px[2], 0)
                else:
                    rp[x, y] = px

        self.status_label.setText("四角同色透明已生效")
        return result

    def alpha_is_empty(self, img):
        return img.getchannel("A").getbbox() is None

    def get_saved_ico_path(self):
        os.makedirs(SAVE_DIR, exist_ok=True)

        shortcut_name = os.path.splitext(os.path.basename(self.shortcut_path))[0]
        shortcut_hash = hashlib.md5(
            self.shortcut_path.encode("utf-8", errors="ignore")
        ).hexdigest()[:8]

        safe_name = "".join(
            c for c in shortcut_name
            if c not in r'\/:*?"<>|'
        ).strip()

        if not safe_name:
            safe_name = "shortcut"

        return os.path.join(SAVE_DIR, f"{safe_name}_{shortcut_hash}.ico")

    def save_icon_to_library(self):
        os.makedirs(SAVE_DIR, exist_ok=True)

        final_img = self.build_final_image()

        if self.alpha_is_empty(final_img):
            raise ValueError("处理后的图标是全透明的，已取消应用")

        ico_path = self.get_saved_ico_path()

        final_img.save(
            ico_path,
            format="ICO",
            sizes=[
                (256, 256),
                (128, 128),
                (64, 64),
                (48, 48),
                (32, 32),
                (16, 16)
            ]
        )

        if not os.path.exists(ico_path) or os.path.getsize(ico_path) == 0:
            raise ValueError(f"ICO 文件生成失败：{ico_path}")

        return ico_path

    def make_checkerboard_preview(self, img, preview_size):
        img = img.copy()
        img.thumbnail((preview_size, preview_size), Image.Resampling.LANCZOS)

        board = Image.new(
            "RGBA",
            (preview_size, preview_size),
            (245, 245, 245, 255)
        )
        draw = ImageDraw.Draw(board)

        block = 12
        for y in range(0, preview_size, block):
            for x in range(0, preview_size, block):
                if (x // block + y // block) % 2 == 0:
                    draw.rectangle(
                        (x, y, x + block, y + block),
                        fill=(225, 225, 225, 255)
                    )

        x = (preview_size - img.width) // 2
        y = (preview_size - img.height) // 2

        board.alpha_composite(img, (x, y))
        return board

    def show_preview(self):
        if os.path.exists(PREVIEW_PNG):
            image = QImage(PREVIEW_PNG)
            self.preview_label.setPixmap(QPixmap.fromImage(image))

    def apply_changes(self):
        if self.source_image is None:
            QMessageBox.warning(self, "提示", "请先选择图片或图标")
            return

        if not self.shortcut_path:
            QMessageBox.warning(self, "提示", "请先选择快捷方式")
            return

        try:
            ico_path = self.save_icon_to_library()

            pythoncom.CoInitialize()

            wsh = win32com.client.Dispatch("WScript.Shell")
            shortcut = wsh.CreateShortcut(self.shortcut_path)
            shortcut.IconLocation = f"{ico_path},0"
            shortcut.Save()

            new_name = self.name_input.text().strip()

            if new_name and new_name != self.original_shortcut_name:
                folder = os.path.dirname(self.shortcut_path)
                new_path = os.path.join(folder, new_name + ".lnk")

                if not os.path.exists(new_path):
                    os.rename(self.shortcut_path, new_path)
                    self.shortcut_path = new_path
                    self.original_shortcut_name = new_name
                    self.shortcut_drop_area.mark_done(
                        f"已选择快捷方式：{os.path.basename(new_path)}"
                    )

            self.refresh_icon_cache()

            self.status_label.setText(f"✔ 图标已修改，ICO保存到：{ico_path}")

        except Exception as e:
            QMessageBox.critical(self, "修改失败", str(e))

    def export_ico(self):
        if self.source_image is None:
            QMessageBox.warning(self, "提示", "请先选择图片或图标")
            return

        path, _ = QFileDialog.getSaveFileName(
            self,
            "导出 ICO",
            "custom_icon.ico",
            "ICO 图标 (*.ico)"
        )

        if not path:
            return

        if not path.lower().endswith(".ico"):
            path += ".ico"

        try:
            final_img = self.build_final_image()

            if self.alpha_is_empty(final_img):
                QMessageBox.warning(self, "提示", "处理后的图标是全透明的，无法导出")
                return

            final_img.save(
                path,
                format="ICO",
                sizes=[
                    (256, 256),
                    (128, 128),
                    (64, 64),
                    (48, 48),
                    (32, 32),
                    (16, 16)
                ]
            )

            self.status_label.setText(f"✔ 已导出 ICO：{os.path.basename(path)}")

        except Exception as e:
            QMessageBox.critical(self, "导出失败", str(e))

    def refresh_icon_cache(self):
        try:
            ctypes.windll.shell32.SHChangeNotify(
                0x08000000,
                0x0000,
                None,
                None
            )
        except Exception:
            pass

    def blue_button_style(self):
        return """
            QPushButton {
                background:#2563EB;
                color:white;
                border:none;
                border-radius:12px;
                font-size:13px;
                font-weight:bold;
            }
            QPushButton:hover {
                background:#1D4ED8;
            }
        """

    def green_button_style(self):
        return """
            QPushButton {
                background:#10B981;
                color:white;
                border:none;
                border-radius:12px;
                font-size:15px;
                font-weight:bold;
            }
            QPushButton:hover {
                background:#059669;
            }
        """

    def small_button_style(self):
        return """
            QPushButton {
                background:#F3F4F6;
                color:#374151;
                border:none;
                border-radius:8px;
                font-size:12px;
            }
            QPushButton:hover {
                background:#DBEAFE;
                color:#1D4ED8;
            }
        """


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ShortcutIconChanger()
    window.show()
    sys.exit(app.exec())