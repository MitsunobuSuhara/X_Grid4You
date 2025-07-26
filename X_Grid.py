import sys
import os
import fiona
import math
from fiona.errors import FionaError
import sqlite3
import xml.etree.ElementTree as ET
from PyQt6.QtWidgets import (
    QApplication, QGraphicsView, QGraphicsScene, QMainWindow, QPushButton,
    QFileDialog, QMessageBox, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QListWidget, QListWidgetItem, QDialog, QDialogButtonBox, QCheckBox, QFrame,
    QLineEdit
)
from PyQt6.QtCore import Qt, QRectF, QPointF, pyqtSignal, QMarginsF, QSizeF, QPoint
from PyQt6.QtGui import (
    QColor, QPen, QBrush, QFont, QPolygonF, QPainter,
    QCursor, QPainterPath, QPageLayout, QPageSize, QFontMetrics
)
from PyQt6.QtPrintSupport import QPrinter

from shapely.geometry import Polygon, MultiPolygon, shape, box
from shapely.ops import unary_union
from shapely.affinity import rotate

DEFAULT_STYLE_INFO = {
    'fill_color': QColor(Qt.GlobalColor.transparent),
    'line_color': QColor(0, 0, 0),
    'line_width': 0.3,
    'pen_style': Qt.PenStyle.SolidLine,
    'dash_pattern': [],
    'line_width_unit': 'MM'
}

def _parse_any_color_string(color_value, default_color=QColor(0, 0, 0)):
    if not color_value or not isinstance(color_value, str): return default_color
    color_str = color_value.strip()
    if ',' in color_str:
        try:
            parts = [int(p.strip()) for p in color_str.split(',') if p.strip().isdigit()]
            if len(parts) == 3: return QColor(parts[0], parts[1], parts[2])
            if len(parts) == 4: return QColor(parts[0], parts[1], parts[2], parts[3])
        except (ValueError, IndexError): pass
    if QColor.isValidColor(color_str): return QColor(color_str)
    return default_color

class LayerSelectionDialog(QDialog):
    def __init__(self, layer_names, parent=None):
        super().__init__(parent)
        self.setWindowTitle("レイヤを選択")
        self.layout = QVBoxLayout(self)
        self.checkboxes = []
        for name in layer_names:
            if name in ['layer_styles', 'gpkg_layer_styles']: continue
            cb = QCheckBox(name); cb.setChecked(True)
            self.checkboxes.append(cb); self.layout.addWidget(cb)
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept); button_box.rejected.connect(self.reject)
        self.layout.addWidget(button_box)
        self.setStyleSheet("""
            QDialog { background-color: #F0F0F0; } 
            QCheckBox { color: #000000; }
            QCheckBox::indicator { border: 1px solid #000000; background-color: #FFFFFF; width: 13px; height: 13px; }
            QCheckBox::indicator:checked { border: 1px solid #000000; background-color: qradialgradient(cx: 0.5, cy: 0.5, radius: 0.5, fx: 0.5, fy: 0.5, stop: 0 #4287F5, stop: 0.4 #4287F5, stop: 0.41 #FFFFFF, stop: 1 #FFFFFF); }
        """)
    def get_selected_layers(self):
        return [cb.text() for cb in self.checkboxes if cb.isChecked()]

class MyGraphicsView(QGraphicsView):
    sceneClicked = pyqtSignal(QPointF); viewZoomed = pyqtSignal()
    def __init__(self, scene, parent=None):
        super().__init__(scene, parent)
        self.setDragMode(QGraphicsView.DragMode.NoDrag)
        self.setRenderHints(QPainter.RenderHint.Antialiasing | QPainter.RenderHint.TextAntialiasing | QPainter.RenderHint.SmoothPixmapTransform)
        self.viewport().setCursor(Qt.CursorShape.CrossCursor)
        self.is_panning = False
        self.last_pan_point = QPoint()
        self.main_window = parent

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
                self.is_panning = True
                self.last_pan_point = event.pos()
                self.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
                if self.main_window and self.main_window.in_area_cells_outline:
                    if self.main_window.in_area_cells_outline.scene():
                        self.main_window.scene.removeItem(self.main_window.in_area_cells_outline)
                    self.main_window.in_area_cells_outline = None
            else:
                self.sceneClicked.emit(self.mapToScene(event.pos()))
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.is_panning:
            delta = event.pos() - self.last_pan_point
            if self.main_window:
                self.main_window.map_offset_x += delta.x()
                self.main_window.map_offset_y += delta.y()
                self.main_window.redraw_all_layers(update_outline=False)
            self.last_pan_point = event.pos()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.is_panning:
            self.is_panning = False
            self.viewport().setCursor(Qt.CursorShape.CrossCursor)
            if self.main_window:
                self.main_window.update_area_outline()
        super().mouseReleaseEvent(event)
        
    def wheelEvent(self, event):
        zoom_factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.scale(zoom_factor, zoom_factor); self.viewZoomed.emit()

class X_Grid(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("X_Grid - 平均集材距離計算システム")
        self.setGeometry(50, 50, 1800, 1000)
        self.cell_size_on_screen = 25
        self.k_value = 25.0
        self.grid_rows_a4, self.grid_cols_a4 = 45, 30
        self.grid_rows_a3, self.grid_cols_a3 = 45, 73 
        self.grid_rows, self.grid_cols = self.grid_rows_a4, self.grid_cols_a4
        self.page_orientation = QPageLayout.Orientation.Portrait
        self.grid_offset_x, self.grid_offset_y = 60, 40
        self.layers = []
        self.master_bbox = None
        self.landing_cell = None
        self.map_rotation = 0
        self.grid_items = []
        self.compass_items = []
        self.calculation_items = []
        self.result_text_items = []
        self.title_items = []
        self.pointer_item = None
        self.in_area_cells_outline = None
        self.last_info_message = ""
        self.has_first_polygon = False
        self.map_offset_x = 0.0
        self.map_offset_y = 0.0
        self.export_file_path = ""
        self.Z_GRID = 0 
        self.Z_DATA_LAYERS_BASE = 1
        self.Z_AREA_OUTLINE = 50
        self.Z_OVERLAYS_BASE = 100 
        self._setup_drawing_styles()
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        top_level_layout = QHBoxLayout(central_widget)
        left_panel_widget = QWidget()
        left_panel_widget.setFixedWidth(380)
        left_panel_layout = QVBoxLayout(left_panel_widget)
        left_panel_layout.setSpacing(15)
        left_panel_layout.setContentsMargins(15, 15, 15, 15)
        
        usage_label = QLabel("<b>【基本操作】</b>")
        usage_text = ("""
        <style>
            ol { padding-left: 1.6em; margin: 0; }
            li { line-height: 1.6; padding-bottom: 6px; }
            b { color: #000000; }
        </style>
        <ol>
            <li><b>「レイヤ追加」</b> でファイル(.shp / .gpkg)選択</li>
            <li><b>  "土場か区域の入口"</b> を地図上でクリック</li>
            <li><b>「計算を実行」</b> ボタンで計算結果を表示</li>
            <li><b>  "林小班名等"</b>を入力し、<b>「表示」</b>を押す</li>
       	    <li><b>「エクスポート」</b> でPDFとして保存</li>
        </ol>
        <p style="margin-left: 1.6em; margin-top: 8px; font-size: 9pt; color: #333;">
            <b>ヒント:</b> Ctrl+ドラッグで地図を微調整できます
        </p>
        """)
        usage_desc = QLabel(usage_text)
        usage_desc.setWordWrap(True)
        left_panel_layout.addWidget(usage_label)
        left_panel_layout.addWidget(usage_desc)
        
        separator1 = QFrame()
        separator1.setFrameShape(QFrame.Shape.HLine)
        separator1.setFrameShadow(QFrame.Shadow.Sunken)
        left_panel_layout.addWidget(separator1)
        
        layer_management_label = QLabel("<b>レイヤ管理 (リストの上が最前面)</b>")
        self.layer_list_widget = QListWidget()
        layer_buttons_layout = QHBoxLayout()
        self.add_layer_button = QPushButton("レイヤ追加")
        self.remove_layer_button = QPushButton("削除")
        self.layer_up_button = QPushButton("↑")
        self.layer_down_button = QPushButton("↓")
        layer_buttons_layout.addWidget(self.add_layer_button)
        layer_buttons_layout.addWidget(self.remove_layer_button)
        layer_buttons_layout.addStretch(1)
        layer_buttons_layout.addWidget(self.layer_up_button)
        layer_buttons_layout.addWidget(self.layer_down_button)
        left_panel_layout.addWidget(layer_management_label)
        left_panel_layout.addWidget(self.layer_list_widget)
        left_panel_layout.addLayout(layer_buttons_layout)
        
        left_panel_layout.addStretch(1)
        
        self.notes_toggle_button = QPushButton("▼ 備考 (クリックで詳細表示)")
        self.notes_toggle_button.setCheckable(True)
        self.notes_toggle_button.setChecked(False)
        self.notes_toggle_button.setStyleSheet("background-color: transparent; border: none; text-align: left; padding: 4px; font-weight: bold; color: #555;")
        
        self.notes_widget = QWidget()
        notes_layout = QVBoxLayout(self.notes_widget)
        notes_layout.setContentsMargins(10, 5, 0, 0)
        
        notes_text = ("""
        <style>
            p, ul, ol { margin: 0; padding: 0; font-size: 9pt;}
            ul { padding-left: 1.6em; margin-bottom: 7px; }
            li { line-height: 1.5; padding-bottom: 4px; }
        </style>
        <p style="font-weight: bold; padding-bottom: 3px;">【重要】データに関する注意</p>
        <ul>
            <li>ベクターファイルは<b>平面直角座標系</b>を使用。</li>
            <li style="color: #000000; padding-top: 4px;">
              <b>属性データについて:</b><br>
             <b>フィールド名（列名）が文字化けしていると、読み込みに失敗します。その場合は "レイヤの文字コードを変更" or "列名を半角英数字に変更" or "文字化けしている属性の削除" 等で対処。</b><br>
            </li>
            <li style="padding-top: 4px;">
              <b>ラインの延長表示:</b><br>
              ラインレイヤの属性に「<b>meter</b>」というフィールドがあると、その値が地図上のラインの横に自動で表示（例: 123m）。
            </li>
        </ul>
        """)
        notes_label = QLabel(notes_text)
        notes_label.setWordWrap(True)
        notes_layout.addWidget(notes_label)
        self.notes_widget.setVisible(False)
        
        left_panel_layout.addWidget(self.notes_toggle_button)
        left_panel_layout.addWidget(self.notes_widget)
        
        right_panel_widget = QWidget()
        right_panel_layout = QVBoxLayout(right_panel_widget)
        right_panel_layout.setContentsMargins(0, 10, 10, 10)
        
        control_panel_layout = QHBoxLayout()
        self.calculate_button = QPushButton("計算を実行")
        self.update_title_button = QPushButton("表示")
        self.export_button = QPushButton("エクスポート")
        
        self.subtitle_input = QLineEdit()
        self.subtitle_input.setPlaceholderText("例：〇〇〇林小班、〇〇伐区")
        self.subtitle_input.setFixedWidth(250)
        
        control_panel_layout.addWidget(self.calculate_button)
        control_panel_layout.addSpacing(20)
        control_panel_layout.addWidget(self.subtitle_input)
        control_panel_layout.addWidget(self.update_title_button)
        control_panel_layout.addStretch(1)
        control_panel_layout.addWidget(self.export_button)
        
        right_panel_layout.addLayout(control_panel_layout)
        
        self.scene = QGraphicsScene(self)
        self.view = MyGraphicsView(self.scene, self)
        self.scene.setBackgroundBrush(QColor("#FFFFFF"))
        right_panel_layout.addWidget(self.view)
        top_level_layout.addWidget(left_panel_widget)
        top_level_layout.addWidget(right_panel_widget, 1)
        
        self.setStyleSheet("""
            QWidget { background-color: #F0F0F0; color: #000000; }
            QLabel { color: #000000; background-color: transparent; }
            QPushButton { background-color: #E1E1E1; border: 1px solid #ADADAD; padding: 5px 12px; border-radius: 3px; }
            QPushButton:hover { background-color: #E9E9E9; }
            QPushButton:pressed { background-color: #D6D6D6; }
            QListWidget { background-color: #FFFFFF; border: 1px solid #ABADB3; }
            QListWidget::item:selected { background-color: #D9E8FB; color: #000000; }
            QListWidget::indicator { border: 1px solid #000000; background-color: #FFFFFF; width: 13px; height: 13px; }
            QListWidget::indicator:checked { border: 1px solid #000000; background-color: qradialgradient(cx: 0.5, cy: 0.5, radius: 0.5, fx: 0.5, fy: 0.5, stop: 0 #4287F5, stop: 0.4 #4287F5, stop: 0.41 #FFFFFF, stop: 1 #FFFFFF); }
            QGraphicsView { border: 1px solid #767676; }
            QLineEdit { 
                background-color: #FFFFFF; 
                border: 1px solid #ABADB3; 
                border-radius: 3px; 
                padding: 2px 4px;
                selection-background-color: #CCCCCC;
                selection-color: #000000;
            }
            QFrame[frameShape="4"] { /* HLine */ border: none; height: 1px; background-color: #D1D1D1; }
        """)

        def toggle_notes():
            is_checked = self.notes_toggle_button.isChecked()
            self.notes_widget.setVisible(is_checked)
            self.notes_toggle_button.setText("▲ 備考 (クリックで非表示)" if is_checked else "▼ 備考 (クリックで詳細表示)")
        
        self.notes_toggle_button.clicked.connect(toggle_notes)
        self.add_layer_button.clicked.connect(self.prompt_add_layer)
        self.remove_layer_button.clicked.connect(self.remove_selected_layer)
        self.layer_up_button.clicked.connect(self.move_layer_up)
        self.layer_down_button.clicked.connect(self.move_layer_down)
        self.layer_list_widget.itemChanged.connect(self.on_layer_item_changed)
        self.view.sceneClicked.connect(self.on_scene_clicked)
        self.calculate_button.clicked.connect(self.run_calculation_and_draw)
        self.export_button.clicked.connect(self.export_results)
        self.update_title_button.clicked.connect(self.update_title_display)
        self.subtitle_input.returnPressed.connect(self.update_title_display)
        self.draw_grid()

    def prompt_add_layer(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "ベクターファイルを選択", "", "ベクターファイル (*.gpkg *.shp)")
        if not file_path: return
        
        layer_names_to_add = []
        try:
            if file_path.lower().endswith('.shp'):
                layer_names_to_add = [None] 
            elif file_path.lower().endswith('.gpkg'):
                all_layer_names = fiona.listlayers(file_path)
                dialog = LayerSelectionDialog(all_layer_names, self)
                if dialog.exec(): 
                    layer_names_to_add = dialog.get_selected_layers()
                else: 
                    return
            else:
                layer_names_to_add = fiona.listlayers(file_path)

        except Exception as e: 
            QMessageBox.critical(self, "エラー", f"ファイルからレイヤリストを取得できませんでした。\n\n詳細: {e}")
            return

        if not layer_names_to_add: 
            QMessageBox.warning(self, "警告", "ファイルから読み込み可能なレイヤが見つかりませんでした。")
            return

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            if self.add_layers_from_file(file_path, layer_names_to_add):
                self.map_offset_x = 0.0
                self.map_offset_y = 0.0
                self.update_layout_and_redraw()
        except Exception as e: 
            QMessageBox.critical(self, "エラー", f"レイヤ追加処理中にエラー: {e}")
        finally: 
            QApplication.restoreOverrideCursor()

    def add_layers_from_file(self, file_path, layer_names):
        new_layers_added = False
        self.layer_list_widget.blockSignals(True)
        
        for layer_name in layer_names:
            try:
                features, geom_type, layer_bbox = [], 'Unknown', None
                try:
                    with fiona.open(file_path, 'r', layer=layer_name, encoding='utf-8') as collection:
                        features = list(collection)
                        geom_type = collection.schema.get('geometry', 'Unknown')
                        layer_bbox = collection.bounds
                except (FionaError, UnicodeDecodeError):
                    with fiona.open(file_path, 'r', layer=layer_name, encoding='cp932') as collection:
                        features = list(collection)
                        geom_type = collection.schema.get('geometry', 'Unknown')
                        layer_bbox = collection.bounds
                if not features: continue
                is_calculable = "Polygon" in geom_type
                total_area = 0
                if is_calculable:
                    polygons = [shape(f['geometry']) for f in features if f.get('geometry')]
                    valid_polygons = [p for p in polygons if p.is_valid]
                    if valid_polygons: total_area = unary_union(valid_polygons).area
                if layer_name is None:
                    internal_name = os.path.splitext(os.path.basename(file_path))[0]
                    item_text = os.path.basename(file_path)
                else:
                    internal_name, item_text = layer_name, f"{os.path.basename(file_path)} ({layer_name})"
                layer_info = {'path': file_path, 'layer_name': internal_name, 'geom_type': geom_type, 'features': features, 'graphics_items': [], 'is_calculable': is_calculable, 'is_calc_target': is_calculable, 'bbox': layer_bbox, 'area': total_area}
                list_item = QListWidgetItem(item_text)
                if is_calculable:
                    list_item.setFlags(list_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                    list_item.setCheckState(Qt.CheckState.Checked)
                else:
                    list_item.setFlags(list_item.flags() & ~Qt.ItemFlag.ItemIsUserCheckable)
                self.layers.insert(0, layer_info)
                self.layer_list_widget.insertItem(0, list_item)
                new_layers_added = True
            except Exception as e: 
                print(f"警告: レイヤ '{layer_name}' の読み込みをスキップ。理由: {e}")
                QMessageBox.warning(self, "読み込みエラー", f"ファイルの読み込みに失敗しました。\nファイル形式またはエンコーディングがサポートされていない可能性があります。\n\n詳細: {e}")
                continue
        self.layer_list_widget.blockSignals(False)
        return new_layers_added
    
    def remove_selected_layer(self):
        current_row = self.layer_list_widget.currentRow()
        if current_row < 0: return
        self.layers.pop(current_row)
        self.layer_list_widget.takeItem(current_row)
        self.update_layout_and_redraw()

    def move_layer_up(self):
        current_row = self.layer_list_widget.currentRow()
        if current_row > 0:
            self.layers.insert(current_row - 1, self.layers.pop(current_row))
            item = self.layer_list_widget.takeItem(current_row)
            self.layer_list_widget.insertItem(current_row - 1, item)
            self.layer_list_widget.setCurrentRow(current_row - 1)
            self.redraw_all_layers()

    def move_layer_down(self):
        current_row = self.layer_list_widget.currentRow()
        if 0 <= current_row < self.layer_list_widget.count() - 1:
            self.layers.insert(current_row + 1, self.layers.pop(current_row))
            item = self.layer_list_widget.takeItem(current_row)
            self.layer_list_widget.insertItem(current_row + 1, item)
            self.layer_list_widget.setCurrentRow(current_row + 1)
            self.redraw_all_layers()

    def on_layer_item_changed(self, item):
        row = self.layer_list_widget.row(item)
        if 0 <= row < len(self.layers):
            is_checked = (item.checkState() == Qt.CheckState.Checked)
            self.layers[row]['is_calc_target'] = is_checked
            self.clear_calculation_results()
            self.update_area_outline()

    def update_layout_and_redraw(self):
        self.update_master_bbox()
        self.determine_layout()
        self.redraw_all_layers()
        self.auto_fit_view()

    def _get_combined_calculable_geom(self):
        all_shapely_polygons = []
        calculable_layers = [layer for layer in self.layers if layer.get('is_calc_target') and layer.get('is_calculable')]
        if not calculable_layers: return None
        for layer in calculable_layers:
            for feature in layer['features']:
                geom_dict = feature.get('geometry')
                if not geom_dict: continue
                try:
                    shapely_geom = shape(geom_dict)
                    if not shapely_geom.is_valid: shapely_geom = shapely_geom.buffer(0)
                    if shapely_geom.is_empty: continue
                    all_shapely_polygons.append(shapely_geom)
                except Exception: continue
        if not all_shapely_polygons: return None
        return unary_union(all_shapely_polygons)

    def _get_combined_all_layers_geom(self):
        all_geoms = []
        for layer in self.layers:
            for feature in layer['features']:
                geom_dict = feature.get('geometry')
                if not geom_dict: continue
                try:
                    shapely_geom = shape(geom_dict)
                    if not shapely_geom.is_valid: shapely_geom = shapely_geom.buffer(0)
                    if shapely_geom.is_empty: continue
                    all_geoms.append(shapely_geom)
                except Exception: continue
        if not all_geoms: return None
        return unary_union(all_geoms)

    def _find_optimal_rotation(self, geom, target_width, target_height):
        if geom is None or geom.is_empty: return None
        for angle in range(1, 90):
            rotated_geom = rotate(geom, angle, origin='center', use_radians=False)
            min_x, min_y, max_x, max_y = rotated_geom.bounds
            width, height = max_x - min_x, max_y - min_y
            if width <= target_width and height <= target_height: return angle
        return None

    def _apply_rotation_to_coords(self, coords):
        if self.map_rotation == 0 or not self.master_bbox: return coords
        orig_center_x = self.master_bbox[0] + (self.master_bbox[2] - self.master_bbox[0]) / 2
        orig_center_y = self.master_bbox[1] + (self.master_bbox[3] - self.master_bbox[1]) / 2
        theta = math.radians(self.map_rotation)
        cos_theta, sin_theta = math.cos(theta), math.sin(theta)
        rotated_ps = []
        for p_x, p_y in coords:
            tx, ty = p_x - orig_center_x, p_y - orig_center_y
            rotated_x = tx * cos_theta - ty * sin_theta + orig_center_x
            rotated_y = tx * sin_theta + ty * cos_theta + orig_center_y
            rotated_ps.append((rotated_x, rotated_y))
        return rotated_ps

    def _check_fit(self, bbox, grid_rows, grid_cols):
        if not bbox: return False
        width, height = bbox[2] - bbox[0], bbox[3] - bbox[1]
        return width <= grid_cols * self.k_value and height <= grid_rows * self.k_value

    def determine_layout(self):
        if not self.master_bbox:
            self.grid_rows, self.grid_cols, self.page_orientation, self.map_rotation = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, 0
            return
        master_geom = self._get_combined_all_layers_geom()
        info_message, layout_found = "", False
        final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation = self.grid_rows, self.grid_cols, self.page_orientation, self.map_rotation
        if master_geom and not master_geom.is_empty:
            bbox, rotated_90_geom = master_geom.bounds, rotate(master_geom, 90, origin='center', use_radians=False)
            rotated_90_bbox, a4_width_m, a4_height_m, a3_width_m, a3_height_m = rotated_90_geom.bounds, self.grid_cols_a4 * self.k_value, self.grid_rows_a4 * self.k_value, self.grid_cols_a3 * self.k_value, self.grid_rows_a3 * self.k_value
            if self._check_fit(bbox, self.grid_rows_a4, self.grid_cols_a4): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, 0, "", True
            if not layout_found and self._check_fit(rotated_90_bbox, self.grid_rows_a4, self.grid_cols_a4): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, 90, "A4縦に収めるため、90°回転しました。", True
            if not layout_found:
                optimal_angle_a4 = self._find_optimal_rotation(master_geom, a4_width_m, a4_height_m)
                if optimal_angle_a4 is not None: final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, optimal_angle_a4, f"A4縦に収めるため、{optimal_angle_a4}°回転しました。", True
            if not layout_found and self._check_fit(bbox, self.grid_rows_a3, self.grid_cols_a3): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 0, "データ範囲が大きいため、A3横モードに切り替えました。", True
            if not layout_found and self._check_fit(rotated_90_bbox, self.grid_rows_a3, self.grid_cols_a3): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 90, "A3横に収めるため、90°回転しました。", True
            if not layout_found:
                optimal_angle_a3 = self._find_optimal_rotation(master_geom, a3_width_m, a3_height_m)
                if optimal_angle_a3 is not None: final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message, layout_found = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, optimal_angle_a3, f"A3横に収めるため、{optimal_angle_a3}°回転しました。", True
            if not layout_found: final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 0, "A3モードでも最適な回転が見つかりませんでした。データの一部が切れて表示される可能性があります。"
        else:
            _, rotated_bbox = self._rotate_points_90_degrees_bbox(self.master_bbox)
            if self._check_fit(self.master_bbox, self.grid_rows_a4, self.grid_cols_a4): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, 0
            elif self._check_fit(rotated_bbox, self.grid_rows_a4, self.grid_cols_a4): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message = self.grid_rows_a4, self.grid_cols_a4, QPageLayout.Orientation.Portrait, 90, "A4縦に収めるため、90°回転しました。"
            elif self._check_fit(self.master_bbox, self.grid_rows_a3, self.grid_cols_a3): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 0, "データ範囲が大きいため、A3横モードに切り替えました。"
            elif self._check_fit(rotated_bbox, self.grid_rows_a3, self.grid_cols_a3): final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 90, "A3横に収めるため、90°回転しました。"
            else: final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation, info_message = self.grid_rows_a3, self.grid_cols_a3, QPageLayout.Orientation.Landscape, 0, "A3モードでもグリッド範囲に収まりません。データの一部が切れて表示される可能性があります。"
        layout_changed = (self.grid_rows != final_grid_rows or self.grid_cols != final_grid_cols or self.map_rotation != final_map_rotation or self.page_orientation != final_page_orientation)
        self.grid_rows, self.grid_cols, self.page_orientation, self.map_rotation = final_grid_rows, final_grid_cols, final_page_orientation, final_map_rotation
        if layout_changed and info_message and info_message != self.last_info_message:
            if "一部が切れて" in info_message or "見つかりませんでした" in info_message: QMessageBox.warning(self, "警告", info_message)
            else: QMessageBox.information(self, "情報", info_message)
            self.last_info_message = info_message
        elif not info_message: self.last_info_message = ""

    def _rotate_points_90_degrees_bbox(self, bbox):
        min_x, min_y, max_x, max_y = bbox
        center_x, center_y = min_x + (max_x - min_x) / 2, min_y + (max_y - min_y) / 2
        points = [(min_x, min_y), (max_x, min_y), (max_x, max_y), (min_x, max_y)]
        rotated_points = [(center_x - (y - center_y), center_y + (x - center_x)) for x, y in points]
        rotated_xs, rotated_ys = [p[0] for p in rotated_points], [p[1] for p in rotated_points]
        return rotated_points, (min(rotated_xs), min(rotated_ys), max(rotated_xs), max(rotated_ys))

    def update_master_bbox(self):
        self.master_bbox = None
        for layer in self.layers:
            if not layer.get('bbox'): continue
            minx, miny, maxx, maxy = layer['bbox']
            if self.master_bbox is None: self.master_bbox = [minx, miny, maxx, maxy]
            else:
                self.master_bbox[0] = min(self.master_bbox[0], minx)
                self.master_bbox[1] = min(self.master_bbox[1], miny)
                self.master_bbox[2] = max(self.master_bbox[2], maxx)
                self.master_bbox[3] = max(self.master_bbox[3], maxy)

    def redraw_all_layers(self, update_outline=True):
        if self.in_area_cells_outline and self.in_area_cells_outline.scene(): self.scene.removeItem(self.in_area_cells_outline)
        self.scene.clear()
        self.grid_items.clear(); self.compass_items.clear(); self.calculation_items.clear(); self.result_text_items.clear(); self.title_items.clear()
        self.pointer_item = None; self.in_area_cells_outline = None
        for layer in self.layers: layer['graphics_items'].clear()
        self.draw_grid()
        if not self.master_bbox: return
        rotated_corners = self._apply_rotation_to_coords([(self.master_bbox[0], self.master_bbox[1]), (self.master_bbox[2], self.master_bbox[1]), (self.master_bbox[2], self.master_bbox[3]), (self.master_bbox[0], self.master_bbox[3])])
        xs, ys = [p[0] for p in rotated_corners], [p[1] for p in rotated_corners]
        bbox_to_use, params = (min(xs), min(ys), max(xs), max(ys)), self._get_transform_parameters_from_bbox((min(xs), min(ys), max(xs), max(ys)))
        if not params: return
        def draw_line_label(q_points, label_text, z_value, layer_dict):
            if len(q_points) < 2: return
            mid_index, p1, p2 = len(q_points) // 2, q_points[len(q_points) // 2 - 1], q_points[len(q_points) // 2]
            mid_point, angle_rad = QPointF((p1.x() + p2.x()) / 2, (p1.y() + p2.y()) / 2), math.atan2(p2.y() - p1.y(), p2.x() - p1.x())
            angle_deg, offset_angle_rad, offset_distance = math.degrees(angle_rad), angle_rad - math.pi / 2, 8
            offset_x, offset_y = offset_distance * math.cos(offset_angle_rad), offset_distance * math.sin(offset_angle_rad)
            label_pos = QPointF(mid_point.x() - offset_x, mid_point.y() - offset_y) if angle_deg > 90 or angle_deg < -90 else QPointF(mid_point.x() + offset_x, mid_point.y() + offset_y)
            if angle_deg > 90 or angle_deg < -90: angle_deg += 180
            font, text_item = QFont("游ゴシック", 8, QFont.Weight.Bold), self.scene.addText(label_text, QFont("游ゴシック", 8, QFont.Weight.Bold))
            text_item.setDefaultTextColor(QColor("black")); text_item.setZValue(z_value + 0.5); text_rect = text_item.boundingRect()
            text_item.setPos(label_pos.x() - text_rect.width() / 2, label_pos.y() - text_rect.height() / 2)
            text_item.setTransformOriginPoint(text_rect.center()); text_item.setRotation(angle_deg)
            layer_dict['graphics_items'].append(text_item)
        for i, layer in enumerate(self.layers):
            z_value = self.Z_DATA_LAYERS_BASE + (len(self.layers) - 1 - i)
            for feature in layer['features']:
                try:
                    geom = feature.get('geometry')
                    if not geom or not geom.get('coordinates'): continue
                    style = self._get_feature_style(feature, layer)
                    pen_width_in_scene_units, unit, width_val = 0.0, style.get('line_width_unit', 'MM').upper(), style.get('line_width', 0)
                    if unit == 'MM': pen_width_in_scene_units = (width_val * 5.0) * params['scale']
                    elif unit in ('PIXEL', 'PX'):
                        current_view_scale = self.view.transform().m11()
                        if current_view_scale > 0: pen_width_in_scene_units = width_val / current_view_scale
                    pen = QPen(style['line_color'], pen_width_in_scene_units)
                    pen.setStyle(style['pen_style'])
                    if style['pen_style'] == Qt.PenStyle.CustomDashLine and style['dash_pattern']: pen.setDashPattern(style['dash_pattern'])
                    pen.setCosmetic(False)
                    brush, path = QBrush(style['fill_color']), QPainterPath()
                    brush.setStyle(Qt.BrushStyle.SolidPattern if style['fill_color'].alpha() != 0 else Qt.BrushStyle.NoBrush)
                    def transform_and_rotate_coords(coords):
                        rotated_coords = self._apply_rotation_to_coords(coords)
                        return [QPointF(params['grid_center_x'] + (p[0] - params['center_x']) * params['scale'] + self.map_offset_x, params['grid_center_y'] - (p[1] - params['center_y']) * params['scale'] + self.map_offset_y) for p in rotated_coords]
                    items_created = []
                    if layer['geom_type'] in ('Polygon', 'MultiPolygon'):
                        path.setFillRule(Qt.FillRule.OddEvenFill)
                        coords_list = geom['coordinates'] if geom['type'] == 'MultiPolygon' else [geom['coordinates']]
                        for poly_rings in coords_list:
                            for ring in poly_rings:
                                if ring and len(ring) >= 3: path.addPolygon(QPolygonF(transform_and_rotate_coords(ring)))
                        items_created.append(self.scene.addPath(path, pen, brush))
                    elif layer['geom_type'] in ('LineString', 'MultiLineString'):
                        coords_list = geom['coordinates'] if geom['type'] == 'MultiLineString' else [geom['coordinates']]
                        for line_coords in coords_list:
                            if len(line_coords) < 2: continue
                            q_points = transform_and_rotate_coords(line_coords); line_path = QPainterPath(); line_path.moveTo(q_points[0])
                            for p in q_points[1:]: line_path.lineTo(p)
                            items_created.append(self.scene.addPath(line_path, pen))
                            properties = feature.get('properties', {})
                            if 'meter' in properties and properties['meter'] is not None:
                                try: label_text = f"{int(float(properties['meter']))}m"
                                except (ValueError, TypeError): label_text = f"{properties['meter']}m"
                                if label_text.strip() != "m": draw_line_label(q_points, label_text, z_value, layer)
                    else: continue
                    for item in items_created:
                        if item: item.setZValue(z_value); setattr(item, 'style_info', style); layer['graphics_items'].append(item)
                except Exception as e: print(f"警告: フィーチャ描画をスキップ。理由: {e}"); continue
        self.draw_compass()
        if update_outline: self.update_area_outline()

    def _get_feature_style(self, feature, layer_info):
        props = feature.get('properties', {})
        final_style = DEFAULT_STYLE_INFO.copy()
        fill_color_prop = props.get('fill_color')
        if fill_color_prop is not None:
            prop_val_str = str(fill_color_prop).strip()
            if not prop_val_str: final_style['fill_color'] = QColor(Qt.GlobalColor.transparent)
            else:
                new_color = _parse_any_color_string(prop_val_str)
                if new_color.isValid(): final_style['fill_color'] = new_color
        line_color_prop = props.get('strk_color') or props.get('stroke_color') or props.get('color')
        if line_color_prop is not None:
            new_line_color = _parse_any_color_string(line_color_prop)
            if new_line_color.isValid(): final_style['line_color'] = new_line_color
        line_width_prop = props.get('strk_width') or props.get('stroke_width')
        if line_width_prop is not None:
            try: final_style['line_width'] = float(line_width_prop)
            except (ValueError, TypeError): pass
        style_key = props.get('strk_style') or props.get('stroke_dash_type') or props.get('stroke_style')
        if style_key is not None:
            style_val, pen_style_map = str(style_key).lower(), {'solid': Qt.PenStyle.SolidLine, 'dot': Qt.PenStyle.DotLine, 'dash': Qt.PenStyle.DashLine, 'dashdot': Qt.PenStyle.DashDotLine, 'dashdotdot': Qt.PenStyle.DashDotDotLine, 'custom': Qt.PenStyle.CustomDashLine, 'none': Qt.PenStyle.NoPen, 'no': Qt.PenStyle.NoPen}
            final_style['pen_style'] = pen_style_map.get(style_val, Qt.PenStyle.SolidLine)
        pattern_prop = props.get('dash_pattn') or props.get('dash_pattern')
        if final_style.get('pen_style') == Qt.PenStyle.CustomDashLine and pattern_prop:
            final_style['dash_pattern'] = []
            try:
                pattern_str = str(pattern_prop).strip().replace('[','').replace(']','')
                scale_factor = final_style.get('line_width', 1.0)
                pattern_list = [float(p.strip()) * scale_factor for p in pattern_str.split(',')]
                if pattern_list: final_style['dash_pattern'] = pattern_list
            except (ValueError, TypeError, AttributeError):
                final_style['dash_pattern'], final_style['pen_style'] = [], Qt.PenStyle.SolidLine
        current_fill_color = final_style['fill_color']
        if current_fill_color.alpha() != 0:
            current_fill_color.setAlpha(100)
            final_style['fill_color'] = current_fill_color
        return final_style

    def auto_fit_view(self):
        all_items_rect = self.scene.itemsBoundingRect()
        grid_rect = QRectF(self.grid_offset_x, self.grid_offset_y, self.grid_cols * self.cell_size_on_screen, self.grid_rows * self.cell_size_on_screen)
        bounding_rect = all_items_rect.united(grid_rect)
        if bounding_rect.isValid(): self.view.fitInView(bounding_rect.adjusted(-20, -20, 20, 20), Qt.AspectRatioMode.KeepAspectRatio)

    def _get_transform_parameters_from_bbox(self, bbox):
        if not bbox: return None
        scale, min_x, min_y, max_x, max_y = self.cell_size_on_screen / self.k_value, bbox[0], bbox[1], bbox[2], bbox[3]
        center_x, center_y = min_x + (max_x - min_x) / 2, min_y + (max_y - min_y) / 2
        grid_center_x, grid_center_y = self.grid_offset_x + (self.grid_cols * self.cell_size_on_screen) / 2, self.grid_offset_y + (self.grid_rows * self.cell_size_on_screen) / 2
        return {'scale': scale, 'center_x': center_x, 'center_y': center_y, 'grid_center_x': grid_center_x, 'grid_center_y': grid_center_y}

    def on_scene_clicked(self, scene_pos):
        grid_rect = QRectF(self.grid_offset_x, self.grid_offset_y, self.grid_cols * self.cell_size_on_screen, self.grid_rows * self.cell_size_on_screen)
        if not any(layer.get('is_calculable') for layer in self.layers):
            QMessageBox.information(self, "情報", "先にポリゴンレイヤを読み込んでください。"); return
        if grid_rect.contains(scene_pos):
            col, row = int((scene_pos.x() - self.grid_offset_x) / self.cell_size_on_screen), int((scene_pos.y() - self.grid_offset_y) / self.cell_size_on_screen)
            if 0 <= row < self.grid_rows and 0 <= col < self.grid_cols:
                self.landing_cell = (row, col)
                if self.pointer_item and self.pointer_item.scene(): self.scene.removeItem(self.pointer_item)
                center_x, center_y, point_size = self.grid_offset_x + col * self.cell_size_on_screen + self.cell_size_on_screen / 2, self.grid_offset_y + row * self.cell_size_on_screen + self.cell_size_on_screen / 2, self.cell_size_on_screen * 0.5
                self.pointer_item = self.scene.addEllipse(center_x - point_size / 2, center_y - point_size / 2, point_size, point_size, QPen(QColor("red"), 1), QBrush(QColor("red")))
                self.pointer_item.setZValue(self.Z_OVERLAYS_BASE + 2)
                self.clear_calculation_results()

    def draw_grid(self):
        for item in self.grid_items:
            if item.scene(): self.scene.removeItem(item)
        self.grid_items.clear()
        pen = QPen(QColor(220, 220, 222)); pen.setCosmetic(False)
        for r in range(self.grid_rows + 1):
            y = self.grid_offset_y + r * self.cell_size_on_screen
            line = self.scene.addLine(self.grid_offset_x, y, self.grid_offset_x + self.grid_cols * self.cell_size_on_screen, y, pen)
            line.setZValue(self.Z_GRID); self.grid_items.append(line)
        for c in range(self.grid_cols + 1):
            x = self.grid_offset_x + c * self.cell_size_on_screen
            line = self.scene.addLine(x, self.grid_offset_y, x, self.grid_offset_y + self.grid_rows * self.cell_size_on_screen, pen)
            line.setZValue(self.Z_GRID); self.grid_items.append(line)

    def draw_compass(self):
        for item in self.compass_items:
            if item.scene(): self.scene.removeItem(item)
        self.compass_items.clear()
        if not self.layers: return
        compass_margin, center_x, center_y, size = self.cell_size_on_screen * 2.0, self.grid_offset_x + self.cell_size_on_screen * 2.0, self.grid_offset_y + self.cell_size_on_screen * 2.0, 35
        compass_group = self.scene.createItemGroup([])
        ns_poly, ew_poly = QPolygonF([QPointF(0, -size/2), QPointF(size/10, 0), QPointF(0, size/2), QPointF(-size/10, 0)]), QPolygonF([QPointF(size/2, 0), QPointF(0, size/10), QPointF(-size/2, 0), QPointF(0, -size/10)])
        dark_brush, light_brush, no_pen = QBrush(QColor(50, 50, 50)), QBrush(QColor(150, 150, 150)), QPen(Qt.PenStyle.NoPen)
        ew_item, ns_item = self.scene.addPolygon(ew_poly, no_pen, light_brush), self.scene.addPolygon(ns_poly, no_pen, dark_brush)
        font, text_item = QFont("游ゴシック", 10, QFont.Weight.Bold), self.scene.addText("N", QFont("游ゴシック", 10, QFont.Weight.Bold))
        text_item.setDefaultTextColor(QColor("black")); text_rect = text_item.boundingRect()
        text_item.setPos(-text_rect.width() / 2, -size / 2 - text_rect.height() + 3)
        compass_group.addToGroup(ew_item); compass_group.addToGroup(ns_item); compass_group.addToGroup(text_item)
        compass_group.setPos(center_x, center_y); compass_group.setRotation(-self.map_rotation)
        compass_group.setZValue(self.Z_OVERLAYS_BASE + 1); self.compass_items.append(compass_group)

    def get_in_area_cells(self):
        if not self.master_bbox: return []
        rotated_corners = self._apply_rotation_to_coords([(self.master_bbox[0], self.master_bbox[1]), (self.master_bbox[2], self.master_bbox[1]), (self.master_bbox[2], self.master_bbox[3]), (self.master_bbox[0], self.master_bbox[3])])
        xs, ys = [p[0] for p in rotated_corners], [p[1] for p in rotated_corners]
        bbox_to_use, params = (min(xs), min(ys), max(xs), max(ys)), self._get_transform_parameters_from_bbox((min(xs), min(ys), max(xs), max(ys)))
        if not params: return []
        combined_world_geom = self._get_combined_calculable_geom()
        if combined_world_geom is None or combined_world_geom.is_empty: return []
        def transform_geom_coords(coords):
            rotated_coords = self._apply_rotation_to_coords(coords)
            return [(params['grid_center_x'] + (p[0] - params['center_x']) * params['scale'] + self.map_offset_x, params['grid_center_y'] - (p[1] - params['center_y']) * params['scale'] + self.map_offset_y) for p in rotated_coords]
        combined_scene_geom = None
        try:
            if combined_world_geom.geom_type == 'Polygon': combined_scene_geom = Polygon(transform_geom_coords(combined_world_geom.exterior.coords), [transform_geom_coords(interior.coords) for interior in combined_world_geom.interiors])
            elif combined_world_geom.geom_type == 'MultiPolygon': polys = [Polygon(transform_geom_coords(p.exterior.coords), [transform_geom_coords(i.coords) for i in p.interiors]) for p in combined_world_geom.geoms if p.exterior]; combined_scene_geom = MultiPolygon(polys)
        except Exception as e: print(f"シーンジオメトリ変換エラー: {e}"); return []
        if not combined_scene_geom or combined_scene_geom.is_empty: return []
        in_area_cells, area_threshold = [], 0.5 * (self.cell_size_on_screen ** 2)
        for r in range(self.grid_rows):
            for c in range(self.grid_cols):
                cell_poly = box(self.grid_offset_x + c * self.cell_size_on_screen, self.grid_offset_y + r * self.cell_size_on_screen, self.grid_offset_x + (c + 1) * self.cell_size_on_screen, self.grid_offset_y + (r + 1) * self.cell_size_on_screen)
                if combined_scene_geom.intersects(cell_poly) and combined_scene_geom.intersection(cell_poly).area >= area_threshold: in_area_cells.append((r, c))
        return in_area_cells

    def update_area_outline(self):
        if self.in_area_cells_outline and self.in_area_cells_outline.scene(): self.scene.removeItem(self.in_area_cells_outline)
        self.in_area_cells_outline = None
        in_area_cells = self.get_in_area_cells()
        if not in_area_cells: return
        cell_polygons = []
        for r, c in in_area_cells:
            min_x, min_y = self.grid_offset_x + c * self.cell_size_on_screen, self.grid_offset_y + r * self.cell_size_on_screen
            cell_polygons.append(box(min_x, min_y, min_x + self.cell_size_on_screen, min_y + self.cell_size_on_screen))
        if not cell_polygons: return
        merged_cells_geom, outline_path = unary_union(cell_polygons), QPainterPath()
        def add_geom_to_path(geom, path):
            if geom.is_empty: return
            if geom.geom_type == 'Polygon':
                exterior_coords = list(geom.exterior.coords)
                if len(exterior_coords) > 1:
                    path.moveTo(QPointF(exterior_coords[0][0], exterior_coords[0][1]))
                    for x, y in exterior_coords[1:]: path.lineTo(QPointF(x, y))
            elif geom.geom_type == 'MultiPolygon':
                for poly in geom.geoms: add_geom_to_path(poly, path)
            elif geom.geom_type in ('LineString', 'MultiLineString', 'LinearRing'):
                 coords_list = [list(geom.coords)] if geom.geom_type in ('LineString', 'LinearRing') else [list(g.coords) for g in geom.geoms]
                 for coords in coords_list:
                     if len(coords) > 1: path.moveTo(QPointF(coords[0][0], coords[0][1])); [path.lineTo(QPointF(x, y)) for x, y in coords[1:]]
        add_geom_to_path(merged_cells_geom.boundary, outline_path)
        outline_pen = QPen(QColor(0, 80, 200, 150), 3, Qt.PenStyle.DashDotLine)
        outline_pen.setCosmetic(True)
        self.in_area_cells_outline = self.scene.addPath(outline_path, outline_pen)
        self.in_area_cells_outline.setZValue(self.Z_AREA_OUTLINE)

    def clear_calculation_results(self):
        for item in self.calculation_items + self.result_text_items + self.title_items:
            if item.scene(): self.scene.removeItem(item)
        self.calculation_items.clear(); self.result_text_items.clear(); self.title_items.clear()

    def update_title_display(self):
        for item in self.title_items:
            if item.scene(): self.scene.removeItem(item)
        self.title_items.clear()
        y_pos, subtitle_text = self.grid_offset_y - 145, self.subtitle_input.text().strip()
        title_font, title_text = QFont("游ゴシック", 16, QFont.Weight.Bold), f"{subtitle_text}  平均集材距離計算表" if subtitle_text else "平均集材距離計算表"
        title_item = self._add_aligned_text(title_text, title_font, self.colors['dark'], QPointF(self.grid_offset_x, y_pos), Qt.AlignmentFlag.AlignLeft)
        self.title_items.append(title_item)
        if subtitle_text:
            metrics, subtitle_width, rect = QFontMetrics(title_font), QFontMetrics(title_font).horizontalAdvance(subtitle_text), title_item.boundingRect()
            underline_y, line1, line2 = y_pos + rect.height(), self.scene.addLine(self.grid_offset_x, y_pos + rect.height() + 1, self.grid_offset_x + QFontMetrics(title_font).horizontalAdvance(subtitle_text), y_pos + rect.height() + 1, QPen(self.colors['dark'])), self.scene.addLine(self.grid_offset_x, y_pos + rect.height() + 3, self.grid_offset_x + QFontMetrics(title_font).horizontalAdvance(subtitle_text), y_pos + rect.height() + 3, QPen(self.colors['dark']))
            self.title_items.extend([line1, line2])

    def run_calculation_and_draw(self):
        self.clear_calculation_results(); self.update_area_outline()
        in_area_cells = self.get_in_area_cells()
        if not in_area_cells: QMessageBox.warning(self, "警告", "計算対象の区域がありません。レイヤ管理リストでポリゴンレイヤにチェックを入れてください。"); return
        if not self.landing_cell: QMessageBox.warning(self, "警告", "土場の位置が選択されていません。"); return
        debug_pen, debug_brush = QPen(QColor("darkgray")), QBrush(QColor("darkgray"))
        for r, c in in_area_cells:
            center_x, center_y = self.grid_offset_x + c * self.cell_size_on_screen + self.cell_size_on_screen / 2, self.grid_offset_y + r * self.cell_size_on_screen + self.cell_size_on_screen / 2
            marker = self.scene.addRect(center_x - 1, center_y - 1, 2, 2, debug_pen, debug_brush)
            marker.setZValue(self.Z_AREA_OUTLINE + 1); self.calculation_items.append(marker)
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            landing_row, landing_col = self.landing_cell
            row_counts, col_counts = {r: 0 for r in range(self.grid_rows)}, {c: 0 for c in range(self.grid_cols)}
            for r, c in in_area_cells: row_counts[r] += 1; col_counts[c] += 1
            total_product_v, total_product_h = sum(abs(r - landing_row) * count for r, count in row_counts.items()), sum(abs(c - landing_col) * count for c, count in col_counts.items())
            total_degree, final_distance = len(in_area_cells), (total_product_v + total_product_h) / len(in_area_cells) * self.k_value if len(in_area_cells) > 0 else 0
            all_rows, all_cols = [r for r, c in in_area_cells] if in_area_cells else [], [c for r, c in in_area_cells] if in_area_cells else []
            calc_data = {"landing_row": landing_row, "landing_col": landing_col, "row_counts": row_counts, "col_counts": col_counts, "total_product_v": total_product_v, "total_product_h": total_product_h, "total_degree": total_degree, "final_distance": final_distance, "min_row": min(all_rows) if all_rows else 0, "max_row": max(all_rows) if all_rows else 0, "min_col": min(all_cols) if all_cols else 0, "max_col": max(all_cols) if all_cols else 0, "subtitle": self.subtitle_input.text().strip()}
            self._draw_calculation_header(calc_data); self._draw_final_result(calc_data); self._draw_calculation_tables(calc_data)
            self.auto_fit_view()
        except Exception as e: QMessageBox.critical(self, "エラー", f"計算または描画中にエラーが発生しました: {e}")
        finally: QApplication.restoreOverrideCursor()

    def _setup_drawing_styles(self):
        self.fonts = { 'title': QFont("游ゴシック", 16, QFont.Weight.Bold), 'legend': QFont("游ゴシック", 9), 'scale': QFont("游ゴシック", 10), 'cell_count': QFont("游ゴシック", 10, QFont.Weight.Bold), 'result': QFont("游ゴシック", 12, QFont.Weight.Bold), 'header': QFont("游ゴシック", 9, QFont.Weight.Bold), 'data': QFont("游ゴシック", 9), 'total': QFont("游ゴシック", 9, QFont.Weight.Bold), 'highlight': QFont("游ゴシック", 9, QFont.Weight.Bold) }
        self.colors = { 'normal': QColor("#333333"), 'dark': QColor("black"), 'highlight': QColor("red") }

    def _add_aligned_text(self, text, font, color, point, alignment=Qt.AlignmentFlag.AlignCenter, is_result=False):
        item = self.scene.addText(text, font); item.setDefaultTextColor(color); text_rect = item.boundingRect(); item_x, item_y = point.x(), point.y()
        if alignment & Qt.AlignmentFlag.AlignHCenter: item_x -= text_rect.width() / 2
        elif alignment & Qt.AlignmentFlag.AlignRight: item_x -= text_rect.width()
        if alignment & Qt.AlignmentFlag.AlignVCenter: item_y -= text_rect.height() / 2
        elif alignment & Qt.AlignmentFlag.AlignBottom: item_y -= text_rect.height()
        item.setPos(item_x, item_y)
        (self.result_text_items if is_result else self.calculation_items).append(item)
        return item

    def _draw_calculation_header(self, calc_data):
        self.update_title_display()
        legend_y, col_widths_v, v_table_width = self.grid_offset_y - 145 + 4, [40, 35, 45], sum([40, 35, 45])
        content_right_edge, k_part_offset_width, scale_text_width = self.grid_offset_x + self.grid_cols * self.cell_size_on_screen + 5 + v_table_width, self.cell_size_on_screen + 65, QFontMetrics(self.fonts['scale']).horizontalAdvance("縮尺: 1/5000")
        legend_block_width, legend_x = k_part_offset_width + scale_text_width, content_right_edge - (k_part_offset_width + scale_text_width) - 10
        legend_pen = QPen(self.colors['dark'], 1.0); legend_pen.setCosmetic(False)
        self.calculation_items.append(self.scene.addRect(legend_x, legend_y, self.cell_size_on_screen, self.cell_size_on_screen, legend_pen))
        dim_pen, tick_size, h_dim_y = QPen(self.colors['dark'], 0.8), 2, legend_y + self.cell_size_on_screen + 5
        dim_pen.setCosmetic(False)
        self.calculation_items.extend([self.scene.addLine(legend_x, h_dim_y, legend_x + self.cell_size_on_screen, h_dim_y, dim_pen), self.scene.addLine(legend_x, h_dim_y - tick_size, legend_x, h_dim_y + tick_size, dim_pen), self.scene.addLine(legend_x + self.cell_size_on_screen, h_dim_y - tick_size, legend_x + self.cell_size_on_screen, h_dim_y + tick_size, dim_pen)])
        v_dim_x = legend_x + self.cell_size_on_screen + 5
        self.calculation_items.extend([self.scene.addLine(v_dim_x, legend_y, v_dim_x, legend_y + self.cell_size_on_screen, dim_pen), self.scene.addLine(v_dim_x - tick_size, legend_y, v_dim_x + tick_size, legend_y, dim_pen), self.scene.addLine(v_dim_x - tick_size, legend_y + self.cell_size_on_screen, v_dim_x + tick_size, legend_y + self.cell_size_on_screen, dim_pen)])
        k_value_text = f"K ({self.k_value:.0f}m)"
        self._add_aligned_text(k_value_text, self.fonts['legend'], self.colors['dark'], QPointF(legend_x + self.cell_size_on_screen + 32, legend_y + self.cell_size_on_screen/2), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignHCenter)
        self._add_aligned_text(k_value_text, self.fonts['legend'], self.colors['dark'], QPointF(legend_x + self.cell_size_on_screen/2, legend_y + self.cell_size_on_screen + 20), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignHCenter)
        self._add_aligned_text("縮尺: 1/5000", self.fonts['scale'], self.colors['dark'], QPointF(legend_x + self.cell_size_on_screen + 65, legend_y + 4), Qt.AlignmentFlag.AlignLeft)

    def _draw_final_result(self, calc_data):
        if calc_data['total_degree'] <= 0: return
        result_area_x, result_area_y, final_dist = self.grid_offset_x, self.grid_offset_y - 70, math.floor(calc_data['final_distance'] * 10) / 10
        dist_str, draw_second_line = f"{int(final_dist)} m" if final_dist * 10 % 10 == 0 else f"{final_dist:.1f} m", not (final_dist * 10 % 10 == 0)
        formula_text, formula_item = f"平均集材距離 = (⑨ + ⑦) ÷ ⑧ × K = ({calc_data['total_product_v']} + {calc_data['total_product_h']}) ÷ {calc_data['total_degree']} × {self.k_value:.0f} = ", self._add_aligned_text(f"平均集材距離 = (⑨ + ⑦) ÷ ⑧ × K = ({calc_data['total_product_v']} + {calc_data['total_product_h']}) ÷ {calc_data['total_degree']} × {self.k_value:.0f} = ", self.fonts['result'], self.colors['normal'], QPointF(self.grid_offset_x, self.grid_offset_y - 70), Qt.AlignmentFlag.AlignLeft, is_result=True)
        result_align_x = result_area_x + formula_item.boundingRect().width()
        self._add_aligned_text(dist_str, self.fonts['result'], self.colors['normal'], QPointF(result_align_x, result_area_y), Qt.AlignmentFlag.AlignLeft, is_result=True)
        if draw_second_line:
            second_line_y, int_dist_str = result_area_y + formula_item.boundingRect().height(), f" {int(calc_data['final_distance'] + 0.5)} m"
            self._add_aligned_text("≒", self.fonts['result'], self.colors['normal'], QPointF(result_align_x, second_line_y), Qt.AlignmentFlag.AlignRight, is_result=True)
            self._add_aligned_text(int_dist_str, self.fonts['result'], self.colors['normal'], QPointF(result_align_x, second_line_y), Qt.AlignmentFlag.AlignLeft, is_result=True)
        for item in self.result_text_items: item.setZValue(self.Z_OVERLAYS_BASE + 20)

    def _draw_calculation_tables(self, calc_data):
        v_table_x, col_widths_v, h_table_y, row_heights_h = self.grid_offset_x + self.grid_cols * self.cell_size_on_screen + 5, [40, 35, 45], self.grid_offset_y + self.grid_rows * self.cell_size_on_screen + 5, [50, 40, 50]
        headers_v_data, current_x = [("①", "走行\n(縦)\n距離"), ("②", "度数"), ("③", "①×②")], v_table_x
        for i, (num, text) in enumerate(headers_v_data):
            self._add_aligned_text(num, self.fonts['header'], self.colors['normal'], QPointF(current_x + col_widths_v[i] / 2, self.grid_offset_y - 110 + 15), Qt.AlignmentFlag.AlignHCenter)
            self._add_aligned_text(text, self.fonts['header'], self.colors['normal'], QPointF(current_x + col_widths_v[i] / 2, self.grid_offset_y - 110 + 45), Qt.AlignmentFlag.AlignHCenter)
            current_x += col_widths_v[i]
        for r in range(self.grid_rows):
            if not (calc_data['min_row'] <= r <= calc_data['max_row']) and r != calc_data['landing_row']: continue
            # ★★★ ここからが修正箇所 ★★★
            is_hl = (r == calc_data['landing_row'])
            font, color = (self.fonts['highlight'], self.colors['highlight']) if is_hl else (self.fonts['data'], self.colors['normal'])
            # ★★★ 修正ここまで ★★★
            count, vals, current_x = calc_data['row_counts'].get(r, 0), [abs(r - calc_data['landing_row']), calc_data['row_counts'].get(r, 0), abs(r - calc_data['landing_row']) * calc_data['row_counts'].get(r, 0)], v_table_x
            for i, val in enumerate(vals): self._add_aligned_text(str(val), font, color, QPointF(current_x + col_widths_v[i]/2, self.grid_offset_y + r * self.cell_size_on_screen + self.cell_size_on_screen/2)); current_x += col_widths_v[i]
        headers_h_data, current_y = [("④", "走行\n(横)距離"), ("⑤", "度数"), ("⑥", "④×⑤")], h_table_y
        for i, (num, text) in enumerate(headers_h_data):
            self._add_aligned_text(num, self.fonts['header'], self.colors['normal'], QPointF(self.grid_offset_x - 65, current_y + row_heights_h[i]/2), Qt.AlignmentFlag.AlignRight|Qt.AlignmentFlag.AlignVCenter)
            self._add_aligned_text(text, self.fonts['header'], self.colors['normal'], QPointF(self.grid_offset_x - 60, current_y + row_heights_h[i]/2), Qt.AlignmentFlag.AlignLeft|Qt.AlignmentFlag.AlignVCenter)
            current_y += row_heights_h[i]
        for c in range(self.grid_cols):
            if not (calc_data['min_col'] <= c <= calc_data['max_col']) and c != calc_data['landing_col']: continue
            # ★★★ ここからが修正箇所 ★★★
            is_hl = (c == calc_data['landing_col'])
            font, color = (self.fonts['highlight'], self.colors['highlight']) if is_hl else (self.fonts['data'], self.colors['normal'])
            # ★★★ 修正ここまで ★★★
            count, vals, current_y = calc_data['col_counts'].get(c, 0), [abs(c - calc_data['landing_col']), calc_data['col_counts'].get(c, 0), abs(c - calc_data['landing_col']) * calc_data['col_counts'].get(c, 0)], h_table_y
            for i, val in enumerate(vals): self._add_aligned_text(str(val), font, color, QPointF(self.grid_offset_x + c * self.cell_size_on_screen + self.cell_size_on_screen/2, current_y + row_heights_h[i]/2)); current_y += row_heights_h[i]
        total_cells_data = [("合計", None, v_table_x, h_table_y, col_widths_v[0], row_heights_h[0]), ("⑧", str(calc_data['total_degree']), v_table_x, h_table_y + row_heights_h[0], col_widths_v[0], row_heights_h[1]), ("⑦", str(calc_data['total_product_h']), v_table_x, h_table_y + sum(row_heights_h[:2]), col_widths_v[0], row_heights_h[2]), ("⑧", str(calc_data['total_degree']), v_table_x + col_widths_v[0], h_table_y, col_widths_v[1], row_heights_h[0]), ("⑨", str(calc_data['total_product_v']), v_table_x + sum(col_widths_v[:2]), h_table_y, col_widths_v[2], row_heights_h[0])]
        for symbol, value, x, y, w, h in total_cells_data:
            if value is None: self._add_aligned_text(symbol, self.fonts['total'], self.colors['normal'], QPointF(x + w/2, y + h/2))
            else: self._add_aligned_text(symbol, self.fonts['total'], self.colors['normal'], QPointF(x + w/2, y + h/3)); self._add_aligned_text(value, self.fonts['total'], self.colors['normal'], QPointF(x + w/2, y + h*2/3))

    def _set_all_pens_cosmetic(self, is_cosmetic):
        items_to_process = self.grid_items + self.calculation_items + self.title_items
        if self.in_area_cells_outline: items_to_process.append(self.in_area_cells_outline)
        for item in items_to_process:
            if hasattr(item, 'pen') and callable(item.pen) and hasattr(item, 'setPen'): pen = item.pen(); pen.setCosmetic(is_cosmetic); item.setPen(pen)

    def export_results(self):
        if not self.subtitle_input.text().strip(): QMessageBox.warning(self, "入力エラー", "見出しが入力されていません。\n入力して「表示」ボタンを押してから、再度エクスポートしてください。"); return
        if not self.calculation_items: QMessageBox.warning(self, "エラー", "エクスポートする内容がありません。「計算を実行」してください。"); return
        self._export_results_recursive()

    def _export_results_recursive(self, force_orientation=None, force_page_size_id=None):
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        if self.pointer_item: self.pointer_item.hide()
        try:
            if force_orientation is None:
                default_filename, (file_path, _) = f"X-Grid_{self.subtitle_input.text().strip()}" or "X-Grid_計算結果", QFileDialog.getSaveFileName(self, "結果をエクスポート", f"X-Grid_{self.subtitle_input.text().strip()}" or "X-Grid_計算結果", "PDF Document (*.pdf)")
                if not file_path:
                    if self.pointer_item: self.pointer_item.show(); QApplication.restoreOverrideCursor(); return
                self.export_file_path = file_path
            else: file_path = self.export_file_path
            content_left, content_top, content_right, content_bottom = self.grid_offset_x - 90, self.grid_offset_y - 145, self.grid_offset_x + self.grid_cols * self.cell_size_on_screen + 5 + sum([40, 35, 45]), self.grid_offset_y + self.grid_rows * self.cell_size_on_screen + 5 + sum([50, 40, 50])
            source_rect = QRectF(content_left, content_top, content_right - content_left, content_bottom - content_top)
            if file_path.lower().endswith(".pdf"):
                printer, orientation, is_a3 = QPrinter(QPrinter.PrinterMode.HighResolution), force_orientation if force_orientation is not None else self.page_orientation, self.grid_cols == self.grid_cols_a3 or (force_page_size_id == QPageSize.PageSizeId.A3)
                printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat); printer.setOutputFileName(file_path)
                page_size_id, page_layout = force_page_size_id if force_page_size_id is not None else (QPageSize.PageSizeId.A3 if is_a3 else QPageSize.PageSizeId.A4), QPageLayout(QPageSize(force_page_size_id if force_page_size_id is not None else (QPageSize.PageSizeId.A3 if is_a3 else QPageSize.PageSizeId.A4)), orientation, QMarginsF(0, 0, 0, 0), QPageLayout.Unit.Millimeter)
                printer.setPageLayout(page_layout)
                full_page_rect_mm, full_page_rect_px = page_layout.fullRect(QPageLayout.Unit.Millimeter), printer.pageRect(QPrinter.Unit.DevicePixel)
                dpmm_x, dpmm_y, margin_mm, mm_per_scene_unit = full_page_rect_px.width() / full_page_rect_mm.width(), full_page_rect_px.height() / full_page_rect_mm.height(), 5.0, (self.k_value / self.cell_size_on_screen) / 5.0
                target_width_mm, target_height_mm, printable_width_mm, printable_height_mm = source_rect.width() * mm_per_scene_unit, source_rect.height() * mm_per_scene_unit, full_page_rect_mm.width() - 10.0, full_page_rect_mm.height() - 10.0
                if target_width_mm > printable_width_mm or target_height_mm > printable_height_mm:
                    msg_box = QMessageBox(self); msg_box.setIcon(QMessageBox.Icon.Warning); msg_box.setWindowTitle("サイズ超過")
                    msg_box.setText(f"1:5000スケールではコンテンツが用紙サイズ({page_layout.pageSize().name()})の印刷可能領域に収まりません。\n\n<b>必要サイズ:</b> {target_width_mm:.1f} x {target_height_mm:.1f} mm\n<b>印刷可能領域 (マージン{margin_mm:.0f}mm):</b> {printable_width_mm:.1f} x {printable_height_mm:.1f} mm\n\nA3サイズでエクスポートを再試行しますか？")
                    retry_button, cancel_button = msg_box.addButton("A3で再試行", QMessageBox.ButtonRole.YesRole), msg_box.addButton("キャンセル", QMessageBox.ButtonRole.NoRole); msg_box.exec()
                    if msg_box.clickedButton() == retry_button: self._export_results_recursive(QPageLayout.Orientation.Landscape, QPageSize.PageSizeId.A3)
                    if self.pointer_item: self.pointer_item.show(); QApplication.restoreOverrideCursor(); return
                target_width_px, target_height_px, offset_x_px, offset_y_px = target_width_mm * dpmm_x, target_height_mm * dpmm_y, margin_mm * dpmm_x + (printable_width_mm * dpmm_x - target_width_mm * dpmm_x) / 2.0, margin_mm * dpmm_y + (printable_height_mm * dpmm_y - target_height_mm * dpmm_y) / 2.0
                target_rect_px, pdf_painter = QRectF(offset_x_px, offset_y_px, target_width_px, target_height_px), QPainter(printer)
                self._set_all_pens_cosmetic(False)
                try: self.scene.render(pdf_painter, target_rect_px, source_rect)
                finally: self._set_all_pens_cosmetic(True); pdf_painter.end()
                QMessageBox.information(self, "成功", f"結果をPDFとして保存しました:\n{file_path}\n\n【重要】\n印刷する際は、必ず印刷設定で「実際のサイズ」または「倍率100%」を選択してください。")
        except Exception as e: QMessageBox.critical(self, "エラー", f"エクスポート中にエラーが発生しました: {e}"); self._set_all_pens_cosmetic(True)
        finally:
            if self.pointer_item: self.pointer_item.show()
            QApplication.restoreOverrideCursor()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("游ゴシック", 10))
    window = X_Grid()
    window.show()
    sys.exit(app.exec())