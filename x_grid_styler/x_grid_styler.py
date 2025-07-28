import os
from qgis.PyQt.QtWidgets import QAction, QDialog, QVBoxLayout, QComboBox, QDialogButtonBox, QMessageBox, QApplication
from qgis.PyQt.QtCore import QVariant, Qt
from qgis.PyQt.QtGui import QColor, QIcon
from qgis.core import (
    QgsProject, QgsVectorLayer, QgsField, QgsRenderContext, QgsSymbol,
    QgsSymbolLayer, QgsSimpleFillSymbolLayer, QgsSimpleLineSymbolLayer,
    QgsGeometryGeneratorSymbolLayer, QgsSingleSymbolRenderer, 
    QgsCategorizedSymbolRenderer, QgsRuleBasedRenderer, QgsVectorDataProvider
)
from qgis.utils import iface

class X_Grid_StylerDialog(QDialog):
    def __init__(self, parent_plugin):
        super().__init__(parent_plugin.iface.mainWindow())
        self.parent_plugin = parent_plugin
        self.setWindowTitle("X-Grid用スタイル書き出し")
        self.layout = QVBoxLayout(self)
        self.layer_combo = QComboBox(self)
        self.button_box = QDialogButtonBox(self)
        self.export_button = self.button_box.addButton("書き出し", QDialogButtonBox.AcceptRole)
        self.close_button = self.button_box.addButton("閉じる", QDialogButtonBox.RejectRole)
        self.layout.addWidget(self.layer_combo)
        self.layout.addWidget(self.button_box)
        self.export_button.clicked.connect(self.on_export)
        self.close_button.clicked.connect(self.close)
        project = QgsProject.instance()
        if hasattr(project, 'layersChanged'):
            project.layersChanged.connect(self.populate_layers)
        else:
            project.layersAdded.connect(self.populate_layers)
            project.layersWillBeRemoved.connect(self.populate_layers)
        self.populate_layers()

    def populate_layers(self, layers=None):
        current_id = self.layer_combo.currentData()
        self.layer_combo.clear()
        layers = [layer for layer in QgsProject.instance().mapLayers().values() if isinstance(layer, QgsVectorLayer)]
        for layer in layers:
            self.layer_combo.addItem(layer.name(), layer.id())
        index = self.layer_combo.findData(current_id)
        if index != -1:
            self.layer_combo.setCurrentIndex(index)

    def on_export(self):
        layer_id = self.layer_combo.currentData()
        if not layer_id:
            self.parent_plugin.iface.messageBar().pushWarning("エラー", "対象レイヤーが選択されていません。")
            return
        layer = QgsProject.instance().mapLayer(layer_id)
        if layer:
            self.parent_plugin.export_styles_to_attributes(layer)
        else:
            self.parent_plugin.iface.messageBar().pushWarning("エラー", "選択されたレイヤーが見つかりません。")

class X_Grid_Styler:
    def __init__(self, iface):
        self.iface = iface
        self.action = None
        self.dialog = None
        try:
            self.dash_map = {
                Qt.PenStyle.NoPen: "none", Qt.PenStyle.SolidLine: "solid",
                Qt.PenStyle.DashLine: "dash", Qt.PenStyle.DotLine: "dot",
                Qt.PenStyle.DashDotLine: "dashdot", Qt.PenStyle.DashDotDotLine: "dashdotdot",
                Qt.PenStyle.CustomDashLine: "custom",
            }
        except AttributeError:
             self.dash_map = {0: "none", 1: "solid", 2: "dash", 3: "dot", 4: "dashdot", 5: "dashdotdot", 6: "custom"}

    def initGui(self):
        icon_path = os.path.join(os.path.dirname(__file__), 'icon.png')
        self.action = QAction(
            QIcon(icon_path), "X-Grid用スタイルを書き出し", self.iface.mainWindow()
        )
        self.action.triggered.connect(self.run)
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("X-Grid", self.action)

    def unload(self):
        self.iface.removePluginMenu("X-Grid", self.action)
        self.iface.removeToolBarIcon(self.action)
        if self.dialog:
            self.dialog.close()

    def run(self):
        if not self.dialog:
            self.dialog = X_Grid_StylerDialog(self)
        self.dialog.populate_layers()
        self.dialog.show()
        self.dialog.raise_()
        self.dialog.activateWindow()

    def _get_style_properties(self, symbol: QgsSymbol) -> dict:
        if not symbol:
            return {}
        
        props = { "fill_color": "", "strk_color": "", "strk_width": "0.0", "strk_style": "solid", "dash_pattn": "" }

        for i in range(symbol.symbolLayerCount()):
            s_layer = symbol.symbolLayer(i)
            if hasattr(s_layer, 'isEnabled') and not s_layer.isEnabled(): continue

            if isinstance(s_layer, QgsSimpleFillSymbolLayer):
                if s_layer.brushStyle() != Qt.BrushStyle.NoBrush:
                    fill_color = s_layer.color()
                    if fill_color.alpha() > 0:
                        props['fill_color'] = fill_color.name(QColor.NameFormat.HexArgb)
                
                stroke_color = s_layer.strokeColor()
                # 枠線が「なし」または「透明」の場合、スタイルを 'none' に設定
                if s_layer.strokeStyle() == Qt.PenStyle.NoPen or stroke_color.alpha() == 0:
                    props['strk_style'] = "none"
                    props['strk_color'] = ""
                    props['strk_width'] = "0.0"
                    props['dash_pattn'] = ""
                else:
                    props['strk_color'] = stroke_color.name(QColor.NameFormat.HexArgb)
                    props['strk_width'] = str(s_layer.strokeWidth())
                    pen_style_val = int(s_layer.strokeStyle())
                    props['strk_style'] = self.dash_map.get(pen_style_val, "solid")
                    if hasattr(s_layer, 'useCustomDashPattern') and s_layer.useCustomDashPattern():
                        props['dash_pattn'] = ",".join(map(str, s_layer.customDashPattern()))
                    else: 
                        props['dash_pattn'] = ""
            
            elif isinstance(s_layer, QgsSimpleLineSymbolLayer):
                color = s_layer.color()
                # 線が「なし」または「透明」の場合、スタイルを 'none' に設定
                if s_layer.penStyle() == Qt.PenStyle.NoPen or color.alpha() == 0:
                    props['strk_style'] = "none"
                    props['strk_color'] = ""
                    props['strk_width'] = "0.0"
                    props['dash_pattn'] = ""
                else:
                    props['strk_color'] = color.name(QColor.NameFormat.HexArgb)
                    props['strk_width'] = str(s_layer.width())
                    pen_style_val = int(s_layer.penStyle())
                    props['strk_style'] = self.dash_map.get(pen_style_val, "solid")
                    if hasattr(s_layer, 'useCustomDashPattern') and s_layer.useCustomDashPattern():
                        props['dash_pattn'] = ",".join(map(str, s_layer.customDashPattern()))
                    else: 
                        props['dash_pattn'] = ""
        return props

    def _update_feature_attributes(self, layer, feature, style_data, field_map):
        fid = feature.id()
        for field_name, value in style_data.items():
            field_idx = field_map.get(field_name, -1)
            if field_idx != -1:
                layer.changeAttributeValue(fid, field_idx, value)

    def export_styles_to_attributes(self, layer):
        if not isinstance(layer, QgsVectorLayer):
            self.iface.messageBar().pushWarning("エラー", "ベクターレイヤーを選択してください。")
            return

        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            required_fields = ["style_cat", "fill_color", "strk_color", "strk_width", "strk_style", "dash_pattn"]
            provider = layer.dataProvider()

            layer.startEditing()
            existing_fields = [field.name() for field in provider.fields()]
            
            if provider.capabilities() & QgsVectorDataProvider.DeleteAttributes:
                fields_to_delete_names = [name for name in existing_fields if name not in required_fields and (name.startswith('style_') or name.startswith('stroke_') or name.startswith('dash_') or name.startswith('strk_'))]
                if fields_to_delete_names:
                    indices_to_delete = [existing_fields.index(name) for name in fields_to_delete_names]
                    provider.deleteAttributes(indices_to_delete)
                    layer.updateFields()
                    existing_fields = [field.name() for field in provider.fields()]

            fields_to_add = [QgsField(f, QVariant.String) for f in required_fields if f not in existing_fields]
            if fields_to_add:
                provider.addAttributes(fields_to_add)
                layer.updateFields()
            
            if layer.isModified():
                layer.commitChanges()

            layer.startEditing()
            renderer = layer.renderer()
            context = QgsRenderContext.fromMapSettings(iface.mapCanvas().mapSettings())
            field_map = {field.name(): i for i, field in enumerate(layer.fields())}

            if isinstance(renderer, QgsSingleSymbolRenderer):
                symbol = renderer.symbol()
                style_props = self._get_style_properties(symbol)
                style_props["style_cat"] = "default"
                for f in layer.getFeatures():
                    self._update_feature_attributes(layer, f, style_props, field_map)
            
            elif isinstance(renderer, (QgsCategorizedSymbolRenderer, QgsRuleBasedRenderer)):
                 for f in layer.getFeatures():
                    symbol = renderer.symbolForFeature(f, context)
                    if not symbol: continue
                    style_props = self._get_style_properties(symbol)
                    
                    if isinstance(renderer, QgsCategorizedSymbolRenderer):
                        category = renderer.category(f)
                        style_props["style_cat"] = category.label() if category else "other"
                    else: # QgsRuleBasedRenderer
                        rule = renderer.ruleForFeature(f)
                        style_props["style_cat"] = rule.label() if rule else "other"
                    
                    self._update_feature_attributes(layer, f, style_props, field_map)
            
            if layer.isModified():
                layer.commitChanges()
                self.iface.messageBar().pushSuccess("完了", f"'{layer.name()}'のスタイル属性を書き出しました。")
            else:
                layer.rollBack()
                self.iface.messageBar().pushInfo("情報", "書き出すスタイルの変更はありませんでした。")

        except Exception as e:
            if layer.isEditable(): layer.rollBack()
            self.iface.messageBar().pushCritical("エラー発生", f"処理中にエラーが発生しました: {e}")
        finally:
            QApplication.restoreOverrideCursor()

def classFactory(iface):
    from .x_grid_styler import X_Grid_Styler
    return X_Grid_Styler(iface)