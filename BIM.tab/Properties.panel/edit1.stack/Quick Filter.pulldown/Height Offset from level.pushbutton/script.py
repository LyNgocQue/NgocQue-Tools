# -*- coding: utf-8 -*-

import clr
import Autodesk
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection  import ObjectType
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox, Form, CheckedListBox, Button, TextBox, Label, Panel, ComboBox
from pyrevit.forms import WPFWindow
from Autodesk.Revit import DB
from System.Collections.Generic import List
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

__title__ = 'Height Offset from level'
color_values = [
	(170, 205, 215),  # Soft Teal
	(255, 230, 200),  # Soft Peach
	(240, 220, 210),  # Soft Blush
	(210, 230, 210),  # Soft Mint
	(230, 200, 200),  # Soft Coral
	(250, 220, 190),  # Soft Apricot
	(200, 220, 240),  # Soft Sky Blue
	(220, 240, 230),  # Soft Mint Cream
	(240, 240, 200),  # Soft Lemon
	(200, 250, 200),  # Soft Light Green
	(215, 210, 255),  # Soft Lilac
	(255, 200, 230),  # Soft Pink
	(210, 200, 240),  # Soft Lavender
	(255, 240, 220),  # Soft Salmon
	(200, 200, 210),  # Soft Steel Blue
	(240, 220, 230),  # Soft Orchid
	(240, 230, 205),  # Soft Beige
	(250, 230, 210),  # Soft Honey
	(220, 220, 200),  # Soft Grey
	(225, 200, 225),  # Soft Mauve
	(200, 240, 220),  # Soft Turquoise
	(210, 210, 240),  # Soft Periwinkle
	(240, 205, 205),  # Soft Rose
	(215, 205, 230),  # Soft Lavender Mist
	(210, 230, 220),  # Soft Aqua
	(200, 250, 255),  # Soft Light Cyan
	(240, 220, 240),  # Soft Orchid
	(230, 240, 210),  # Soft Pale Green
	(200, 225, 255),  # Soft Light Blue
	(240, 240, 230),  # Soft Mist
	(255, 250, 240)   # Soft Floral White
]

# Tạo danh sách màu bằng cách sử dụng vòng lặp
list_color = [Autodesk.Revit.DB.Color(*color) for color in color_values] * 6

try:
	class ModalForm(WPFWindow):
		def __init__(self, xaml_file_name):
			WPFWindow.__init__(self, xaml_file_name)

			self.text1 = self.FindName("text1")
			self.text2 = self.FindName("text2")
			self.text3 = self.FindName("text3")
			
			self.OK = self.FindName("OK")  # Tìm nút Create
			self.ApplyToView = self.FindName("ApplyToView")  # Tìm nút Apply to View
			
			self.OK.Click += self.on_ok_click  # Liên kết sự kiện nhấp nút Create
			self.ApplyToView.Click += self.on_apply_to_view_click  # Liên kết sự kiện cho nút Apply to View
			self.selected_elements = []  # Lưu danh sách các đối tượng đã chọn
			self.select_elements()  # Gọi hàm để chọn đối tượng ngay khi mở form

		def select_elements(self):
			# Cho phép người dùng chọn các đối tượng trong Revit
			selection = uidoc.Selection.GetElementIds()
			self.selected_elements1 = selection  # Lưu thông tin đối tượng đã chọn

			for obj in selection:
				ele = doc.GetElement(obj)
				category = (ele.Category).Name
				if category not in ["Floors", "Ceilings"]:
					MessageBox.Show("Đối tượng " +category.upper()+ " không có thuộc tính HEIGHT OFFSET FROM LEVEL", "Thông Báo !")
					self.Close()  # Đóng cửa sổ nếu không hỗ trợ
					return  # Thoát hàm
				
				# Cập nhật text1 dựa vào category
				self.update_text1(category)
		def update_text1(self, category):
			# Thiết lập text_1 dựa vào category
			text_map = {
						"Floors": "FloorS_height_e_",
						"Ceilings": "Ceiling_height_e_",
					}
			text_1 = text_map.get(category)
			if text_1 is None:
				return  # Bỏ qua nếu category không hỗ trợ
			
			# Cập nhật TextBlock để hiển thị text_1
			self.text1.Text = text_1  # Cập nhật text1 ngay lập tức

		def on_ok_click(self, sender, args):
			text_3 = self.text3.Text  # Lấy nội dung từ TextBox 3 
			t = Transaction (doc, "Create filter ")
			t.Start()
			for obj in self.selected_elements1:
				ele2 = doc.GetElement(obj)
				category = (ele2.Category).Name
				Height_offset_from_level = ele2.LookupParameter("Height Offset From Level").AsDouble()*304.8
				new_name = self.text1.Text + str(Height_offset_from_level) + text_3
				cats = List[ElementId]()
				category_map = {
									"Floors": BuiltInCategory.OST_Floors,
									"Ceilings": BuiltInCategory.OST_Ceilings,

								}
				if category in category_map:
					cats.Add(ElementId(category_map[category]))
				if category == "Ceilings":
					sym_name_param = ElementId(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
				elif category == "Floors":
					sym_name_param = ElementId(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
				rules = List[FilterRule]()
				rules.Add(ParameterFilterRuleFactory.CreateEqualsRule(sym_name_param, float(Height_offset_from_level)/304.8, False))

				for rule in rules:
					pfilter = ElementParameterFilter(rule)
					# filter = ParameterFilterElement.Create(doc, new_name, category, pfilter)
					filters_and = List[DB.ElementFilter]()
					filters_and.Add(pfilter)
					rule_set_and = DB.LogicalAndFilter(filters_and)
					all_par_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
					all_par_filters_names = [f.Name for f in all_par_filters]
						
					if new_name not in all_par_filters_names:
						filter = ParameterFilterElement.Create(doc, new_name, cats, rule_set_and)
					else:
						pass
			t.Commit()
			self.Close()  # Đóng cửa sổ sau khi xử lý xong

		def on_apply_to_view_click(self, sender, args):
			text_3 = self.text3.Text  # Lấy nội dung từ TextBox 3 

			# Bắt đầu giao dịch cho thao tác áp dụng filter
			t = Transaction(doc, "Apply filters to view")
			t.Start()

			# Danh sách để lưu các filter đã tạo
			created_filters = []

			# Duyệt qua các đối tượng đã chọn để tạo filter nếu chưa tồn tại
			for obj in self.selected_elements1:
				ele2 = doc.GetElement(obj)
				category = (ele2.Category).Name
				Height_offset_from_level = ele2.LookupParameter("Height Offset From Level").AsDouble()*304.8
				new_name = self.text1.Text + str(Height_offset_from_level) + text_3
				cats = List[ElementId]()
				category_map = {
									"Floors": BuiltInCategory.OST_Floors,
									"Ceilings": BuiltInCategory.OST_Ceilings,

								}
				if category in category_map:
					cats.Add(ElementId(category_map[category]))
				if category == "Ceilings":
					sym_name_param = ElementId(BuiltInParameter.CEILING_HEIGHTABOVELEVEL_PARAM)
				elif category == "Floors":
					sym_name_param = ElementId(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM)
				rules = List[FilterRule]()
				rules.Add(ParameterFilterRuleFactory.CreateEqualsRule(sym_name_param, float(Height_offset_from_level)/304.8, False))
				for rule in rules:
					pfilter = ElementParameterFilter(rule)
					filters_and = List[DB.ElementFilter]()
					filters_and.Add(pfilter)
					rule_set_and = DB.LogicalAndFilter(filters_and)
					all_par_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
					all_par_filters_names = [f.Name for f in all_par_filters]
					if new_name not in all_par_filters_names:
						view_filter = ParameterFilterElement.Create(doc, new_name, cats, rule_set_and)
						created_filters.append(view_filter)  # Lưu filter đã tạo
					else:
						existing_filter = next(f for f in all_par_filters if f.Name == new_name)
						created_filters.append(existing_filter)
						

			# Áp dụng tất cả filter vào view hiện tại
			active_view = doc.ActiveView
			existing_filters = active_view.GetFilters()
			filter_ids = [f.IntegerValue for f in existing_filters]  # Lấy ID dạng Integer
			# print(filter_ids)

			# Kiểm tra xem view hiện tại có template không
			template_id = active_view.ViewTemplateId
			# print(template_id)

			# Lưu danh sách filter có thể áp dụng, bao gồm cả filter đã tạo và đã tồn tại

			all_filters_to_apply = created_filters 

			# Xác định view để áp dụng filter
			target_view = doc.GetElement(template_id) if template_id != ElementId.InvalidElementId else active_view
			# print(target_view)

			for view_filter in all_filters_to_apply:
				if view_filter.Id.IntegerValue not in filter_ids:  # Không áp dụng nếu đã có
					GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<Solid fill>")
					if str(GetFillPattern) == "None":
						GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<塗り潰し>")
					index_i = all_filters_to_apply.index(view_filter)
					random_color = list_color[index_i]
					override_settings = OverrideGraphicSettings()
					# override_settings.SetCutLineColor(random_color)
					# override_settings.SetCutLineWeight(5)
					# override_settings.SetSurfaceTransparency(50)
					override_settings.SetSurfaceForegroundPatternColor(random_color)
					override_settings.SetSurfaceForegroundPatternId(GetFillPattern.Id)
					override_settings.SetCutForegroundPatternColor(random_color)
					override_settings.SetCutForegroundPatternId(GetFillPattern.Id)

					try:
						# Áp dụng ngay cả khi filter bỏ dấu tick VG overrides Filters
						active_view.AddFilter(view_filter.Id)
						active_view.SetFilterOverrides(view_filter.Id, override_settings)
						# Áp dụng vào view template
						target_view.AddFilter(view_filter.Id)
						target_view.SetFilterOverrides(view_filter.Id, override_settings)
					except Exception as e:
						pass
						# print("Lỗi khi áp dụng filter '{view_filter.Name}': {e}")

			t.Commit()  # Kết thúc giao dịch
			self.Close()  # Đóng cửa sổ sau khi xử lý xong

	# Khởi chạy ứng dụng
	if __name__ == "__main__":
		form = ModalForm('Height Offset from level.xaml')
		form.ShowDialog()  # Hiển thị cửa sổ
except :
	pass