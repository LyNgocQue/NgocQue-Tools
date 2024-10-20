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
	def on_ok_click(self, sender, args):
		text_1 = self.text1.Text
		text_3 = self.text3.Text
		# Cho phép người dùng chọn các đối tượng trong Revit
		selection = uidoc.Selection.GetElementIds()
		for obj in selection:
			ele = doc.GetElement(obj)
			elementtype = doc.GetElement(ele.GetTypeId())
			family_name = elementtype.FamilyName
			family_type = elementtype.get_Parameter (BuiltInParameter. ALL_MODEL_TYPE_NAME)

			type_name =  family_type.AsString()
			ten = "Sections" + " (" + type_name + ")"
			new_name = str(text_1) + type_name + str(text_3) 
			category = List[ElementId]()
			category.Add(ElementId(BuiltInCategory.OST_Sections))
			# category.Add(ElementId(BuiltInCategory.OST_Callouts))
			sym_name_param = ElementId(BuiltInParameter.VIEW_TYPE)
			rules = List[FilterRule]()
			rules.Add(ParameterFilterRuleFactory.CreateNotEqualsRule(sym_name_param, ten, False))
			for rule in rules:
				pfilter = ElementParameterFilter(rule)
				# filter = ParameterFilterElement.Create(doc, new_name, category, pfilter)
				filters_and = List[DB.ElementFilter]()
				filters_and.Add(pfilter)
				rule_set_and = DB.LogicalAndFilter(filters_and)
				all_par_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
				all_par_filters_names = [f.Name for f in all_par_filters]
				t = Transaction (doc, "Create filter Section")
				t.Start()
				if new_name not in all_par_filters_names:
					filter = ParameterFilterElement.Create(doc, new_name, category, rule_set_and)
				else:
					pass
				t.Commit()
				self.Close()
	def on_apply_to_view_click(self, sender, args):
		text_3 = self.text3.Text  # Lấy nội dung từ TextBox 3 
		text_1 = self.text1.Text
		# Bắt đầu giao dịch cho thao tác áp dụng filter
		t = Transaction(doc, "Apply filters to view")
		t.Start()

		# Danh sách để lưu các filter đã tạo
		created_filters = []			
		selection = uidoc.Selection.GetElementIds()
		for obj in selection:
			ele = doc.GetElement(obj)
			elementtype = doc.GetElement(ele.GetTypeId())
			family_name = elementtype.FamilyName
			family_type = elementtype.get_Parameter (BuiltInParameter. ALL_MODEL_TYPE_NAME)

			type_name =  family_type.AsString()
			ten = "Sections" + " (" + type_name + ")"
			new_name = str(text_1) + type_name + str(text_3) 
			category = List[ElementId]()
			category.Add(ElementId(BuiltInCategory.OST_Sections))
			# category.Add(ElementId(BuiltInCategory.OST_Callouts))
			sym_name_param = ElementId(BuiltInParameter.VIEW_TYPE)
			rules = List[FilterRule]()
			rules.Add(ParameterFilterRuleFactory.CreateNotEqualsRule(sym_name_param, ten, False))
			created_filters = []
			for rule in rules:
				pfilter = ElementParameterFilter(rule)
				# filter = ParameterFilterElement.Create(doc, new_name, category, pfilter)
				filters_and = List[DB.ElementFilter]()
				filters_and.Add(pfilter)
				rule_set_and = DB.LogicalAndFilter(filters_and)
				all_par_filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
				all_par_filters_names = [f.Name for f in all_par_filters]
				if new_name not in all_par_filters_names:
					view_filter = ParameterFilterElement.Create(doc, new_name, category, rule_set_and)
					created_filters.append(view_filter)
				else:
					pass


				active_view = doc.ActiveView
				existing_filters = active_view.GetFilters()
				filter_ids = [f.IntegerValue for f in existing_filters]  # Lấy ID dạng Integer


				# Kiểm tra xem view hiện tại có template không
				template_id = active_view.ViewTemplateId

				# Lưu danh sách filter có thể áp dụng, bao gồm cả filter đã tạo và đã tồn tại
				all_filters_to_apply = created_filters + [f for f in all_par_filters if f.Name == new_name]

				# Xác định view để áp dụng filter
				target_view = doc.GetElement(template_id) if template_id != ElementId.InvalidElementId else active_view


				for view_filter in all_filters_to_apply:
					if view_filter.Id.IntegerValue not in filter_ids:  # Không áp dụng nếu đã có
						try:
							target_view.SetFilterVisibility(view_filter.Id,False)
						except Exception as e:
							pass
							# print("Lỗi khi áp dụng filter '{view_filter.Name}': {e}")

		t.Commit()  # Kết thúc giao dịch
		self.Close() 
if __name__ == "__main__":
	form = ModalForm('secNotEquals.xaml')
	form.ShowDialog()
