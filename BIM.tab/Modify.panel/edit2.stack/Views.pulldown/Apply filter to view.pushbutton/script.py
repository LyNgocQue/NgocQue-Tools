# -*- coding: utf-8 -*-
import Autodesk
import clr
clr.AddReference("RevitAPI")
from Autodesk.Revit.DB import *
from pyrevit import forms
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView
select_view = view.Id

__title__ = 'Apply filter'

color_values = [
    (162, 196, 201),
    (244, 204, 204),
    (252, 229, 205),
    (180, 235, 250),
    (203, 242, 102),
    (234, 209, 220),
    (234, 153, 153),
    (249, 203, 156),
    (255, 229, 153),
    (205, 221, 181),
    (135, 231, 176),
    (159, 197, 232),
    (225, 220, 236),
    (213, 166, 189),
    (234, 209, 220),
    (111, 168, 220),
    (220, 220, 220),
    (151, 151, 255),
    (183, 255, 183),
    (255, 155, 255),
    (145, 214, 210),
    (250, 250, 210),
    (182, 166, 137),
    (173, 216, 230),
    (183, 166, 253),
    (175, 238, 238),
    (245, 222, 179),
    (255, 228, 196),
    (124, 194, 73),
    (151, 205, 255),
    (111, 168, 220)
]

# Tạo danh sách màu bằng cách sử dụng vòng lặp
list_color = [Autodesk.Revit.DB.Color(*color) for color in color_values] * 6
try:
	t = Transaction (doc, __title__)
	t.Start()
	from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox,
									Separator, Button, CheckBox)
	components = [Label('Nhập Đối Tượng Cần Thực Hiện:'),
					ComboBox('combobox1', ["Walls","Ceilings","Floors"]),
					Label('Nhập Kiểu Filter:'),
					ComboBox('combobox2', ["Type Name","Height Offset From Level"]),
					Separator(),
					Button('OK')]
	form = FlexForm('LNQ tools', components)
	form.show()
	form.values
	input_category = str(form.values["combobox1"])
	if input_category == "Walls":
		type_element = Wall
	elif input_category == "Ceilings":
		type_element = Ceiling
	elif input_category == "Floors":
		type_element = Floor

	input_type_filter= str(form.values["combobox2"])
	if input_type_filter == "Type Name":
		list_filter_on_view = []
		list_filter_on_set = []
	# Lấy Tất cả tường, sàn, trần trong dự án 
		def apply_filter_to_view(category):
			list_name_category = []
			all_filter_category = []
			all_walls =  FilteredElementCollector(doc,view.Id).OfClass(category).WhereElementIsNotElementType().ToElements()
			for wall in all_walls:
				wall_name = wall.Name
				if wall_name not in list_name_category:
					list_name_category.append(wall_name)

	# Lấy tất cả filter trong dự án
			all_filter = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
			for i in all_filter:
				list_filter_on_set.append(i)
				filter_name = i.Name
				cateoryIds = i.GetCategories()
				for cateoryId in cateoryIds:
					category = Category.GetCategory(doc,cateoryId)
	#  lấy category của filter
					category_name = category.Name
					if category_name == input_category :
						# all_filter_category.append(filter_name)
						element_filter = i.GetElementFilter()
						# print(element_filter)
						if element_filter:
							element_parameter_filters = element_filter.GetFilters()
							# print(element_parameter_filters)
							for z in element_parameter_filters:
								rules = z.GetRules()
								for rule in rules:
									dict_BIPs =  {str(i.value__) : i for i in BuiltInParameter.GetValues(BuiltInParameter)}
									rule_param_id = rule.GetRuleParameter()
									rule_shared_param = doc.GetElement(rule_param_id)
	# Lấy parameter mà filter dùng đề so sánh
									list_rule = []
									if rule_shared_param:
										list_rule.append(rule_shared_param.Name)
										# print(rule_shared_param.Name)
									elif rule_shared_param is None:
										bip_rule_param = dict_BIPs[str(rule_param_id)]
										readable_bip = LabelUtils.GetLabelFor(bip_rule_param)
										# print(readable_bip)
										list_rule.append(readable_bip)
									for y in list_rule:
										if y == "Type Name":
	#  Lấy phép so sánh equals của filter
											if hasattr(rule, "GetEvaluator"):
												thiseval = rule.GetEvaluator().ToString().replace("Autodesk.Revit.DB.Filter","")
												thiseval = thiseval.replace("Numeric","")
												thiseval = thiseval.replace("String","")
												if thiseval == "Equals":
													if isinstance(rule, Autodesk.Revit.DB.FilterStringRule):
														all_filter_category.append(filter_name)
														value = rule.RuleString
														evaluator = rule.GetRuleParameter
														if value in list_name_category:
															list_filter_on_view.append(filter_name)
															# list_filter_on_set.append(i)
													elif isinstance(rule, Autodesk.Revit.DB.FilterValueRule):
														value = rule.RuleValue
														evaluator = rule.GetEvaluator()
														# print(evaluator)

													else:
														continue
			return all_filter_category
	# Gọi Hàm tiến hành chạy chương trình
		all_filter_category = apply_filter_to_view (type_element)
		select_filter = forms.SelectFromList.show(
			{'All': all_filter_category,
			'Only View': list_filter_on_view ,
			},
			title='MultiGroup List',
			group_selector_title='Chọn Filter:',
			multiselect=True
		)
		
		GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<Solid fill>")
		if str(GetFillPattern) == "None":
			GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<塗り潰し>")
		for i_name in select_filter:
			# print(i_name)
			override_settings = OverrideGraphicSettings()
			index_i = select_filter.index(i_name)
			random_color = list_color[index_i]
			override_settings.SetSurfaceForegroundPatternColor(random_color)
			override_settings.SetSurfaceForegroundPatternId(GetFillPattern.Id)
			override_settings.SetCutForegroundPatternColor(random_color)
			override_settings.SetCutForegroundPatternId(GetFillPattern.Id)
			for wall in list_filter_on_set:
				wall_name = Autodesk.Revit.DB.Element.Name.GetValue(wall)
				if i_name == wall_name:
					try:
						view_template = doc.GetElement(view.ViewTemplateId)
						view_template_name = Autodesk.Revit.DB.Element.Name.GetValue(view_template)
						if view_template:
							view_template.AddFilter(wall.Id)
							view_template.SetFilterOverrides(wall.Id, override_settings)
							index = list_filter_on_set.index(wall)
							list_filter_on_set.pop(index)
						else:
							view.AddFilter(wall.Id)
							view.SetFilterOverrides(wall.Id, override_settings)
							# index = list_filter_on_set.index(wall)
							# list_filter_on_set.pop(index)
					except:
						pass
	elif input_type_filter == "Height Offset From Level":
		list_filter_on_view = []
		list_filter_on_set = []
		def apply_filter_to_view(category):
			list_name_category = []
			all_filter_category = []
			all_walls =  FilteredElementCollector(doc,view.Id).OfClass(category).WhereElementIsNotElementType().ToElements()
			for wall in all_walls:
				caodo_1 = wall.LookupParameter("Height Offset From Level").AsValueString()
				cao_do = float(caodo_1)
				wall_name = wall.Name
				list_name_category.append(cao_do)
				# print(cao_do)
			# list_name_category = list(set(list_name_category))
			# print(list_name_category)
			all_filter = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()
			for i in all_filter:
				list_filter_on_set.append(i)
				filter_name = i.Name
				cateoryIds = i.GetCategories()
				for cateoryId in cateoryIds:
					category = Category.GetCategory(doc,cateoryId)
					category_name = category.Name
					if category_name == input_category :
						element_filter = i.GetElementFilter()
						if element_filter:
							element_parameter_filters = element_filter.GetFilters()
							for z in element_parameter_filters:
								rules = z.GetRules() 
								for rule in rules:
									dict_BIPs =  {str(i.value__) : i for i in BuiltInParameter.GetValues(BuiltInParameter)}
									rule_param_id = rule.GetRuleParameter()
									rule_shared_param = doc.GetElement(rule_param_id)
									list_rule = []
									if rule_shared_param:
										list_rule.append(rule_shared_param.Name)
										# print(rule_shared_param.Name)
									elif rule_shared_param is None:
										bip_rule_param = dict_BIPs[str(rule_param_id)]
										readable_bip = LabelUtils.GetLabelFor(bip_rule_param)
										# print(readable_bip)
										list_rule.append(readable_bip)
									for y in list_rule:
										if y == "Height Offset From Level":
											if hasattr(rule, "GetEvaluator"):
												thiseval = rule.GetEvaluator().ToString().replace("Autodesk.Revit.DB.Filter","")
												thiseval = thiseval.replace("Numeric","")
												thiseval = thiseval.replace("String","")
												if thiseval == "Equals":
													# list_value = []
													if isinstance(rule, Autodesk.Revit.DB.FilterStringRule):
														value = rule.RuleString
														evaluator = rule.GetRuleParameter
													elif isinstance(rule, Autodesk.Revit.DB.FilterValueRule):
														all_filter_category.append(filter_name)
														value = round(rule.RuleValue * 304.8, 4)
														# list_value.append(value)
														evaluator = rule.GetEvaluator()
														if value in list_name_category:
															list_filter_on_view.append(filter_name)
													else:
														continue


			return all_filter_category
		all_filter_category = apply_filter_to_view (type_element)
		select_filter = forms.SelectFromList.show(
			{'All': all_filter_category,
			'Only View': list_filter_on_view ,
			},
			title='MultiGroup List',
			group_selector_title='Chọn Filter:',
			multiselect=True
		)
		GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<Solid fill>")
		if str(GetFillPattern) == "None":
			GetFillPattern = Autodesk.Revit.DB.FillPatternElement.GetFillPatternElementByName(doc,FillPatternTarget.Drafting,"<塗り潰し>")
		for i_name in select_filter:
			# print(i_name)
			override_settings = OverrideGraphicSettings()
			index_i = select_filter.index(i_name)
			random_color = list_color[index_i]
			override_settings.SetSurfaceForegroundPatternColor(random_color)
			override_settings.SetSurfaceForegroundPatternId(GetFillPattern.Id)
			override_settings.SetCutForegroundPatternColor(random_color)
			override_settings.SetCutForegroundPatternId(GetFillPattern.Id)
			for wall in list_filter_on_set:
				wall_name = Autodesk.Revit.DB.Element.Name.GetValue(wall)
				if i_name == wall_name:
					view_template = doc.GetElement(view.ViewTemplateId)
					view_template_name = Autodesk.Revit.DB.Element.Name.GetValue(view_template)
					if view_template:
						view_template.AddFilter(wall.Id)
						view_template.SetFilterOverrides(wall.Id, override_settings)
						index = list_filter_on_set.index(wall)
						list_filter_on_set.pop(index)
					else:
						view.AddFilter(wall.Id)
						view.SetFilterOverrides(wall.Id, override_settings)
						# index = list_filter_on_set.index(wall)
						# list_filter_on_set.pop(index)
	t.Commit() 
except:
	pass
