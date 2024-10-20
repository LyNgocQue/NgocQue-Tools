# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *
from pyrevit import forms
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

__title__ = 'Copy Filters'

# Lấy view hiện hành đang mở
src_view = view

# Kiểm tra xem view hiện hành có chứa bộ lọc không
filter_ids = src_view.GetFilters()
if not filter_ids:
	forms.alert('View này không có filters.', exitscript=True)

# Lấy danh sách các bộ lọc trong view hiện hành
filters = [doc.GetElement(f_id) for f_id in filter_ids]
dict_filters = {f.Name:f for f in filters}

# Cho phép người dùng chọn các bộ lọc cần sao chép
selected_filters = forms.SelectFromList.show(sorted(dict_filters),
											title='Select filters to copy',
											multiselect=True,
											button_name='OK')
if not selected_filters:
	forms.alert('Không có filters nào được chọn. Vui Lòng Thử Lại.', exitscript=True)
filter_to_copy = [dict_filters[f_name] for f_name in selected_filters]

# Cho phép người dùng chọn các view đích để sao chép bộ lọc
# tat_ca_view = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
# floor_plan_views = []
# for view in tat_ca_view:
# 	if view.ViewType == ViewType.FloorPlan:
# 		floor_plan_views.append(view.Name)

# for name in floor_plan_views:
# 	print(name)
dict_all_views = {v.Name:v for v in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
					WhereElementIsNotElementType().ToElements()}
floor_plan_views = {view.Name:view for view in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
					WhereElementIsNotElementType().ToElements()
					if view.ViewType == ViewType.FloorPlan}
section_views = {view.Name:view for view in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
					WhereElementIsNotElementType().ToElements()
					if view.ViewType == ViewType.Section}
Elevation_views = {view.Name:view for view in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
					WhereElementIsNotElementType().ToElements()
					if view.ViewType == ViewType.Elevation}
Ceiling_Plan_views = {view.Name:view for view in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
					WhereElementIsNotElementType().ToElements()
					if view.ViewType == ViewType.CeilingPlan}
structural_plan_views = {view.Name:view for view in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).
							WhereElementIsNotElementType().ToElements()
							if view.ViewType == ViewType.EngineeringPlan}
selected_dest_views = forms.SelectFromList.show({'All view': sorted(dict_all_views),
												'Floor_Plan' :sorted(floor_plan_views),
												'Section' :sorted(section_views),
												'Elevation' :sorted(Elevation_views),
												'Ceiling' :sorted(Ceiling_Plan_views),
												'Structural Plan' :sorted(structural_plan_views)},
												title='Chọn View Để Paste filters',
												group_selector_title='Select List:',
												multiselect=True,
												button_name='OK')
if not selected_dest_views:
	forms.alert('Chưa chọn view để paste filters. Vui Lòng Thử Lại.', exitscript=True)
dest_views = [dict_all_views[v_name] for v_name in selected_dest_views]

# Sao chép các bộ lọc sang các view đích
with Transaction(doc, __title__) as t:
	t.Start()
	for view_filter in filter_to_copy:
		filter_overrides = src_view.GetFilterOverrides(view_filter.Id)
		for view in dest_views:
			# Kiểm tra xem view có template không
			if view.ViewTemplateId != ElementId.InvalidElementId:
				# Nếu view có template, lấy template
				template_view = doc.GetElement(view.ViewTemplateId)	
				# Áp dụng thiết lập bộ lọc vào template
				template_view.SetFilterOverrides(view_filter.Id, filter_overrides)
				view.SetFilterOverrides(view_filter.Id, filter_overrides)
			else:
				# Nếu view không có template, kiểm tra và thêm bộ lọc vào view đó
				if view_filter.Id not in view.GetFilters():
					view.AddFilter(view_filter.Id)  # Thêm bộ lọc vào view hiện tại
				
				# Áp dụng các thiết lập bộ lọc vào view hiện tại
				view.SetFilterOverrides(view_filter.Id, filter_overrides)

	t.Commit()









































# with Transaction(doc, __title__) as t:
# 	t.Start()
# 	for view_filter in filter_to_copy:
# 		filter_overrides = src_view.GetFilterOverrides(view_filter.Id)
# 		for view in dest_views:
# 			# Kiểm tra xem view có template không
# 			if view.ViewTemplateId != ElementId.InvalidElementId:
# 				# Nếu view có template, lấy template
# 				template_view = doc.GetElement(view.ViewTemplateId)

# 				# Kiểm tra xem VG Overrides Filter của template có được bật hay không
# 				current_filters = template_view.GetFilters()
# 				if view_filter.Id not in current_filters:
# 					# Nếu bộ lọc chưa có trong template, thêm vào template
# 					template_view.AddFilter(view_filter.Id)

# 				# Áp dụng thiết lập bộ lọc vào template
				
# 				template_view.SetFilterOverrides(view_filter.Id, filter_overrides)
# 			else:
# 				# Nếu view không có template, kiểm tra và thêm bộ lọc vào view đó
# 				if view_filter.Id not in view.GetFilters():
# 					view.AddFilter(view_filter.Id)  # Thêm bộ lọc vào view hiện tại
				
# 				# Áp dụng các thiết lập bộ lọc vào view hiện tại
# 				view.SetFilterOverrides(view_filter.Id, filter_overrides)

# 	t.Commit()




# # -*- coding: utf-8 -*-

# import clr
# clr.AddReference('RevitAPI')
# from Autodesk.Revit.DB import *
# from rpw.ui.forms import TextInput, Alert # import giao diện
# from Autodesk.Revit.UI.Selection.Selection import PickObject
# from Autodesk.Revit.UI.Selection  import ObjectType
# from pyrevit import forms
# uidoc = __revit__.ActiveUIDocument
# doc = uidoc.Document
# view = doc.ActiveView
# __title__ = 'Nháp'

# all_views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
# views_with_filters = [v for v in all_views if v.GetFilters()]
# # print(views_with_filters)
# if not views_with_filters:
#     # print("chans")
# 	forms.alert('khong co filter duoc chon', exitscript = True)
# dict_views_with_filters = {v.Name:v for v in views_with_filters}

# # step 2: select source view/viewTemplate

# selected_src_view = forms.SelectFromList.show(dict_views_with_filters.keys(),
# 								title = 'Select view/viewTempalte',			  
# 								multiselect=False,
# 								button_name='Select view/viewTempalte')

# if not selected_src_view:
# 	forms.alert('khong co filter duoc chon', exitscript = True)
# src_view = dict_views_with_filters[selected_src_view]
# # print('selected view: {}'.format(selected_src_view))
# # step 3 select filters to copy

# filter_ids =src_view.GetFilters()
# filters = [doc.GetElement(f_id) for f_id in filter_ids]
# # filter=[]
# # for f_id in filter_ids:
# # 	f=doc.GetElement(f_id)
# # 	filter.append(f)
# dict_filters = {f.Name:f for f in filters}

# selected_filters = forms.SelectFromList.show(sorted(dict_filters),
# 								title = 'Select to copy',
# 								multiselect=True,
# 								button_name='Select to copy')

# if not selected_filters:
# 	forms.alert('No filter were selected. please try again', exitscript = True)
# filter_to_copy = [dict_filters[f_name] for f_name in selected_filters]
# # print(selected_filters)
# # print(filter_to_copy)

# # step4: select destination view
# dict_all_views = {v.Name:v for v in all_views}
# selected_dest_views = forms.SelectFromList.show(sorted(dict_all_views),
# 								title = 'Select Destinaton views/viewTemplates',			  
# 								multiselect=True,
# 								button_name='Select Destinaton views/viewTemplates')

# if not selected_dest_views:
# 	forms.alert('No Destinaton views/viewTemplates was selected. please try again', exitscript = True)
# dest_views = [dict_all_views[v_name] for v_name in selected_dest_views]
# print(dest_views)

# # step 5 : copy view Filters
# with Transaction(doc,__title__) as t:
# 	t.Start()
# 	for view_filter in filter_to_copy:
# 		filter_overrides = src_view.GetFilterOverrides(view_filter.Id)
# 		for View in dest_views:
# 			View.SetFilterOverrides(view_filter.Id, filter_overrides)
# 	t.Commit()
