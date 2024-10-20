# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import *
from System.Collections.Generic import *
uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
active_view = doc.ActiveView 

try:
	all_sport_elevation = FilteredElementCollector(doc, active_view.Id).OfCategory(BuiltInCategory.OST_SpotElevations).WhereElementIsNotElementType()
	from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, TextBox,
								Separator, Button, CheckBox)
	components = [Label('CHỌN SPORT ELEVATION:'),
					Label('Chọn kiểu cao độ:'),
					ComboBox('combobox1', ["=0",">0","<0"]),
					Separator(),
					Button('Run')]
	form = FlexForm('LNQ tools', components)
	form.show()
	form.values
	kieu_sport= (form.values["combobox1"])

	def filter_sport_elevation_by_kieu_sport(kieu_sport):
		selected_ids = List[ElementId]()
		
		for i in all_sport_elevation:
			elementtype = doc.GetElement(i.GetTypeId())
			caodo_1 = i.get_Parameter(BuiltInParameter.SPOT_ELEV_SINGLE_OR_UPPER_VALUE).AsValueString()
			caodo = float(caodo_1)
			ten_type = Element.Name.GetValue(elementtype)
			
			if kieu_sport == "=0" and caodo == 0:
				selected_ids.Add(i.Id)
			elif kieu_sport == ">0" and caodo > 0:
				selected_ids.Add(i.Id)
			else: 
				if kieu_sport == "<0" and caodo < 0:
					selected_ids.Add(i.Id)

		return selected_ids

	kieu_sport = form.values["combobox1"]
	selected_ids = filter_sport_elevation_by_kieu_sport(kieu_sport)
	uidoc.Selection.SetElementIds(selected_ids)
except: 
	pass


















# selected_ids =  List[ElementId]()
# for i in all_sport_elevation:
# 	elementtype = doc.GetElement(i.GetTypeId())
# 	# print(i.Category.Name)
# 	caodo_1= i.get_Parameter (BuiltInParameter.SPOT_ELEV_SINGLE_OR_UPPER_VALUE).AsValueString()
# 	caodo = int(caodo_1)
# 	# print(type(caodo))
# 	ten_type = Element.Name.GetValue(elementtype)
#  	# print(ten_type)
# 	if caodo == 0:
# 		selected_ids.Add(i.Id)
# uidoc.Selection.SetElementIds(selected_ids)



# Chọn tất cả các đối tượng Spot Elevation
# selected_ids = []
# for spot_elevation in spot_elevations:
# 	selected_ids.append(spot_elevation.Id)

# uidoc.Selection.SetElementIds(selected_ids)






# selection = uidoc.Selection.GetElementIds()
# list_chon = []
# for i in selection:
# 	ele = doc.GetElement(i)
# 	elementtype = doc.GetElement(ele.GetTypeId())
# 	category = ele.Category.Name
# 	print(category)
# 	caodo = ele.get_Parameter (BuiltInParameter.SPOT_ELEV_SINGLE_OR_UPPER_VALUE)
# 	ten_type = Element.Name.GetValue(elementtype)
# 	print(caodo.AsValueString())
# 	print(ten_type)
# 	if caodo.AsDouble() == 0:
# 		list_chon.append(ele)
# 		uidoc.Selection.SetElementIds(list_chon)