# -*- coding: utf-8 -*-
import sys
from Autodesk.Revit.DB import *
from pyrevit.forms import select_views,select_sheets
from pyrevit import forms
from System.Windows.Forms import MessageBox
from pyrevit import revit, DB, UI
__doc__ = 'LNQ'
__title__ = 'Copy 2D on Sheet'

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
view = doc.ActiveView

# try:
transform = Transform.Identity
opts = CopyPasteOptions()

elementToCopy = uidoc.Selection.GetElementIds()
list_category = []
for j in elementToCopy:
	element_copy = doc.GetElement(j)
	category = (element_copy.Category).Name
	list_category.append(category)
	if category == "Viewports":
		type_element_copy = (element_copy.ViewId).IntegerValue
	else:
		type_element_copy = (element_copy.GetTypeId()).IntegerValue
	

selected_sheets = select_sheets(__title__,multiple=True,
						 button_name='Select Sheets')
if not selected_sheets:
	sys.exit()
t = Transaction(doc, "copy element")
t.Start()
cursheet = revit.uidoc.ActiveGraphicalView
for v in selected_sheets:
	if cursheet.Id == v.Id:
		selected_sheets.remove(v)
for x in list_category: 
	if x == "Viewports":
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Viewports).OwnedByView(sht.Id).ToElements()

			list_type_element=[]
			for k in existing_vps:
				type_element = (k.ViewId).IntegerValue
				list_type_element.append(type_element)

			if type_element_copy not in list_type_element:
				print("tạo mới")
				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
			elif type_element_copy  in list_type_element:
				print("di chuyển")
				existing_element = next((e for e in existing_vps if ((e.ViewId).IntegerValue == type_element_copy)), None) 
				if existing_element:
					doc.Delete(existing_element.Id)
					ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
	elif x == "Title Blocks":
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OwnedByView(sht.Id).ToElements()

			list_type_element=[]
			for k in existing_vps:
				type_element = (k.GetTypeId()).IntegerValue
				list_type_element.append(type_element)
			if type_element_copy not in list_type_element:
				print("tạo mới")
				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
			elif type_element_copy  in list_type_element:
				print("di chuyển")
				existing_element = next((e for e in existing_vps if (e.GetTypeId().IntegerValue == type_element_copy)), None) 
				if existing_element:
					doc.Delete(existing_element.Id)
					ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
	elif x == "Generic Annotations":
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_GenericAnnotation).OwnedByView(sht.Id).ToElements()

			list_type_element=[]
			for k in existing_vps:
				type_element = (k.GetTypeId()).IntegerValue
				list_type_element.append(type_element)
			if type_element_copy not in list_type_element:
				print("tạo mới")
				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
			elif type_element_copy  in list_type_element:
				print("di chuyển")
				existing_element = next((e for e in existing_vps if (e.GetTypeId().IntegerValue == type_element_copy)), None) 
				if existing_element:
					doc.Delete(existing_element.Id)
					ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
	elif x == "Lines":
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Lines).OwnedByView(sht.Id).ToElements()

			list_type_element=[]
			for k in existing_vps:
				type_element = (k.GetTypeId()).IntegerValue
				list_type_element.append(type_element)
			if type_element_copy not in list_type_element:
				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
	elif x == "Text Notes":
		for sht in selected_sheets:
			existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TextNotes).OwnedByView(sht.Id).ToElements()

			list_type_element=[]
			for k in existing_vps:
				type_element = (k.GetTypeId()).IntegerValue
				list_type_element.append(type_element)
			if type_element_copy not in list_type_element:
				print("tạo mới")
				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
			elif type_element_copy  in list_type_element:
				print("di chuyển")
				existing_element = next((e for e in existing_vps if (e.GetTypeId().IntegerValue == type_element_copy)), None) 
				if existing_element:
					doc.Delete(existing_element.Id)
					ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
	else:
		MessageBox.Show("Đối tượng " +category.upper()+ " không được hỗ trợ trong Tool này", "LNQ-Tools")

t.Commit()
# except:
# 	pass
































# elementToCopy = uidoc.Selection.GetElementIds()
# list_element_copy = []
# title_block_locations = {}
# for j in elementToCopy:
# 	element_copy = doc.GetElement(j)
# 	type_element_copy = (element_copy.GetTypeId()).IntegerValue
# 	list_element_copy .append(type_element_copy)
# 	title_block_locations[type_element_copy] = element_copy.Location.Point
# 	# print(list_element_copy)

# # all_title_block = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElementIds()
# # for i in all_title_block:
# # 	ele = doc.GetElement(i)
# # 	typeid = (ele.GetTypeId()).IntegerValue
# 	# print(typeid)

# # t = Transaction(doc, __title__)
# # t.Start()
# # src_view = doc.ActiveView
# dest_view = select_sheets(__title__,multiple=True,
# 						 button_name='Select Sheets')
# for sht in dest_view:
# 	existing_vps = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).OwnedByView(sht.Id).ToElements()
# 	# existing_schedules = [x for x in allSheetedSchedules
# 	# 						if x.OwnerViewId == sht.Id]
# 	list_type_element=[]
# 	for k in existing_vps:
# 		type_element = (k.GetTypeId()).IntegerValue
# 		list_type_element.append(type_element)
# 	# print(type_element_copy)
# 		# print(list_type_element)

# # 	if i:
# 		for x in list_element_copy:
# 			if x in list_type_element:
# 				print("lần 1")
# 				# t1 = Transaction(doc, "lan 1")
# 				# t1.Start()
# 				# k.Location.Move(title_block_locations[type_element])
# 				# t1.Commit()
# 				# break
# 			else:
# 			# elif x not in list_type_element:
# 				t2 = Transaction(doc, "lan 2")
# 				t2.Start()
# 				transform = Transform.Identity
# 				opts = CopyPasteOptions()
# 				ElementTransformUtils.CopyElements(view, elementToCopy, sht, transform, opts)
# 				t2.Commit()
# 				print("lần 2")




# selected_sheets = forms.select_sheets(title='Select Target Sheets',
#                                       include_placeholder=False,
#                                       button_name='Select Sheets')



# for sheet in selected_sheets:
#     print(sheet.Name)
# selection = uidoc.Selection.GetElementIds() #Trả về 1 list ElementId
# for i in selection:
# 	print(i)
# 	ele = doc.GetElement(i)
# 	category = ele.Category
# 	print(category.Name)
# wallsToCopy = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElementIds()
# # print(all_walls)

# vector = XYZ(25, 50, 0)


# t = Transaction(doc, __title__)
# t.Start()

# ElementTransformUtils.CopyElements(doc, wallsToCopy, vector)


# t.Commit()









# logger = script.get_logger()


# def is_placable(view):
#     if view and hasattr(view, 'ViewType') and view.ViewType in [DB.ViewType.Schedule,
#                                                                DB.ViewType.DraftingView,
#                                                                DB.ViewType.Legend,
#                                                                DB.ViewType.CostReport,
#                                                                # DB.ViewType.TitleBlocks,
#                                                                DB.ViewType.LoadsReport,
#                                                                DB.ViewType.ColumnSchedule,
#                                                                DB.ViewType.PanelSchedule]:
#         return True
#     return False


# def update_if_placed(vport, exst_vps):
#     for exst_vp in exst_vps:
#         if vport.ViewId == exst_vp.ViewId:
#             exst_vp.SetBoxCenter(vport.GetBoxCenter())
#             exst_vp.ChangeTypeId(vport.GetTypeId())
#             return True
#     return False


# selViewports = []

# allSheetedSchedules = DB.FilteredElementCollector(revit.doc)\
#                         .OfClass(DB.ScheduleSheetInstance)\
#                         .ToElements()
# # print(allSheetedSchedules)


# selected_sheets = forms.select_sheets(title='Select Target Sheets',
#                                       include_placeholder=False,
#                                       button_name='Select Sheets')

# # # get a list of viewports to be copied, updated
# if selected_sheets and len(selected_sheets) > 0:
#     if int(__revit__.Application.VersionNumber) > 2014:
#         cursheet = revit.uidoc.ActiveGraphicalView
#         for v in selected_sheets:
#             if cursheet.Id == v.Id:
#                 selected_sheets.remove(v)
#     else:
#         cursheet = selected_sheets[0]
#         selected_sheets.remove(cursheet)

#     revit.uidoc.ActiveView = cursheet
#     selected_vps = revit.pick_elements()

#     if selected_vps:
#         with revit.Transaction('Copy Viewports to Sheets'):
#             for sht in selected_sheets:
#                 existing_vps = [revit.doc.GetElement(x)
#                                 for x in sht.GetAllViewports()]
#                 existing_schedules = [x for x in allSheetedSchedules
#                                       if x.OwnerViewId == sht.Id]
#                 for vp in selected_vps:
#                     if isinstance(vp, DB.FamilyInstance):
#                         src_view = revit.doc.GetElement(vp.Id)
#                         print(src_view)
#                         # check if viewport already exists
#                         # and update location and type
#                         if update_if_placed(vp, existing_vps):
#                             break
#                         # if not, create a new viewport
#                         elif is_placable(src_view):
#                             new_vp = \
#                                 DB.Viewport.Create(revit.doc,
#                                                    sht.Id,
#                                                    vp.Id,
#                                                    vp.GetBoxCenter())

#                             new_vp.ChangeTypeId(vp.GetTypeId())
#                         else:
#                             logger.warning('Skipping %s. This view type '
#                                            'can not be placed on '
#                                            'multiple sheets.',
#                                            revit.query.get_name(src_view))
#                     elif isinstance(vp, DB.ScheduleSheetInstance):
#                         # check if schedule already exists
#                         # and update location
#                         for exist_sched in existing_schedules:
#                             if vp.ScheduleId == exist_sched.ScheduleId:
#                                 exist_sched.Point = vp.Point
#                                 break
#                         # if not, place the schedule
#                         else:
#                             DB.ScheduleSheetInstance.Create(revit.doc,
#                                                             sht.Id,
#                                                             vp.ScheduleId,
#                                                             vp.Point)
#     else:
#         forms.alert('At least one viewport must be selected.')
# else:
#     forms.alert('At least one sheet must be selected.')