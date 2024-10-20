
# -*- coding: utf-8 -*-
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.DB import*

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document
app = __revit__.Application

# lấy tên legend

t = Transaction(doc, 'Replace Door Accessory')
t.Start()
selection = uidoc.Selection.GetElementIds()

for i in selection:
	ele = doc.GetElement(i)
	elementtype = doc.GetElement(ele.GetTypeId())
	category = ele.Category.Name
	if category == "Generic Annotations" :
		height = ele.LookupParameter("Height")#.AsDouble()*304.8
		width = ele.LookupParameter("Width")#.AsDouble()*304.8
		type_mark = ele.LookupParameter("Type Mark")#.AsString()
		vi_tri_cua = ele.LookupParameter("Vị Trí Cửa").AsString()
		type_comments = ele.LookupParameter("Type Comments")#.AsString()
		phu_kien = ele.LookupParameter("Phụ Kiện Cửa")#.AsString()
		description = ele.LookupParameter("Description")#.AsString()

		# print(height)
		# print(type_mark)
		# print(type_comments)
		# print(vi_tri_cua)
		# print(type(phu_kien))
		# print(description)
	elif category == "Legend Components" :
		component_type = ele.LookupParameter("Component Type").AsElementId()
		hostcua= doc.GetElement(component_type)
		tentype = hostcua.get_Parameter (BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString() # lấy ra tên ten type của cửa của mà legend hiển thị
		# print(type(tentype))

	else: 
		print("chọn sai đối tượng")

cua = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Doors).WhereElementIsElementType() # Lấy Ra toàn bộ cửa trong dự án

for i in cua:
	tentype_all = i.get_Parameter (BuiltInParameter.SYMBOL_NAME_PARAM).AsValueString() # lấy ra tên tên type của cửa trong tất cả dự án
	if tentype_all == tentype:
		height_cua = i.get_Parameter (BuiltInParameter.DOOR_HEIGHT).AsValueString()
		description_cua = i.get_Parameter (BuiltInParameter.ALL_MODEL_DESCRIPTION).AsValueString()
		type_mark_cua = i.get_Parameter (BuiltInParameter.WINDOW_TYPE_ID).AsValueString()
		type_comments_cua = i.get_Parameter (BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsValueString()
		phu_kien_cua = i.LookupParameter("Phụ Kiện Cửa").AsString()
		width_cua = i.get_Parameter (BuiltInParameter.FURNITURE_WIDTH).AsValueString()


# Đổi gia trị để set

		width_cua_set = int(width_cua)
		height_cua_set = int(height_cua)
		phu_kien_cua_set = str(phu_kien_cua)
		type_comments_cua_set = str(type_comments_cua)
		description_cua_set = str(description_cua)
		type_mark_cua_set = str(type_mark_cua)

# Set giá trị mới

		width.Set(width_cua_set/304.8)
		height.Set(height_cua_set/304.8)
		phu_kien.Set(phu_kien_cua_set)
		type_comments.Set(type_comments_cua_set)
		description.Set(description_cua_set)
		type_mark.Set(type_mark_cua_set)
print("_______________ĐÃ HOÀN THÀNH_______________")


t.Commit()


# print(tentype_all)
# print(float(height_cua))
# print(description_cua)
# print(type_mark_cua)
# print(type_comments_cua)
# print(phu_kien_cua_set.AsValueString())



# for i in selection:
# 	ele = doc.GetElement(i)
# 	if isinstance(ele, ViewSheet):
# 		sheet_name = ele.LookupParameter("Sheet Name").AsString()
# 		sheet_number = ele.LookupParameter("Sheet Number").AsString()
# 		sheet_number_para = ele.LookupParameter("Sheet Number")
# 		# print(sheet_number)
# 		list_sheet.append((ele, sheet_number))  # Lưu trữ cả ViewSheet và giá trị Sheet Number tương ứng
# if kieu_seri == 1:
# 	count = so_bat_dau
# 	sorted_list_sheet = sorted(list_sheet, key=lambda x: x[1])  # Sắp xếp danh sách theo thứ tự tăng dần của Sheet Number
# 	for sheet in sorted_list_sheet:
# 		sheet_number_formatted = str(count).zfill(1)  # Định dạng số đếm thành chuỗi có độ dài 3 ký tự, bằng cách thêm số 0 vào đầu nếu cần
# 		print(sheet_number_formatted, "-", sheet[1])
# 		sheet_number_new = str(ky_tu_ban_ve + str(count).zfill(1) + Hau_to)
# 		sheet[0].LookupParameter("Sheet Number").Set(sheet_number_new)  # Gán giá trị mới cho Sheet Number của ViewSheet
# 		count += 1

