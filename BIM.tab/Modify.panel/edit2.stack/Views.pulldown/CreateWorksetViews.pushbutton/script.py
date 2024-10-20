# -*- coding: utf-8 -*-
__title__ = "Create Workset Views"   # Name of the button displayed in Revit


# ╦╔╦╗╔═╗╔═╗╦═╗╔╦╗╔═╗
# ║║║║╠═╝║ ║╠╦╝ ║ ╚═╗
# ╩╩ ╩╩  ╚═╝╩╚═ ╩ ╚═╝ IMPORTS
# ==================================================
import sys
from Autodesk.Revit.DB import  (View3D ,FilteredWorksetCollector, WorksetKind, WorksetVisibility,ViewFamilyType, FilteredElementCollector,
								Transaction, SubTransaction, BuiltInParameter, BuiltInCategory)
# pyRevit
from pyrevit import forms
from rpw.ui.forms import*

# Custom
from Snippets._context_manager import ef_Transaction

# ╦  ╦╔═╗╦═╗╦╔═╗╔╗ ╦  ╔═╗╔═╗
# ╚╗╔╝╠═╣╠╦╝║╠═╣╠╩╗║  ║╣ ╚═╗
#  ╚╝ ╩ ╩╩╚═╩╩ ╩╚═╝╩═╝╚═╝╚═╝ VARIABLES
# ==================================================

doc            = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

try:
	value = TextInput('Nhập tên view', default= "vn_08")
	all_views      = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
	all_view_names = [view.Name for view in all_views]
	view           = None
	# ╔═╗╦ ╦╔╗╔╔═╗╔╦╗╦╔═╗╔╗╔╔═╗
	# ╠╣ ║ ║║║║║   ║ ║║ ║║║║╚═╗
	# ╚  ╚═╝╝╚╝╚═╝ ╩ ╩╚═╝╝╚╝╚═╝ FUNCTIONS
	# ==================================================
	def get_view_type_3D():
		"""Function to get ViewType - 3D View"""
		all_view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
		for view_type in all_view_types:
			if '3D' in view_type.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString():
				return view_type

	# ╔╦╗╔═╗╦╔╗╔
	# ║║║╠═╣║║║║
	# ╩ ╩╩ ╩╩╝╚╝ MAIN
	# ==================================================
	view_type_3D = get_view_type_3D()

	if __name__ == '__main__':
		with ef_Transaction(doc,__title__, debug=True):
			#>>>>>>>>>>>>>>>>>>>> GET WORKSETS
			all_worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets() #ToWorksets #ToWorksets()
			if not all_worksets:
				forms.alert('No Worksets found in the current project.')
				print("No Worksets found in the current project.")
				sys.exit()

			#>>>>>>>>>>>>>>>>>>>> LOOP THROUGH WORKSETS
			print("CREATING WORKSETS: ")
			for workset in all_worksets:
				view_new_name = "3D_Workset_{}_{}".format(workset.Name,value)
				if view_new_name in all_view_names:
					print("--- Workset 3DView already exists: {}".format(view_new_name))
					continue

				# CREATE NEW WORKSET VIEW
				view      = View3D.CreateIsometric(doc, view_type_3D.Id)
				view.Name = view_new_name
				print("--- Workset 3DView created: {}".format(view_new_name))

				#>>>>>>>>>> SET WORKSET VISIBILITIES
				for ws in all_worksets:
					if workset.Name == ws.Name:
						view.SetWorksetVisibility(ws.Id, WorksetVisibility.Visible)
					else:
						view.SetWorksetVisibility(ws.Id, WorksetVisibility.Hidden)

		if view:
			uidoc.ActiveView = view
except:
	pass
