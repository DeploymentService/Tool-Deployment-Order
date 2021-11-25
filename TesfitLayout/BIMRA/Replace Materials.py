import unreal

swap1_material = unreal.EditorAssetLibrary.load_asset("Material'/Engine/Game/V2/Content/V2/2021-ON-002-MEADOWS-ARCH-3DView-_3D_-_anurag_saini_ltl_/Materials/Steel_ASTM_A992'")
if swap1_material is None:
	print "The swap1 material can't be loaded"
	quit()

swap2_material = unreal.EditorAssetLibrary.load_asset("Material'/Engine/Game/V2/Content/StarterContent/Materials/M_Metal_Steel'")
if swap2_material is None:
	print "The swap2 material can't be loaded"
	quit()

# Find all StaticMesh Actor in the Level
actor_list = unreal.EditorLevelLibrary.get_all_level_actors()
actor_list = unreal.EditorFilterLibrary.by_class(actor_list, unreal.StaticMeshActor.static_class())
unreal.EditorLevelLibrary.replace_mesh_components_materials_on_actors(actor_list, swap1_material, swap2_material)
