import win32com.client as win32
import os

list_to_replace = {r"T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133398.ipt": r"\\alfa\Eng\Ped_0005_Cooxupe\08.Silo_Metalico_23854\26927.idw"}

inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
oPropSets = inv.PropertySets
PropertySet = oPropSets.Item("Design Tracking Properties")

for key, value in list_to_replace.items():
    if os.path.exists(key) and os.path.exists(value):
        idw = inv.Open(r'T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133205.ipt')
        #idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)
        #inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
        #inv.FileSaveAs.ExecuteSaveAs()
        print(PropertySet.Item(2).Value)
    else:
        print('ops')


try:
    idw = inv.Open(r'T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133398.ipt')
except Exception as e:
    print('num deu', e)
