#import win32com.client as win32
#
#list_to_replace = {"C:\\Users\\pcp03\\Desktop\\Nova pasta\\133133.ipt": "T:\\01_Projetos\\Ped_1360_Cooxupe\\11.Silo_Mistura\\133130.idw"}
#
#inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
#
#
##copiar idw para pasta em quetao   salvar com nome novo
#
#for key, value in list_to_replace.items():
#    if key and value:
#        idw = inv.Open(value)
#        idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)
#        inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
#        inv.FileSaveAs.ExecuteSaveAs()
#
import pythoncom

if hasattr(pythoncom, '__file__'):
    print(pythoncom.__file__)
    