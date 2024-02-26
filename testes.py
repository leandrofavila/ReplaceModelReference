import win32com.client as win32
import os
import codecs

#list_to_replace = {r"T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133398.ipt": r"\\alfa\Eng\Ped_0005_Cooxupe\08.Silo_Metalico_23854\26927.idw"}
#
#inv = win32.gencache.EnsureDispatch("Inventor.ApprenticeServer")
#oPropSets = inv.PropertySets
#PropertySet = oPropSets.Item("Design Tracking Properties")
#
#for key, value in list_to_replace.items():
#    if os.path.exists(key) and os.path.exists(value):
#        idw = inv.Open(r'T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133205.ipt')
#        #idw.ReferencedDocumentDescriptors(1).ReferencedFileDescriptor.ReplaceReference(key)
#        #inv.FileSaveAs.AddFileToSave(idw, (key[:-3] + "idw"))
#        #inv.FileSaveAs.ExecuteSaveAs()
#        print(PropertySet.Item(2).Value)
#    else:
#        print('ops')
#
#
#try:
#    idw = inv.Open(r'T:\01_Projetos\Ped_1360_Cooxupe\11.Silo_Mistura\133398.ipt')
#except Exception as e:
#    print('num deu', e)
#
#
#
#path = r'C:\Users\pcp03\PycharmProjects\ReplaceModelReference\tmp\133212.ipt'
#file = os.path.basename(path)
#
#caminhos = set(codecs.open(r"R:\Rubens\30_Prog\Caminhos.txt", "r", "utf-8").readlines())
#tapa = ''
#
#for line in caminhos:
#    if ("\\" + (str(file))) in line:
#        tapa = line.rstrip()
#        break
#
#print('ta procurando', file, 'retornou', tapa)

import win32com.client

# Crie uma instância do objeto COM usando o executável
#objeto = win32com.client.Dispatch("C:\PROGRA~1\SigmaTEK\SIGMAN~1.3SP\SIGMAN~1.EXE")

# Agora você pode usar métodos e propriedades do objeto como de costume
#resultado = objeto.LoadWorkspaceFile(r"Y:\Cnc\Puncionadeira_Cnc\1158042.ws")
#print(resultado)

# Não se esqueça de liberar o objeto depois de terminar de usá-lo
#objeto = None


#C:\PROGRA~1\SigmaTEK\SIGMAN~1.3SP\SIGMAN~1.EXE
import winreg

# Abrir a chave do registro onde a informação do executável está armazenada
with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\CLSID\{B912AA00-902A-11D2-B1DD-0060978DE86F}\LocalServer32") as key:
    # Obter o valor do registro que contém o caminho do executável
    path, _ = winreg.QueryValueEx(key, "")
    print(path)
