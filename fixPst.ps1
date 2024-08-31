
Sub ProcesarPSTs()
    Dim objNamespace As Outlook.NameSpace
    Dim objStore As Outlook.Store
    Dim objPSTFile As String
    Dim objFileSystem As Object
    Dim objFolder As Object
    Dim objFile As Object
    Debug.Print "Iniciando ..."
 
    ' Especifica la ruta de la carpeta que contiene los archivos PST
    Dim rutaCarpetaPST As String
    rutaCarpetaPST = "C:\PST_GREEN\pst_sin_arreglar\"
 
    ' Obtiene el objeto Namespace
    Set objNamespace = Application.GetNamespace("MAPI")
 
    ' Obtiene el objeto FileSystem
    Set objFileSystem = CreateObject("Scripting.FileSystemObject")
 
    ' Obtiene el objeto Folder (carpeta) que contiene los archivos PST
    Set objFolder = objFileSystem.GetFolder(rutaCarpetaPST)
 
    ' Itera a trav?s de los archivos PST en la carpeta
    For Each objFile In objFolder.Files
        If LCase(Right(objFile.Name, 4)) = ".pst" Then
            ' Construye la ruta completa del archivo PST
            objPSTFile = rutaCarpetaPST & objFile.Name
 
            ' Adjunta el archivo PST a Outlook
            openPST rutaCarpetaPST, objFile.Name
 
            ' Realiza algunas operaciones con el archivo PST adjunto (puedes agregar m?s aqu?)
            MoveFolder rutaCarpetaPST, objFile.Name
            MoveFolderEnglish rutaCarpetaPST, objFile.Name
            MoveFolderPortug rutaCarpetaPST, objFile.Name
 
            ' Elimina el archivo PST adjunto de Outlook
            closePST rutaCarpetaPST, objFile.Name
        End If
    Next objFile
    ' Itera a trav?s de los archivos PST en la carpeta
    For Each objFile In objFolder.Files
        If LCase(Right(objFile.Name, 4)) = ".pst" Then
            ' Construye la ruta completa del archivo PST
            objPSTFile = rutaCarpetaPST & objFile.Name
            ' Adjunta el archivo PST a Outlook
            openPST rutaCarpetaPST, objFile.Name
        End If
    Next objFile
 
    MsgBox "Proceso completado.", vbInformation
End Sub
 
 
Sub openPST(folder As String, file As String)
    Dim objNamespace As Outlook.NameSpace
    Set objNamespace = Application.GetNamespace("MAPI")
 
    Source = folder & file
    objNamespace.AddStore Source
    Set objPST = objNamespace.Folders.GetLast
    objPST.Name = file
End Sub
Sub closePST(folder As String, file As String)
    Dim objNamespace As Outlook.NameSpace
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objPST = objNamespace.Folders.GetLast
    objPST.Name = file
    objNamespace.RemoveStore objPST
    'source = folder & file
    'objNamespace.RemoveStore source
End Sub
 
 
Sub MoveFolder(folder As String, file As String)
    'Dim objOL As Object
    'Set objOL = CreateObject("Outlook.Application")
    Dim objOL As Outlook.Application
    Set objOL = New Outlook.Application
    Set objNS = objOL.GetNamespace("MAPI")
    'Set objFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set psts = objNS.Folders
  Set objPSTOrigen = objNS.Folders.Item(file)
  On Error GoTo Catch
    Set rootFolders = objPSTOrigen.Folders
    For Each rootFolder In rootFolders
        If InStr(rootFolder.Name, "(Pr") > 0 Then
            'Debug.Print "Find " & rootFolder.Name
            folderPrimary = rootFolder.Name
        End If
    Next
    Debug.Print folderPrimary
    Set objCarpetaRaiz = objPSTOrigen.Folders.Item(folderPrimary)
        foldersToMove = 1
        While foldersToMove > 0
          Debug.Print "Buscando Principio"
          Set objCarpetaOrigen = objCarpetaRaiz.Folders.Item("Principio del almacén de información")
          Set Folders = objCarpetaOrigen.Folders
          foldersToMove = Folders.Count
          For Each subfolder In objCarpetaOrigen.Folders
            subfolder.MoveTo objPSTOrigen
            Debug.Print "Moviendo " & subfolder
          Next
        Wend
    objCarpetaRaiz.Name = "Trash"
 
    'MsgBox "OK"
Continue:
    Exit Sub
Catch:
    Debug.Print "Not found Principio"
    Exit Sub
 
End Sub
 
Sub MoveFolderEnglish(folder As String, file As String)
    'Dim objOL As Object
    'Set objOL = CreateObject("Outlook.Application")
    Dim objOL As Outlook.Application
    Set objOL = New Outlook.Application
    Set objNS = objOL.GetNamespace("MAPI")
    'Set objFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set psts = objNS.Folders
  Set objPSTOrigen = objNS.Folders.Item(file)
  On Error GoTo Catch
    Set rootFolders = objPSTOrigen.Folders
    For Each rootFolder In rootFolders
        If InStr(rootFolder.Name, "(Pr") > 0 Then
            'Debug.Print "Find " & rootFolder.Name
            folderPrimary = rootFolder.Name
        End If
    Next
    Debug.Print folderPrimary
    Set objCarpetaRaiz = objPSTOrigen.Folders.Item(folderPrimary)
        foldersToMove = 1
        While foldersToMove > 0
          Debug.Print "Buscando Top"
          Set objCarpetaOrigen = objCarpetaRaiz.Folders.Item("Top of Information Store")
          Set Folders = objCarpetaOrigen.Folders
          foldersToMove = Folders.Count
          For Each subfolder In objCarpetaOrigen.Folders
            subfolder.MoveTo objPSTOrigen
            Debug.Print "Moviendo " & subfolder
          Next
        Wend
    objCarpetaRaiz.Name = "Trash"
 
    'MsgBox "OK"
Continue:
    Exit Sub
Catch:
    Debug.Print "Not found Top"
    Exit Sub
 
End Sub

Sub MoveFolderPortug(folder As String, file As String)
    'Dim objOL As Object
    'Set objOL = CreateObject("Outlook.Application")
    Dim objOL As Outlook.Application
    Set objOL = New Outlook.Application
    Set objNS = objOL.GetNamespace("MAPI")
    'Set objFolder = objNS.GetDefaultFolder(olFolderContacts)
    Set psts = objNS.Folders
  Set objPSTOrigen = objNS.Folders.Item(file)
  On Error GoTo Catch
    Set rootFolders = objPSTOrigen.Folders
    For Each rootFolder In rootFolders
        If InStr(rootFolder.Name, "(Pr") > 0 Then
            'Debug.Print "Find " & rootFolder.Name
            folderPrimary = rootFolder.Name
        End If
    Next
    Debug.Print folderPrimary
    Set objCarpetaRaiz = objPSTOrigen.Folders.Item(folderPrimary)
        foldersToMove = 1
        While foldersToMove > 0
          Debug.Print "Buscando Topinho"
          Set objCarpetaOrigen = objCarpetaRaiz.Folders.Item("Início do Repositório de Informações")
          Set Folders = objCarpetaOrigen.Folders
          foldersToMove = Folders.Count
          For Each subfolder In objCarpetaOrigen.Folders
            subfolder.MoveTo objPSTOrigen
            Debug.Print "Moviendihnoo " & subfolder
          Next
        Wend
    objCarpetaRaiz.Name = "Trash"
 
    'MsgBox "OK"
Continue:
    Exit Sub
Catch:
    Debug.Print "Not found carpetinha"
    Exit Sub
 
End Sub

