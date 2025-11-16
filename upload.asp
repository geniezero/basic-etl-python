<%@ Language=VBScript %>
<%
' ========================================
' API UPLOAD - upload.asp
' ========================================
Option Explicit
Response.Buffer = True
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

' CORS - Si nécessaire pour appeler depuis un autre domaine
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Access-Control-Allow-Methods", "POST, OPTIONS"
Response.AddHeader "Access-Control-Allow-Headers", "Content-Type"

' Gérer les requêtes OPTIONS (preflight CORS)
If Request.ServerVariables("REQUEST_METHOD") = "OPTIONS" Then
    Response.End
End If

' ========================================
' CONFIGURATION
' ========================================
Const UPLOAD_PATH = "C:\uploads\"
Const MAX_FILE_SIZE = 10485760        ' 10 MB
Const ALLOWED_EXTENSIONS = "csv,txt"

' ========================================
' TRAITEMENT DE L'UPLOAD
' ========================================
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim objResult
    Set objResult = ProcessUpload()
    
    ' Retourner la réponse en JSON
    If objResult("success") Then
        Response.Status = "200 OK"
        Response.Write JsonSuccess(objResult)
    Else
        Response.Status = "400 Bad Request"
        Response.Write JsonError(objResult("error"))
    End If
    
    Set objResult = Nothing
Else
    ' Méthode non autorisée
    Response.Status = "405 Method Not Allowed"
    Response.Write JsonError("Seule la méthode POST est autorisée")
End If

' ========================================
' FONCTION PRINCIPALE D'UPLOAD
' ========================================
Function ProcessUpload()
    Dim objDict, objStream, objFSO
    Dim lngBytesCount, strData, strFileName, strFileExt
    Dim strSavePath, strBoundary, intFileStart, intFileEnd
    
    Set objDict = Server.CreateObject("Scripting.Dictionary")
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    
    objDict("success") = False
    
    On Error Resume Next
    
    ' Créer le dossier si nécessaire
    If Not objFSO.FolderExists(UPLOAD_PATH) Then
        objFSO.CreateFolder(UPLOAD_PATH)
    End If
    
    ' Récupérer la taille
    lngBytesCount = Request.TotalBytes
    
    ' Validations
    If lngBytesCount = 0 Then
        objDict("error") = "Aucun fichier sélectionné"
        Set ProcessUpload = objDict
        Exit Function
    End If
    
    If lngBytesCount > MAX_FILE_SIZE Then
        objDict("error") = "Fichier trop volumineux. Maximum: " & FormatBytes(MAX_FILE_SIZE)
        Set ProcessUpload = objDict
        Exit Function
    End If
    
    ' Créer le stream et lire les données binaires
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1
    objStream.Open
    objStream.Write Request.BinaryRead(lngBytesCount)
    
    ' Convertir temporairement en texte pour analyser le multipart/form-data
    objStream.Position = 0
    objStream.Type = 2
    objStream.Charset = "iso-8859-1"
    strData = objStream.ReadText
    
    ' Extraire le nom du fichier
    strFileName = ExtractFileName(strData)
    strFileExt = LCase(objFSO.GetExtensionName(strFileName))
    
    ' Valider l'extension
    If Not IsExtensionAllowed(strFileExt) Then
        objDict("error") = "Extension non autorisée. Formats acceptés: " & ALLOWED_EXTENSIONS
        Set ProcessUpload = objDict
        objStream.Close
        Exit Function
    End If
    
    ' Nettoyer le nom du fichier
    strFileName = SanitizeFileName(strFileName)
    
    ' Trouver où commence et se termine le contenu du fichier
    strBoundary = Left(strData, InStr(strData, vbCrLf) - 1)
    intFileStart = InStr(strData, vbCrLf & vbCrLf) + 4
    intFileEnd = InStr(intFileStart, strData, strBoundary) - 2
    
    ' Créer un nouveau stream pour sauvegarder le fichier
    Dim objFileStream
    Set objFileStream = Server.CreateObject("ADODB.Stream")
    objFileStream.Type = 1
    objFileStream.Open
    
    ' Repositionner le stream en mode binaire
    objStream.Position = 0
    objStream.Type = 1
    objStream.Position = intFileStart - 4
    
    ' Lire uniquement les octets du fichier
    Dim bytFileData
    bytFileData = objStream.Read(intFileEnd - intFileStart + 4)
    objFileStream.Write bytFileData
    
    ' Définir le chemin avec le nom original
    strSavePath = UPLOAD_PATH & strFileName
    
    ' Sauvegarder le fichier
    objFileStream.SaveToFile strSavePath, 2
    objFileStream.Close
    objStream.Close
    
    ' Récupérer les informations
    Dim objFile
    Set objFile = objFSO.GetFile(strSavePath)
    
    objDict("success") = True
    objDict("fileName") = objFSO.GetFileName(strSavePath)
    objDict("filePath") = strSavePath
    objDict("fileSize") = objFile.Size
    objDict("fileSizeFormatted") = FormatBytes(objFile.Size)
    objDict("uploadDate") = Now()
    
    Set objFile = Nothing
    Set objFileStream = Nothing
    Set objStream = Nothing
    Set objFSO = Nothing
    
    Set ProcessUpload = objDict
End Function

' ========================================
' FONCTIONS UTILITAIRES
' ========================================
Function ExtractFileName(strData)
    Dim intPos, strTemp
    intPos = InStr(strData, "filename=""")
    If intPos > 0 Then
        strTemp = Mid(strData, intPos + 10)
        intPos = InStr(strTemp, """")
        ExtractFileName = Left(strTemp, intPos - 1)
        If InStr(ExtractFileName, "\") > 0 Then
            ExtractFileName = Mid(ExtractFileName, InStrRev(ExtractFileName, "\") + 1)
        End If
    Else
        ExtractFileName = "fichier_inconnu.txt"
    End If
End Function

Function SanitizeFileName(strFileName)
    Dim strClean
    strClean = strFileName
    strClean = Replace(strClean, "..", "")
    strClean = Replace(strClean, "/", "")
    strClean = Replace(strClean, "\", "")
    strClean = Replace(strClean, ":", "")
    strClean = Replace(strClean, "*", "")
    strClean = Replace(strClean, "?", "")
    strClean = Replace(strClean, """", "")
    strClean = Replace(strClean, "<", "")
    strClean = Replace(strClean, ">", "")
    strClean = Replace(strClean, "|", "")
    SanitizeFileName = strClean
End Function

Function IsExtensionAllowed(strExt)
    IsExtensionAllowed = (InStr("," & ALLOWED_EXTENSIONS & ",", "," & strExt & ",") > 0)
End Function

Function FormatBytes(lngBytes)
    If lngBytes < 1024 Then
        FormatBytes = lngBytes & " B"
    ElseIf lngBytes < 1048576 Then
        FormatBytes = FormatNumber(lngBytes / 1024, 2) & " KB"
    Else
        FormatBytes = FormatNumber(lngBytes / 1048576, 2) & " MB"
    End If
End Function

' ========================================
' FONCTIONS JSON
' ========================================
Function JsonSuccess(objData)
    Dim strJson
    strJson = "{"
    strJson = strJson & """success"": true,"
    strJson = strJson & """message"": ""Fichier uploadé avec succès"","
    strJson = strJson & """data"": {"
    strJson = strJson & """fileName"": """ & JsonEscape(objData("fileName")) & ""","
    strJson = strJson & """filePath"": """ & JsonEscape(objData("filePath")) & ""","
    strJson = strJson & """fileSize"": " & objData("fileSize") & ","
    strJson = strJson & """fileSizeFormatted"": """ & objData("fileSizeFormatted") & ""","
    strJson = strJson & """uploadDate"": """ & objData("uploadDate") & """"
    strJson = strJson & "}"
    strJson = strJson & "}"
    JsonSuccess = strJson
End Function

Function JsonError(strMessage)
    Dim strJson
    strJson = "{"
    strJson = strJson & """success"": false,"
    strJson = strJson & """error"": """ & JsonEscape(strMessage) & """"
    strJson = strJson & "}"
    JsonError = strJson
End Function

Function JsonEscape(strText)
    Dim strResult
    strResult = strText
    strResult = Replace(strResult, "\", "\\")
    strResult = Replace(strResult, """", "\""")
    strResult = Replace(strResult, vbCrLf, "\n")
    strResult = Replace(strResult, vbCr, "\n")
    strResult = Replace(strResult, vbLf, "\n")
    strResult = Replace(strResult, vbTab, "\t")
    JsonEscape = strResult
End Function
%>