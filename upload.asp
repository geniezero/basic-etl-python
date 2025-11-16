<%@ Language=VBScript %>
<%
' ========================================
' API UPLOAD - upload.asp
' Version finale - Transfert binaire sans modification
' ========================================
Option Explicit
Response.Buffer = True
Response.Charset = "UTF-8"
Response.ContentType = "application/json"

' ========================================
' CONFIGURATION DYNAMIQUE SELON SERVER_NAME
' ========================================
Dim UPLOAD_PATH, MAX_FILE_SIZE, ALLOWED_EXTENSIONS
Dim strServerName

' Récupérer le nom du serveur
strServerName = LCase(Request.ServerVariables("SERVER_NAME"))

' Configurer selon le nom du serveur
If InStr(strServerName, "prod") > 0 Then
    ' Configuration PRODUCTION
    UPLOAD_PATH = "E:\production\uploads\"
    MAX_FILE_SIZE = 104857600  ' 100 MB
Else
    ' Configuration TEST/DEV (par défaut)
    UPLOAD_PATH = "C:\uploads\"
    MAX_FILE_SIZE = 104857600  ' 100 MB
End If

' Extensions autorisées
ALLOWED_EXTENSIONS = "csv,txt,asp,zip,tar,gz,gzip,tgz"

' ========================================
' CORS - Cross-Origin Resource Sharing
' ========================================
Response.AddHeader "Access-Control-Allow-Origin", "*"
Response.AddHeader "Access-Control-Allow-Methods", "POST, OPTIONS"
Response.AddHeader "Access-Control-Allow-Headers", "Content-Type"

' Gérer les requêtes OPTIONS (preflight CORS)
If Request.ServerVariables("REQUEST_METHOD") = "OPTIONS" Then
    Response.End
End If

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
    Dim strUploadPath, lngFileSize
    
    Set objDict = Server.CreateObject("Scripting.Dictionary")
    Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
    
    objDict("success") = False
    
    On Error Resume Next
    
    ' Utiliser le chemin configuré
    strUploadPath = UPLOAD_PATH
    
    ' Créer le dossier si nécessaire
    If Not objFSO.FolderExists(strUploadPath) Then
        objFSO.CreateFolder(strUploadPath)
    End If
    
    ' Récupérer la taille totale de la requête
    lngBytesCount = Request.TotalBytes
    
    ' === VALIDATIONS ===
    
    ' Vérifier qu'un fichier est présent
    If lngBytesCount = 0 Then
        objDict("error") = "Aucun fichier sélectionné"
        Set ProcessUpload = objDict
        Exit Function
    End If
    
    ' Vérifier la taille maximale
    If lngBytesCount > MAX_FILE_SIZE Then
        objDict("error") = "Fichier trop volumineux. Maximum: " & FormatBytes(MAX_FILE_SIZE)
        Set ProcessUpload = objDict
        Exit Function
    End If
    
    ' === LECTURE DES DONNÉES ===
    
    ' Créer le stream et lire les données binaires
    Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Type = 1 ' adTypeBinary
    objStream.Open
    objStream.Write Request.BinaryRead(lngBytesCount)
    
    ' Convertir temporairement en texte pour analyser le multipart/form-data
    objStream.Position = 0
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "iso-8859-1"
    strData = objStream.ReadText
    
    ' === EXTRACTION DU NOM DE FICHIER ===
    
    strFileName = ExtractFileName(strData)
    strFileExt = LCase(objFSO.GetExtensionName(strFileName))
    
    ' Valider l'extension
    If Not IsExtensionAllowed(strFileExt) Then
        objDict("error") = "Extension non autorisée. Formats acceptés: " & ALLOWED_EXTENSIONS
        Set ProcessUpload = objDict
        objStream.Close
        Exit Function
    End If
    
    ' Nettoyer le nom du fichier (sécurité)
    strFileName = SanitizeFileName(strFileName)
    
    ' === EXTRACTION DU CONTENU BINAIRE DU FICHIER ===
    
    ' Trouver le boundary de début
    strBoundary = Left(strData, InStr(strData, vbCrLf) - 1)
    
    ' Trouver le début des headers HTTP (Content-Disposition)
    intFileStart = InStr(strData, "Content-Disposition")
    
    ' Trouver la fin des headers (double CRLF = début du fichier)
    intFileStart = InStr(intFileStart, strData, vbCrLf & vbCrLf) + 4
    
    ' Trouver le boundary de fin (marque la fin du fichier)
    intFileEnd = InStr(intFileStart, strData, vbCrLf & strBoundary) - 1
    
    ' Calculer la taille exacte du fichier (en caractères dans la string)
    lngFileSize = intFileEnd - intFileStart + 1
    
    ' === SAUVEGARDE DU FICHIER ===
    
    ' Créer un nouveau stream pour sauvegarder le fichier
    Dim objFileStream
    Set objFileStream = Server.CreateObject("ADODB.Stream")
    objFileStream.Type = 1 ' adTypeBinary
    objFileStream.Open
    
    ' Repositionner le stream original en mode binaire
    objStream.Position = 0
    objStream.Type = 1 ' adTypeBinary
    
    ' Se positionner au début du fichier (après les headers HTTP)
    Dim lngBinaryPos
    lngBinaryPos = intFileStart - 1
    objStream.Position = lngBinaryPos
    
    ' Lire EXACTEMENT les octets du fichier (sans headers ni boundaries)
    Dim bytFileData
    bytFileData = objStream.Read(lngFileSize)
    objFileStream.Write bytFileData
    
    ' Définir le chemin complet de sauvegarde
    strSavePath = strUploadPath & strFileName
    
    ' Sauvegarder le fichier sur le disque
    objFileStream.SaveToFile strSavePath, 2 ' adSaveCreateOverWrite
    objFileStream.Close
    objStream.Close
    
    ' === INFORMATIONS DE RETOUR ===
    
    ' Récupérer les informations du fichier sauvegardé
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
' FONCTION: Extraire le nom du fichier
' ========================================
Function ExtractFileName(strData)
    Dim intPos, strTemp
    
    ' Chercher filename=" dans les headers
    intPos = InStr(strData, "filename=""")
    
    If intPos > 0 Then
        ' Extraire le nom entre les guillemets
        strTemp = Mid(strData, intPos + 10)
        intPos = InStr(strTemp, """")
        ExtractFileName = Left(strTemp, intPos - 1)
        
        ' Supprimer le chemin complet si présent (ancien navigateurs)
        If InStr(ExtractFileName, "\") > 0 Then
            ExtractFileName = Mid(ExtractFileName, InStrRev(ExtractFileName, "\") + 1)
        End If
    Else
        ' Nom par défaut si non trouvé
        ExtractFileName = "fichier_inconnu.txt"
    End If
End Function

' ========================================
' FONCTION: Nettoyer le nom de fichier
' ========================================
Function SanitizeFileName(strFileName)
    Dim strClean
    strClean = strFileName
    
    ' Supprimer les caractères dangereux pour le système de fichiers
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

' ========================================
' FONCTION: Vérifier l'extension
' ========================================
Function IsExtensionAllowed(strExt)
    ' Vérifie si l'extension est dans la liste autorisée
    IsExtensionAllowed = (InStr("," & ALLOWED_EXTENSIONS & ",", "," & strExt & ",") > 0)
End Function

' ========================================
' FONCTION: Formater les octets
' ========================================
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
' FONCTION: Réponse JSON succès
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

' ========================================
' FONCTION: Réponse JSON erreur
' ========================================
Function JsonError(strMessage)
    Dim strJson
    strJson = "{"
    strJson = strJson & """success"": false,"
    strJson = strJson & """error"": """ & JsonEscape(strMessage) & """"
    strJson = strJson & "}"
    JsonError = strJson
End Function

' ========================================
' FONCTION: Échapper les caractères JSON
' ========================================
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
