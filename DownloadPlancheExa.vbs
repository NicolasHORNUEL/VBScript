myProcess "C:\Users\Administrateur\Google Drive\EXA", "C:\Users\Administrateur\Desktop", "D:\DropFolder\EXA\70x100 2UP"



Sub myProcess(myGoogleExaFolder,myTempFolder,myDropFolder)
	Dim oFSO,oFld 
	Dim FolderName,FolderPath,FilePathZIP,NameXMLFile,FilePathXML
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	If oFSO.FolderExists(myGoogleExaFolder) Then 
    For Each oFld in oFSO.GetFolder(myGoogleExaFolder).SubFolders
      FolderName = oFld.Name
      oFSO.CreateFolder(myTempFolder & "\" & FolderName)
      FilePathZIP = myTempFolder & "\" & FolderName & "\" & FolderName & ".zip"
      FileNameXML = FolderName & ".xml"
      FilePathXML = myTempFolder & "\" & FolderName & "\PEOPLE_Offset_" & FolderName & "\" & FolderName & ".xml"
      If oFSO.FileExists(FilePathZIP) Then
        UnZip myTempFolder & "\" & FolderName, FilePathZIP
        SendMail FileNameXML, FilePathXML
        ReadRenameMoveDeleteFile FilePathXML,myGoogleExaFolder,FolderName,myTempFolder,myDropFolder
      Else
        Download FolderName, FilePathZIP
        UnZip myTempFolder & "\" & FolderName, FilePathZIP
        SendMail FileNameXML, FilePathXML
        ReadRenameMoveDeleteFile FilePathXML,myGoogleExaFolder,FolderName,myTempFolder,myDropFolder
      End If
    Next 
	End If
  Set oFSO = Nothing
End Sub



Sub ReadRenameMoveDeleteFile(FilePathXML,myGoogleExaFolder,FolderName,myTempFolder,myDropFolder)
  Dim xmlDoc,objFSO
  Dim Substrat,Quantite,Grammage
  Dim FilePathPDF,FilePathNewPDF
  'LECTURE DU FICHIER XML https://docs.microsoft.com/en-us/previous-versions/windows/desktop/ms757828(v%3Dvs.85)
  Set xmlDoc = CreateObject("Microsoft.XMLDOM")
  xmlDoc.Async = "false"
  xmlDoc.Load(FilePathXML)
  Set Substrat = xmlDoc.getElementsByTagName("substrat")
  Quantite = Substrat(0).getAttribute("quantite")
  Grammage = Substrat(0).getAttribute("support")
  Grammage = Replace(Grammage, "g ", "")
  Set xmlDoc = Nothing
  'Renomme et D�place le fichier d'impression
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  FilePathPDF = myTempFolder & "\" & FolderName & "\PEOPLE_Offset_" & FolderName & "\" & FolderName & ".pdf"
  FilePathNewPDF = myDropFolder & "\" & FolderName & " " & Quantite & "f " & Grammage & ".pdf"
  objFSO.MoveFile FilePathPDF , FilePathNewPDF
  'SUPPRIMER FICHIER ET DOSSIER
  objFSO.DeleteFolder(myGoogleExaFolder & "\" & FolderName)
  Set objFSO = Nothing
End Sub



Sub Download(FolderName, FilePath)
  ' https://docs.microsoft.com/en-us/windows/win32/winhttp/winhttprequest
  ' https://www.w3schools.com/asp/ado_ref_stream.asp 
	Dim Url01, authorization, Url02, direct_link, Url03, QueryString
	Dim FileNum
	Dim WHTTP
	Dim FileData	
  ' 01 Post LOGIN TO GET AUTHORIZATION RESPONSE
	Set WHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
	Url01 = "https://api-ftps3.exa.io/login"
	WHTTP.Open "POST", Url01, False
	WHTTP.SetRequestHeader "x-api-key", "*************************************" 
	WHTTP.Send "{""username"":""people"",""password"":""******""}"
	authorization = WHTTP.ResponseText
	authorization = Mid(authorization, 2, 239)
	' 02 Get DIRECT LINK To AMAZON SERVER With AUTHORIZATION
	Url02 = "https://api-ftps3.exa.io/file?key=PEOPLE/PEOPLE%20Offset/PEOPLE_Offset_" & FolderName & ".zip"
	WHTTP.Open "GET", Url02, False
	WHTTP.SetRequestHeader "authorization", authorization
	WHTTP.SetRequestHeader "x-api-key", "*************************************"
	WHTTP.Send
	direct_link = WHTTP.ResponseText
  Url03 =  Mid(direct_link, 2, 106)
  QueryString = Mid(direct_link, 109, Len(direct_link)-109)
	' 03 Get BLOB    
	WHTTP.Open "GET", Url03, False
	WHTTP.Send QuerySTring
	FileData = WHTTP.ResponseBody
	' 04 Save FILE
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
	BinaryStream.Type = 1
	BinaryStream.Open
	BinaryStream.Write FileData
	BinaryStream.SaveToFile FilePath, 2
  Set WHTTP = Nothing
End Sub



Sub UnZip(FolderPath, FilePath)
  ' https://www.robvanderwoude.com/vbstech_files_zip.php#X-UNZIP
  ' These are the available CopyHere options, according to MSDN
  ' (http://msdn2.microsoft.com/en-us/library/ms723207.aspx). On my test systems, however, the options were completely ignored.
  '      4: Do not display a progress dialog box.
  '      8: Give the file a new name in a move, copy, or rename operation if a file with the target name already exists.
  '     16: Click "Yes to All" in any dialog box that is displayed.
  '     64: Preserve undo information, if possible.
  '    128: Perform the operation on files only if a wildcard file name (*.*) is specified.
  '    256: Display a progress dialog box but do not show the file names.
  '    512: Do not confirm the creation of a new directory if the operation requires one to be created.
  '   1024: Do not display a user interface if an error occurs.
  '   4096: Only operate in the local directory. Don't operate recursively into subdirectories.
  '   8192: Do not copy connected files as a group. Only copy the specified files.   
  Dim intOptions, objShell, objSource, objTarget
  Set objShell = CreateObject( "Shell.Application" )
  Set objSource = objShell.NameSpace( FilePath ).Items( )
  Set objTarget = objShell.NameSpace( FolderPath )
  intOptions = 256 
  objTarget.CopyHere objSource, intOptions
  Set objSource = Nothing
  Set objTarget = Nothing
  Set objShell  = Nothing
End Sub
    
 
    
Function SendMail(FileNameXML, FilePathXML)  
  Dim msg
  Dim conf
  Dim config
  Dim myTargetDir
  Set msg = CreateObject("CDO.Message")
  Set conf = CreateObject("CDO.Configuration")
  Set config = conf.Fields
  With config
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "webmail.ags-hosting.fr"
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465 '
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = "true"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "fichier@pressepeople.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "******"
    .Update
  End With
  With msg
    Set .Configuration = conf
        .From = "AutoEmailerVBScript<fichier@pressepeople.com>"
        .Subject = FileNameXML
        .To = "ctp.pressepeople@gmail.com"
        .TextBody = "Bonjour, " & vbNewLine & "Cordialement, " & vbNewLine & "PAO PRESSE PEOPLE" & vbNewLine & vbNewLine
        .AddAttachment (FilePathXML)
        .Send
  End With
  Set msg = Nothing
  Set conf = Nothing
  Set config = Nothing
 End Function