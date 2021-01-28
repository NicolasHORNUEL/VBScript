myProcess "\\MASTER-XMF\DropFolder\EXPORT PDF brut pour GS\agent sur windows", "\\MASTER-XMF\DropFolder\EXPORT PDF compressé"

Sub myProcess(InputFolder, OutputFolder)

	Dim objShell, objFSO
	Dim strExe, strParams01, strParams02
	Dim objFolder, colFiles, objFile, inputFilePath, outputFilePath, FilePathBAT

	Set objShell = CreateObject("Wscript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	strExe = "C:\Users\Administrateur\Desktop\PRESSEPEOPLE\gs9.50\bin\gswin64c.exe"
	strParams01 = " -dSAFER -dNOPAUSE -dBATCH -sDEVICE=pdfwrite -dPDFSETTINGS=/prepress -sOUTPUTFILE="
	strParams02 = " -dColorImageDownsampleType=/Bicubic -dColorImageResolution=150 "

	Set objFolder = objFSO.GetFolder(InputFolder)
	Set colFiles = objFolder.Files
	For Each objFile in colFiles
		If UCase(objFSO.GetExtensionName(objFile.name)) = "PDF" Then

			inputFilePath = Chr (34) & InputFolder & "\" & objFile.Name & Chr (34)
			outputFilePath =  Chr (34) & InputFolder & "\" & "temp.pdf"  & Chr (34)
			objShell.Run strExe & strParams01 & outputFilePath & strParams02 & inputFilePath, 1, true

			FilePathBAT = OutputFolder & "\" & objFile.Name
 			objFSO.MoveFile InputFolder & "\" & "temp.pdf", FilePathBAT
  			objFSO.DeleteFile InputFolder & "\" & objFile.Name, true
		End If
	Next
	Set objFSO = Nothing
End Sub



'https://www.ghostscript.com/doc/9.50/Use.htm#Platforms
'https://sourceforge.net/p/ghostscript/discussion/5452/thread/5a6f3978/?limit=25


