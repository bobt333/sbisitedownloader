strXMLURL = "https://www.mysite.com/n7SX2pWQ.xml"

Set objHTTP = CreateObject("MSXML2.XMLHTTP")
Dim objStream
Set objStream = CreateObject("ADODB.Stream")
Set objFSO = CreateObject("Scripting.FileSystemObject")

strAppPath = Left(WScript.ScriptFullName, instrrev(WScript.ScriptFullName, "\"))
objHTTP.open "get", strXMLURL, false
objHTTP.send
arrXML = split(objHTTP.responseText, chr(10))

'Create files folder... overwrite if neccessary...
strSaveToFolder = strAppPath & "files"
If objFSO.FolderExists(strSaveToFolder) then
    objFSO.DeleteFolder strSaveToFolder, True  'True indicates forceful deletion
End If
objFSO.CreateFolder strSaveToFolder
strSaveToFolder = strSaveToFolder & "\" 'add trailing slash

n = 0
For i = 1 to ubound(arrXML)
	strLine = trim(arrXML(i))
    if left(strLine, 13) = "<loc>https://"  or left(strLine, 12) = "<loc>http://" then
        n = n + 1
        nStart = 6 'start and len strips off <loc>...</loc>
        nLen = len(strLine) - 11
        strURL = mid(strLine, nStart, nLen)
		'msgbox "x" & strURL & "x"
		DownloadLink(strURL)
    end if
Next
Set objFSO = Nothing
set objStream = Nothing
WScript.Echo "Finished. Downloaded " & n & " files to " & strSaveToFolder
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub DownloadLink(strLink)

	strSaveName = Mid(strLink, InStrRev(strLink,"/") + 1, Len(strLink))
	strTLD = right(strSaveName, 4)
	if len(strSaveName) = 0 or strTLD = ".com" or strTLD = ".net" or strTLD = ".org" or strTLD = ".xyz" then
		strSaveName = "index.html"
	end if

	strSaveTo = strSaveToFolder & strSaveName

	objHTTP.open "GET", strLink, False
	objHTTP.send

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FileExists(strSaveTo) Then
	  objFSO.DeleteFile(strSaveTo)
	End If
	If objHTTP.Status = 200 Then
	  With objStream
		.Type = 1 'adTypeBinary
		.Open
		.Write objHTTP.responseBody
		.SaveToFile strSaveTo
		.Close
	  End With
	End If

	If objFSO.FileExists(strSaveTo) Then
	  'WScript.Echo "Downloaded " & strSaveName
	End If

End Sub
