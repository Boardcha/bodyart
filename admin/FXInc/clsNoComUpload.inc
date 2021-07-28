<SCRIPT LANGUAGE="VBScript" RUNAT="Server">

'***************************************************************************
    ' FELIXONE © 2002/2005 - All rights reserved.
'***************************************************************************

Class NoComUpload
	
    Private AllData
    Private Head
    Private q
    Private Sent

   '***************************************************************************
    ' Constructor
   '***************************************************************************
    Private Sub Class_Initialize()
	If (Request.TotalBytes > 0) Then
	    Dim AllBin, x
		
	    ' Flag for posted data
	    Sent = True
			
	    ' Retrieve posted data
	    q = Chr(34)
	    AllBin = Request.BinaryRead(Request.TotalBytes)
			
	    ' Convert it to Unicode
	    For x = 1 To LenB(AllBin)
	        AllData = AllData & Chr(AscB(MidB(AllBin, x, 1)))
	    Next
			
	    ' Get header
	    Head = Left(AllData, InStr(AllData, vbCrLf) + 1)
	Else
	    Sent = False
	End If
    End Sub
    

   '***************************************************************************
    ' Get the Request.Form
   '***************************************************************************
    Public Function ReqForm()
        If (Sent) Then
            Dim myArray, ReqStr, thename, thevalue, y, i
            myArray = split(AllData,"Content-Disposition: form-data; ")

            For i = 0 To UBound(myArray)
                If (i > 0) Then
                    x = InStr(myArray(i), q)
                    y = InStr(x+1, myArray(i), q)
                    thename = Mid(myArray(i), x+1, y-(x+1))
                    If (InStr(myArray(i), "filename=") > 0) Then
                        x = InStr(y+1, myArray(i), q)
                        y = InStr(x+1, myArray(i), q)
                        thevalue = escape(Mid(myArray(i), x+1, y-(x+1)))
                    Else
                        x = InStr(y+1, myArray(i), vbCrLf)
                        y = InStr(x, myArray(i), Replace(Head, vbCrLf, ""))
                        thevalue = escape(Mid(myArray(i), x+4, y-(x+6)))
                    End If
                    Del = "&"
                    If (ReqStr = "") Then Del = ""
                    ReqStr = ReqStr & Del & thename & "=" & thevalue
                End If
            Next

            ReqForm = ReqStr
        End If
    End Function


   '***************************************************************************
    ' Get the form element values
   '***************************************************************************
    Public Function Form(field)
        If (Sent) Then
            Dim reqStr, mStr, i
            reqStr = split(ReqForm, "&")
            For i=0 To UBound(reqStr)
                x = split(reqStr(i), "=")
                If (field = x(0)) Then
                    Del = ", "
                    If (mStr = "") Then Del = ""
                    mStr = mStr & Del & unescape(x(1))
                End If
            Next
            Form = mStr
        End If
    End Function


   '***************************************************************************
    ' Retrieve the filename
   '***************************************************************************
    Public Function GetFileName(field)
	If (Sent) Then
	    Dim x2, i

	    x = InStr(AllData, Head & "Content-Disposition: form-data; name=" & q & field & q)
	    If (x > 0) Then
		x = InStr(x, AllData, "filename=" & q)
		x2 = InStr(x, AllData, vbCrLf)
				
		For i = x2 To x Step -1
		    If Mid(AllData, i, 1) = "\" Then
			x = i - 9
			Exit For
		    End If
		Next
				
		GetFileName = Mid(AllData, x + 10, x2 - (x + 11))
	    End If
	End If
    End Function
    

   '***************************************************************************
    ' Create the path where to save the file
   '***************************************************************************
    Public Function MakeFilePath(directory, field)
	MakeFilePath = directory & "/" & GetFileName(field)
    End Function
    

   '***************************************************************************
    ' Get data from a specified field
   '***************************************************************************
    Private Function GetFileData(field)
	Dim x2
	
	x = InStr(AllData, Head & "Content-Disposition: form-data; name=" & q & field & q)
	
	If (x > 0) Then
	    x = InStr(x, AllData, vbCrLf)
	    x = InStr(x + 1, AllData, vbCrLf)
	    x = InStr(x + 1, AllData, vbCrLf) + 2
	    x2 = InStr(x, AllData, Left(Head, Len(Head) - 2))
			
	    GetFileData = Mid(AllData, x + 2, x2 - x - 4)
	End If
    End Function
    

   '***************************************************************************
    ' save data to disk
   '***************************************************************************
    Private Sub SaveFile(aField, aDirectory, overwrite) 'As Boolean
	Dim FSO, FU, aFilename
		
	Set FSO = CreateObject("Scripting.FileSystemObject")
        If (Not FSO.FolderExists(aDirectory)) Then FSO.CreateFolder(aDirectory)
        aFilename = MakeFilePath(aDirectory, aField)
	Set FU = FSO.CreateTextFile(aFilename, overwrite, False) ' last parameter is Unicode
	FU.Write GetFileData(aField)
	FU.Close
	Set FU = Nothing
	Set FSO = Nothing

    End Sub


   '***************************************************************************
    ' Save it with its name
   '***************************************************************************
    Public Function Save(directory, field, overwrite) 'As Boolean
        If (Sent And Form(field) <> "") Then
	    SaveFile field,directory,overwrite
	End If
    End Function
    
End Class
	

'***************************************************************************
    ' Make it available for JScript
'***************************************************************************
    set NoComUploadJs = new NoComUpload

</SCRIPT>