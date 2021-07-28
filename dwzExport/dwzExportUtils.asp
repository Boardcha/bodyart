<%

class XmlError
	Private ErrNumber()
	Private ErrDescription()
	Private ErrPosition()
	Private ItemCount
	
	Private Sub Class_Initialize()
		ItemCount = -1
	End Sub

	Public sub Add(byval Number, byval Description, byval Position)
		ItemCount = ItemCount + 1
		ReDim Preserve ErrNumber(ItemCount)
		ReDim Preserve ErrDescription(ItemCount)
		ReDim Preserve ErrPosition(ItemCount)
		
		ErrNumber(ItemCount) = Number
		ErrDescription(ItemCount) = Description
		ErrPosition(ItemCount) = Position
	end sub
	
	public function hasErrors()
		if ItemCount = -1 then
			hasErrors = false
		else
			hasErrors = true
		end if
	end function
	
	Public Function getTable(header) 
		retStr = "<table border=1>"
		retStr = retStr & "<tr><td colspan=3>" & header & "</td></tr>"
		retStr = retStr & "<tr>"
		retStr = retStr & "<td>Number</td>"
		retStr = retStr & "<td>Description</td>"
		retStr = retStr & "<td>Position</td>"
		retStr = retStr & "</tr>"
		for J=0 to ItemCount
			tmp = "<tr>"
			tmp = tmp & "<td>" & ErrNumber(J) & "</td>"
			tmp = tmp & "<td>" & ErrDescription(J) & "</td>"
			tmp = tmp & "<td>" & ErrPosition(J) & "</td>"
			tmp = tmp & "</tr>"
			retStr = retStr & tmp
		next
		retStr = retStr & "</table>"
		getTable = retStr
	End Function
	
end class

class XmlContent
	Private ContArray()
	Private ItemCount
	
	Private Sub Class_Initialize()
		ItemCount = -1
	End Sub

	Public sub Add(byval Value)
		ItemCount = ItemCount + 1
		ReDim Preserve ContArray(ItemCount)
		ContArray(ItemCount) = Value
	end sub
	
	Public Function ToString() 
		ToString = Join(ContArray, "")
	End Function
	
end class

function dwz_Exp_DoDateTime(str, nNamedFormat, nLCID)				
	dim strRet, nOldLCID
	strRet = str
	If (nLCID > -1) Then
		oldLCID = Session.LCID
	End If
	On Error Resume Next
	If (nLCID > -1) Then							
		Session.LCID = nLCID
	End If
	If ((nLCID < 0) Or (Session.LCID = nLCID)) Then
		strRet = FormatDateTime(str, nNamedFormat)
	End If
	If (nLCID > -1) Then
		Session.LCID = oldLCID
	End If
	dwz_Exp_DoDateTime = strRet
end function


%>