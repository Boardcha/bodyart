<%
'
' ADOBE SYSTEMS INCORPORATED
' Copyright 2007 Adobe Systems Incorporated
' All Rights Reserved
' 
' NOTICE:  Adobe permits you to use, modify, and distribute this file in accordance with the 
' terms of the Adobe license agreement accompanying it. If you have received this file from a 
' source other than Adobe, then your use, modification, or distribution of it requires the prior 
' written permission of Adobe.
'

'
'	Copyright (c) InterAKT Online 2000-2005
'

	Class KT_FakeRecordset
		Public Fields
		
		Public CursorLocation
		Public CursorType
		Public LockType
		
		Public ActiveConnection
		Public Source
		
		' // constructor 
		private Sub Class_Initialize()
			' // initialisation of all properties goes here
			Set Fields = new KT_FakeItemsClass
		End Sub

		Public Function Close()
			Set Fields = Nothing
		End Function
		
		Public Function MoveFirst
			Fields.m_position = 1
			If Fields.m_position <= Fields.m_noRows Then
				Fields.EOF = False
				Fields.BOF = False
			End If
		End Function
		
		Public Function Requery
			MoveFirst()
		End Function
		
		Public Function MoveNext
			If Fields.m_position = Fields.m_noRows Then
				Fields.EOF = True
				'Fields.BOF = True
			End If	
			Fields.m_position = Fields.m_position + 1 
		End Function
		
		Public Function Move(newPosition, fromWhere)
			If fromWhere <> 0 And fromWhere <> 1 Then
				' move not supported
				Move = False
				Exit Function
			End If
			
			If fromWhere = 0 Then ' Starts at the current record
				If (Fields.m_position + newPosition) <= Fields.m_noRows And (Fields.m_position + newPosition) >= 1 Then
					Fields.m_position = Fields.m_position + newPosition
					Move = True
				Else
					Move = False
				End If
				Exit Function
			End If
			
			If fromWhere = 1 Then '  Starts at the first record
				If (newPosition + 1) <= Fields.m_noRows And newPosition >=0  Then
					Fields.m_position = newPosition + 1
					Move = True
				Else
					Move = False
				End If
				Exit Function
			End If
			
		End Function
		
		Public Function RecordCount
			RecordCount = Fields.m_noRows
		End Function
		
		Public Function AddNew	
			Fields.AddNew
		End Function
		
		Public Function BOF
			BOF = Fields.BOF
		End Function
		
		Public Function EOF
			EOF = Fields.EOF
		End Function
		
		Public Function Open
			
		End Function
		
		Public Default Property Get Value (p_fieldName)
			Set Value = Fields (p_fieldName)
		End Property
	End Class



	Class KT_FakeItemClass
		Private fieldValue
		private Sub Class_Initialize()
			fieldValue = ""
		End Sub
	
		public Default Property Get Value()
			Value = fieldValue
		End Property
		
		public Property Let Value(p_fieldValue)	
			If isnull(p_fieldValue) Then
				p_fieldValue = ""
			End If
			fieldValue = p_fieldValue
		End Property
	End Class



	Class KT_FakeItemsClass
		Public m_FieldValues	' dictionary with indexed by rows / fieldNames
		Public m_Fields		' current used fields

		Public m_noRows	
		Public m_noFields
		Public m_position
		Public EOF
		Public BOF			
		
		
		Private Sub Class_Initialize()
			Set m_Fields = Server.CreateObject("Scripting.Dictionary")
			Set m_FieldValues = Server.CreateObject("Scripting.Dictionary")
			m_noRows = 0
			m_noFields = 0
			m_position = 0
			EOF = True
			BOF = True
		End Sub
		
		Public default property Get Value(p_fieldName)
			If m_FieldValues.Exists(m_position) Then
				If m_FieldValues(m_position).Exists (lcase(p_fieldName)) Then
					Set Value = m_FieldValues(m_position)(lcase(p_fieldName))
				Else
					Response.write KT_getResource("ASP_FAKERS_MISSING_FIELD", "FakeRecordset", array(p_fieldName))
					Response.End()
				End If
			Else
				Response.write KT_getResource("ASP_FAKERS_NO_RECORDS", "FakeRecordset", array())
				Response.end
			End If
		End property

	
		Public Function Item(p_fieldName)
			If m_FieldValues.Exists(m_position) Then
				If m_FieldValues(m_position).Exists (lcase(p_fieldName)) Then
					Set Item = m_FieldValues(m_position)(lcase(p_fieldName))
				Else
					Response.write KT_getResource("ASP_FAKERS_MISSING_FIELD", "FakeRecordset", array(p_fieldName))
					Response.End()
				End If
			Else
				Response.write KT_getResource("ASP_FAKERS_NO_RECORDS", "FakeRecordset", array())
				Response.end
			End If
		End Function
		
		
		
		Public Function Append (p_fieldName)
			If m_Fields.Exists (lcase(p_fieldName)) Then
				Response.write KT_getResource("ASP_FAKERS_FIELD_EXISTS", "FakeRecordset", array())
				Response.End()
			Else
				m_Fields.Add lcase(p_fieldName), ""
				m_noFields = m_noFields + 1
			End If
		End Function		
		
		Public Function AddNew
			If m_noFields = 0 Then
				Response.write KT_getResource("ASP_FAKERS_NO_FIELDS", "FakeRecordset", array())
				Response.end
			Else
				m_position = m_position + 1
				Set m_FieldValues(m_position) = Server.CreateObject("Scripting.Dictionary")
				' add the column values
				For each fieldname in m_Fields
					Set m_FieldValues(m_position)(fieldName) = new KT_FakeItemClass
				Next
				m_noRows = m_noRows + 1
				
				EOF = false
				BOF = false
			End If
		End Function
	End Class
 			
%>
