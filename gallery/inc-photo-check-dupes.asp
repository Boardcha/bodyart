<%
	' Function that checks customer gallery photos in database for duplicate image names
	Function CheckDupe(pass_code)
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT filename FROM TBL_PhotoGallery WHERE filename = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,200,pass_code))
		set rsPhotoCheckDupe = objCmd.Execute()

		Unique = ""
		 IF rsPhotoCheckDupe.EOF AND rsPhotoCheckDupe.BOF THEN
			
		  ELSE
			DO UNTIL Unique = true
			  ' call function again that was originally variable strRandomCode to get a new random string
			  var_new_filename = sha256(filename & "x" & getSalt(32, extraChars))
			  var_function_count = var_function_count + 1
			  strSQL2 = "SELECT filename FROM TBL_PhotoGallery WHERE filename = '" & var_new_filename & " ' "
			  SET objRS = DataConn.Execute(strSQL2)
			  IF objRS.EOF THEN
				Unique = true
			  ELSE
				intCount = intCount
			  END IF
			LOOP
			SET objRS = Nothing 
		  END IF
		  SET rsPhotoCheckDupe = Nothing 
	End function 
%>