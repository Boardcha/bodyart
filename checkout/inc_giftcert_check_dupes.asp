<%
If var_giftcert = "yes"  Then

	' Function that checks for duplicates
	Function CheckDupe(pass_code)
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT code FROM TBLCredits WHERE code = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,50,pass_code))
		set rsCertCheckDupe = objCmd.Execute()

		Unique = ""
		 IF rsCertCheckDupe.EOF AND rsCertCheckDupe.BOF THEN
			nFinalUsername = strUsername
		  ELSE
			DO UNTIL Unique = true
			  ' call function again that was originally variable strRandomCode to get a new random string
			  var_cert_code = getPassword(passwordLength, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
			  var_function_count = var_function_count + 1
			  strSQL2 = "SELECT code FROM TBLCredits WHERE code = '" & var_cert_code & " ' "
			  SET objRS = DataConn.Execute(strSQL2)
			  IF objRS.EOF THEN
				Unique = true
			  ELSE
				intCount = intCount
			  END IF
			LOOP
			SET objRS = Nothing 
		  END IF
		  SET rsCertCheckDupe = Nothing 
	End function 

end if
%>