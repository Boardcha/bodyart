<%
	' Function that checks for duplicates
	Function CheckDupe(pass_code)
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "SELECT DiscountCode FROM TBLDiscounts WHERE DiscountCode = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("Code",200,1,50,pass_code))
		'set rsCodeCheckDupe = objCmd.Execute()

		set rsCodeCheckDupe = Server.CreateObject("ADODB.Recordset")
		rsCodeCheckDupe.CursorLocation = 3 'adUseClient
		rsCodeCheckDupe.Open objCmd

		Unique = ""
		IF NOT rsCodeCheckDupe.EOF THEN
			DO UNTIL Unique = true
				' call function again that was originally variable strRandomCode to get a new random string
				var_cert_code = getPassword(15, extraChars, firstNumber, firstLower, firstUpper, firstOther, latterNumber, latterLower, latterUpper, latterOther)
				var_function_count = var_function_count + 1
				strSQL2 = "SELECT DiscountCode FROM TBLDiscounts WHERE DiscountCode = '" & var_cert_code & " ' "
				SET GenerateNewCode = DataConn.Execute(strSQL2)
				IF GenerateNewCode.EOF THEN
					Unique = true
				END IF
			LOOP
			SET GenerateNewCode = Nothing 
		END IF
		SET rsCodeCheckDupe = Nothing 
		CheckDupe = var_cert_code
	End function 
%>