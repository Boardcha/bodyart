<%

if request("date1") = "" then
  var_date1 = DatePart("yyyy",date()-30) _
  & "-" & Right("0" & DatePart("m",date()-30), 2) _
  & "-" & Right("0" & DatePart("d",date()-30), 2)
else
  var_date1 = Request.querystring("date1")
end if


if request("date2") = "" then
  var_date2 = DatePart("yyyy",Date) _
  & "-" & Right("0" & DatePart("m",Date), 2) _
  & "-" & Right("0" & DatePart("d",Date), 2)
else
  var_date2 = Request.querystring("date2")
end if

If request("reviewed") <> "on" Then ReviewdItems = " AND problem_reviewed = 0"

Error_flip = 0
Error_flip_total = 0
Error_matching = 0
Error_matching_total = 0
Error_broken = 0
Error_broken_total = 0
Error_wrong = 0
Error_wrong_total = 0
Error_missing = 0
Error_missing_total = 0
varError = 0
TotalErrorPoints = 0
ErrorAdd = 0

  '===== GET ALL ERRORS ====================================
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM QRY_FaultyOrders WHERE (date_sent >= ? AND date_sent <= ?) AND PackagedBy = ? " & ReviewdItems & " ORDER BY item_problem ASC, date_sent DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("date1",133,1,20, var_date1 ))
  objCmd.Parameters.Append(objCmd.CreateParameter("date2",133,1,20, var_date2 ))
  objCmd.Parameters.Append(objCmd.CreateParameter("PackagedBy",200,1,30, var_packer ))

	set rsGetErrors = Server.CreateObject("ADODB.Recordset")
	rsGetErrors.CursorLocation = 3 'adUseClient
	rsGetErrors.Open objCmd
	total_errors = rsGetErrors.RecordCount


  '==== GET TOTAL AMOUNT OF ORDERS ========================
  set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ID FROM sent_items WHERE (date_sent >= ? AND date_sent <= ?) AND PackagedBy = ?" 
	objCmd.Parameters.Append(objCmd.CreateParameter("date1",133,1,20, var_date1 ))
  objCmd.Parameters.Append(objCmd.CreateParameter("date2",133,1,20, var_date2 ))
  objCmd.Parameters.Append(objCmd.CreateParameter("PackagedBy",200,1,30, var_packer ))

	set rsGetTotalOrders = Server.CreateObject("ADODB.Recordset")
	rsGetTotalOrders.CursorLocation = 3 'adUseClient
	rsGetTotalOrders.Open objCmd
	total_orders = rsGetTotalOrders.RecordCount

  While NOT rsGetErrors.EOF 

  if (rsGetErrors.Fields.Item("item_problem").Value) = "Flip-flop" then 
	varError = 5
	Error_flip = 1
	else
	Error_flip = 0
end if
if (rsGetErrors.Fields.Item("item_problem").Value) = "Mis-matched" then 
	varError = 4
  	Error_matching = 1
	else
	Error_matching = 0

end if
if (rsGetErrors.Fields.Item("item_problem").Value) = "Broken" then 
	varError = 3
	Error_broken = 1
	else
	Error_broken = 0
end if
if (rsGetErrors.Fields.Item("item_problem").Value) = "Wrong" then 
	varError = 2
	Error_wrong = 1
	else
	Error_wrong = 0
end if
if (rsGetErrors.Fields.Item("item_problem").Value) = "Missing" then 
	varError = 2 * rsGetErrors.Fields.Item("ErrorQtyMissing").Value
	Error_missing = 1
	else
	Error_missing = 0
end if
if (rsGetErrors.Fields.Item("item_problem").Value) = "Misc" then 
	varError = 1
  Error_misc = 1
  else
	Error_misc = 0
end if
	
  ErrorAdd = ErrorAdd + varError
	Error_flip_total = Error_flip_total + Error_flip
	Error_matching_total = Error_matching_total + Error_matching
	Error_broken_total = Error_broken_total + Error_broken
	Error_wrong_total = Error_wrong_total + Error_wrong
	Error_missing_total = Error_missing_total + Error_missing
  Error_misc_total = Error_misc_total + Error_misc

  rsGetErrors.MoveNext()
Wend
rsGetErrors.requery()

if NOT rsGetErrors.EOF then
  var_error_percentage = Replace(Left((FormatNumber((total_orders / 10 - ErrorAdd) / (total_orders / 10),4, 0) * 100),3),".", "") & "%"
else
  var_error_percentage = "100%"
end if
%>