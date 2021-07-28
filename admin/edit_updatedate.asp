<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
If Request("MM_recordId") <> "" Then

	SqlString = "UPDATE jewelry SET date_added = '" & now() & "' WHERE ProductID = " + Request.Form("MM_recordId") 
	DataConn.Execute(SqlString)

End If
%>
<html>
<head>
<title>Update date</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#666699" text="#CCCCCC">
<%
If Request("MM_recordId") = "" Then


	Set rsUpdateDate = Server.CreateObject("ADODB.Recordset")
	rsUpdateDate.ActiveConnection = MM_bodyartforms_sql_STRING
	rsUpdateDate.Source = "SELECT * FROM jewelry WHERE ProductID = " + Request.QueryString("ID") + ""
	rsUpdateDate.CursorLocation = 3 'adUseClient
	rsUpdateDate.LockType = 1 'Read-only records
	rsUpdateDate.Open()
%>
<form name="frm_updatedate" method="POST" action="<%=MM_editAction%>">

  <p><font size="2" face="Verdana"><strong><%=(rsUpdateDate.Fields.Item("title").Value)%></strong></font></p>
  <p> 
    <input type="submit" name="Submit" value="Update into What's New section">
  </p>
  <input type="hidden" name="MM_recordId" value="<%= rsUpdateDate.Fields.Item("ProductID").Value %>">

</form>
<%
	rsUpdateDate.Close()
	Set rsUpdateDate = Nothing

Else
%>
	<script language=javascript>
		window.close();
	</script>	
<%
End if
%>
</body>
</html>
