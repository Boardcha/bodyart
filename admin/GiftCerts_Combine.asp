<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


Dim rsGetCertificates__MMColParam
rsGetCertificates__MMColParam = "1"
If (Request.Form("code1") <> "") Then 
  rsGetCertificates__MMColParam = Request.Form("code1")
End If
%>
<%
Dim rsGetCertificates__Code2
rsGetCertificates__Code2 = "1"
If (Request.Form("code2") <> "") Then 
  rsGetCertificates__Code2 = Request.Form("code2")
End If
%>
<%
Dim rsGetCertificates__Code3
rsGetCertificates__Code3 = "1"
If (Request.Form("code3") <> "") Then 
  rsGetCertificates__Code3 = Request.Form("code3")
End If
%>
<%
Dim rsGetCertificates__Code4
rsGetCertificates__Code4 = "1"
If (Request.Form("code4") <> "") Then 
  rsGetCertificates__Code4 = Request.Form("code4")
End If
%>
<%
Dim rsGetCertificates__Code5
rsGetCertificates__Code5 = "1"
If (Request.Form("code5") <> "") Then 
  rsGetCertificates__Code5 = Request.Form("code5")
End If
%>
<%
Dim rsGetCertificates
Dim rsGetCertificates_cmd
Dim rsGetCertificates_numRows

Set rsGetCertificates_cmd = Server.CreateObject ("ADODB.Command")
rsGetCertificates_cmd.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetCertificates_cmd.CommandText = "SELECT ID, amount, code FROM dbo.TBLcredits WHERE code = ? OR code = ? OR code = ? OR code = ? OR code = ?" 
rsGetCertificates_cmd.Prepared = true
rsGetCertificates_cmd.Parameters.Append rsGetCertificates_cmd.CreateParameter("param1", 200, 1, 50, rsGetCertificates__MMColParam) ' adVarChar
rsGetCertificates_cmd.Parameters.Append rsGetCertificates_cmd.CreateParameter("param2", 200, 1, 255, rsGetCertificates__Code2) ' adVarChar
rsGetCertificates_cmd.Parameters.Append rsGetCertificates_cmd.CreateParameter("param3", 200, 1, 255, rsGetCertificates__Code3) ' adVarChar
rsGetCertificates_cmd.Parameters.Append rsGetCertificates_cmd.CreateParameter("param4", 200, 1, 255, rsGetCertificates__Code4) ' adVarChar
rsGetCertificates_cmd.Parameters.Append rsGetCertificates_cmd.CreateParameter("param5", 200, 1, 255, rsGetCertificates__Code5) ' adVarChar

Set rsGetCertificates = rsGetCertificates_cmd.Execute
rsGetCertificates_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
rsGetCertificates_numRows = rsGetCertificates_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Combine gift certificates</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
  <h5>Combine gift certificates</h5>

<% if request.querystring("status") = "results" then %>
<% If Not rsGetCertificates.EOF Or Not rsGetCertificates.BOF Then %>
<form action="GiftCerts_Combine.asp?status=email" method="post">                  
<%
varcode = 0 
varamount = 0
While ((Repeat1__numRows <> 0) AND (NOT rsGetCertificates.EOF)) 
%>
                    <strong><%=FormatCurrency(rsGetCertificates.Fields.Item("amount").Value,2)%></strong>&nbsp;&nbsp;<%=(rsGetCertificates.Fields.Item("code").Value)%><br>
                    <% 
  varcode = rsGetCertificates.Fields.Item("code").Value
  varamount = varamount + rsGetCertificates.Fields.Item("amount").Value

set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE TBLCredits SET amount = 0, CombinedInto = '" + varcode + "' WHERE code = '" + rsGetCertificates.Fields.Item("code").Value + "'"
commUpdate.Execute()
  
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsGetCertificates.MoveNext()
Wend
%>
<%
set commUpdate = Server.CreateObject("ADODB.Command")
commUpdate.ActiveConnection = MM_bodyartforms_sql_STRING
commUpdate.CommandText = "UPDATE TBLCredits SET amount = "& varamount &" WHERE code = '" + varcode + "'"
commUpdate.Execute()
%>
<span class="text-success font-weight-bold"><%= FormatCurrency(varamount,2) %> has been combined into CODE: <%= varcode %></span>
    <input name="amount" type="hidden" id="amount" value="<%= FormatCurrency(varamount,2) %>">
    <input name="code" type="hidden" id="code" value="<%= varcode %>">

</form>  <% End If ' end Not rsGetCertificates.EOF Or NOT rsGetCertificates.BOF %>

              <% end if ' show if there are results %>
<% if request.querystring("status") = "" then %>            <form action="GiftCerts_Combine.asp?status=results" method="post" name="form1">
                  CODE #1:
                <input name="code1" type="text" class="form-control form-control-sm my-1" id="code1" size="30" style="width: 300px">
                CODE #2:
                <input name="code2" type="text" class="form-control form-control-sm my-1" id="code2" size="30" style="width: 300px">
                CODE #3:
                <input name="code3" type="text" class="form-control form-control-sm my-1" id="code3" size="30" style="width: 300px">
                CODE #4:
                <input name="code4" type="text" class="form-control form-control-sm my-1" id="code4" size="30" style="width: 300px">
                CODE #5:
                <input name="code5" type="text" class="form-control form-control-sm my-1" id="code5" size="30" style="width: 300px">

      <div class="alert alert-info">Once submitting, it CAN NOT be undone! So be sure the codes above are the ones you want to combine!</div>
        <button class="btn btn-sm btn-secondary" type="submit" name="SearchCerts" id="SearchCerts">COMBINE</button>
     
            </form> 
            <% end if ' display if there are no results %>
</div>
</body>
</html>
<%
rsGetCertificates.Close()
Set rsGetCertificates = Nothing
%>
