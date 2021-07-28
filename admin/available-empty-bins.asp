<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

'===============================================================================================
' This page is a difficult one. We have to find the active items first and then build out a system to kinda reverse show anything that's available in between the active items
'===============================================================================================

if request.querystring("letter") = "" then
    bin_letter = "A"
else
    bin_letter = request.querystring("letter")
end if


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT b.ID_Description, b.ID_Number, d.location, d.qty, CASE WHEN d.active = 1 THEN 'Active' ELSE 'Inactive' END as 'active', CASE WHEN p.active = 1 THEN 'Active' ELSE 'Inactive' END as 'p_active', p.title, p.ProductID FROM ProductDetails AS d INNER JOIN TBL_Barcodes_SortOrder as b ON d.DetailCode = b.ID_Number INNER JOIN jewelry AS p ON d.ProductID = p.ProductID WHERE b.ID_Description = ? ORDER BY location ASC, d.active DESC"
objCmd.Parameters.Append(objCmd.CreateParameter("bin_letter",200,1,10, bin_letter ))
set rsGetEmptyBins = Server.CreateObject("ADODB.Recordset")
rsGetEmptyBins.CursorLocation = 3 'adUseClient
rsGetEmptyBins.Open objCmd
rsGetEmptyBins.PageSize = 1
total_records = rsGetEmptyBins.RecordCount
intPageCount = rsGetEmptyBins.PageCount

%>

<html>
<head>
<title>Available empty bins</title>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="m-3">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<h5>Available empty bins</h5>

<div class="btn-toolbar mb-4" role="toolbar">
    <div class="btn-group btn-group-sm text-white" role="group">
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=A">A</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=B">B</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=C">C</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=D">D</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=E">E</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=F">F</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=G">G</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=H">H</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=I">I</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=J">J</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=K">K</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=L">L</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=M">M</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=N">N</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=O">O</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=P">P</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=Q">Q</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=R">R</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=S">S</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=T">T</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=U">U</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=V">V</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=W">W</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=X">X</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=Y">Y</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=Z">Z</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=AA">AA</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=BB">BB</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=CC">CC</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=DD">DD</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=Large">Large</a></button>
      <button type="button" class="btn btn-secondary"><a class="text-white" href="?letter=Party">Party</a></button>
    </div>
  </div>




<table class="table table-bordered table-hover table-sm small">
    <thead>
        <tr>
            <th>Location</th>
            <th>Current stock</th>
            <th>Active status</th>
            <th>Description</th>
        </tr>
    </thead>
<%

while not rsGetEmptyBins.eof
current_location = rsGetEmptyBins.Fields.Item("location").Value
bin_difference = current_location - prior_location - 1

'==== Highlight row if it's a duplicate location ======
if current_location = prior_location then
    row_styling = "class='d-none'"
else
    row_styling = ""
end if

'==== Highlight row if it's a duplicate location ======
if current_location = prior_location AND rsGetEmptyBins.Fields.Item("active").Value = "Active" then
    row_styling = "class='alert alert-info'"
end if

if rsGetEmptyBins.Fields.Item("active").Value = "Active" then
    row_styling = "class='d-none'"
end if

if rsGetEmptyBins.Fields.Item("qty").Value  > 0 then
    qty_styling = "class='alert alert-warning'"
else
    qty_styling = ""
end if

if bin_difference > 1 then
For i = 1 To bin_difference
    calc_location_start = current_location - bin_difference - 1
    calc_location = calc_location_start + i
%>
<tr>
    <td>
        <span class="mr-3"><%= bin_letter %></span><%= calc_location %>
    </td>
    <td>-</td>
    <td>
        Unassigned bin
    </td>
    <td>
        Unassigned bin
    </td>
</tr>
<% 
    If i=30 Then Exit For
Next
end if
%>
<tr <%= row_styling %>>
    <td>
            <span class="mr-3"><%= bin_letter %></span><%= current_location  %>
    </td>
    <td <%= qty_styling %>>
        <%= rsGetEmptyBins.Fields.Item("qty").Value  %>
    </td>
    <td>
        <span class="mr-4"><%= rsGetEmptyBins.Fields.Item("active").Value  %> detail</span>
        <%= rsGetEmptyBins.Fields.Item("p_active").Value  %> product
    </td>
    <td>
        <a href="product-edit.asp?ProductID=<%= rsGetEmptyBins.Fields.Item("ProductID").Value %>"><%= rsGetEmptyBins.Fields.Item("title").Value  %></a>
    </td>
</tr>
<%   
    prior_location = current_location
    rsGetEmptyBins.MoveNext()
    wend
%>
</table>



<% else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</div>

</body>
</html>
<%
rsGetEmptyBins.Close()
Set rsGetUser = Nothing

%>
