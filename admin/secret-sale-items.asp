<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM jewelry WHERE secret_sale = 1 ORDER BY title ASC"
set rsCheck = objCmd.Execute()
%>
<html>
<head>
<title>Secret sale items</title>
</head>
<body style="background-color:#fff">
<!--#include file="admin_header.asp"-->
<link href="/CSS/baf.min.css?v=040919" rel="stylesheet" type="text/css" />
<div class="p-4">
        <h5 class="mb-3">Secret sale items</h5>
<div>
    Secret sale items will only show to customers that have linked from a secret sale URL. In order to create a URL for a secret sale it MUST have secret=yes in it. Also note, that these settings aren't immediate. They go live with each 11am product push each day.
    <br/><br/>
    For example: https://bodyartforms.com/products.asp?jewelry=basics&jewelry=septum<strong>&secret=yes</strong>
    <br/>
    OR https://bodyartforms.com/productdetails.asp?ProductID=32371<strong>&secret=yes</strong>
    <br/>

</div>

  <table class="table">
    <thead class="thead-dark">
    <tr>
      <th>Item</th>
    </tr>
	</thead>
        <% 
While NOT rsCheck.EOF
%>
    <tr>
           <td><a href="product-edit.asp?ProductID=<%=(rsCheck.Fields.Item("ProductID").Value)%>&info=less">
               <img src="http://bafthumbs-400.bodyartforms.com/<%=(rsCheck.Fields.Item("picture").Value)%>" width="90" height="90">
            <%=(rsCheck.Fields.Item("title").Value)%>
        </a></td>
    </tr>
    <% 
  rsCheck.MoveNext()
Wend
%>
</table>
</div>
</body>
</html>
<%
rsCheck.Close()
Set rsCheck = Nothing
%>