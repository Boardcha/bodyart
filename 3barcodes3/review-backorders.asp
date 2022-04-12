
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if rsGetUser.bof AND rsGetUser.eof then
    response.redirect "login.asp"
end if 

Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Barcodes_SortOrder" 
Set rs_getsections = objCmd.Execute()
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
    <meta name="mobile-web-app-capable" content="yes">
    <script src="https://use.fortawesome.com/dc98f184.js"></script>
    <link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
    <title>Review backorders</title>
</head>
<body>
 <!--#include file="includes/scanners-header.asp" -->
<div class="p-2">

<!--#include virtual="/includes/inc-review-backorders.inc"-->
</div>



</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/bootstrap-v4.min.js"></script>
<!--#include virtual="/includes/review-backorders.asp"-->
