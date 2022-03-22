
<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
if rsGetUser.bof AND rsGetUser.eof then
    response.redirect "login.asp"
end if 
%>

<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
    <meta name="mobile-web-app-capable" content="yes">
    <script src="https://use.fortawesome.com/dc98f184.js"></script>
    <link href="/CSS/baf.min.css?v=092519" rel="stylesheet" type="text/css" />
    <title>Tag Products</title>
</head>
<body>
 <!--#include file="includes/scanners-header.asp" -->          
    <h5 class="p-2">Tag Products</h5>
 <div class="px-2">
    <input class="form-control form-control-sm mb-2"  type="text" id="scan-item" placeholder="Scan ITEM">
    <div id="load-message" class="h5 mb-2"></div>
    <div id="load-body"></div>
</div>

</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="scripts/tag-items.js?v=031622"></script>