<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"


set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn

%>

<html>
<head>
<title>Review backorders</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
</head>
<body>
<!--#include file="admin_header.asp"-->
<style>
@page{size:landscape}
</style>
<div class="p-2">
        
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

<!--#include virtual="/includes/inc-review-backorders.inc"-->


<% else ' unathorized access error %>
Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript" src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/bootstrap-v4.min.js"></script>
<!--#include virtual="/includes/review-backorders.asp"-->
<script type="text/JavaScript" src="/js/jQuery.print.min.js"></script>
<script type="text/javascript">

$("#btn_print").click(function(){

	$("#bo-print").print({
		globalStyles: true,
		mediaPrint: false,
		stylesheet: null,
		noPrintSelector: ".no-print",
		iframe: true,
		append: null,
		prepend: null,
		manuallyCopyFormValues: true,
		deferred: $.Deferred(),
		timeout: 750,
		title: null,
		doctype: '<!doctype html>'
	});
});
</script>
</body>
</html>
