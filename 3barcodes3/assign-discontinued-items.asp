<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<html>
<head>
<title>Assign discontinued items</title>
<meta name="viewport" content="width=device-width, minimum-scale=1.0, maximum-scale=1.0" />
<meta name="mobile-web-app-capable" content="yes">
<link href="css/scanners.css" rel="stylesheet" type="text/css" />
</head>
<body class="discontinued">
<div class="page-header">
	Assign discontinued product items to bin
</div>
	<input type="text" name="productid" id="productid" placeholder="Scan discontinued barcode" autofocus>
<form id="frm-discontinued" name="frm-discontinued" method="post" action="assign-discontinued-items.asp">
	<input type="text" name="bin" id="bin" placeholder="Scan bin number">
	<div id="loaded-info"></div>
	<br/>
	<br/>
	<button type="submit">Submit</button>
</form>
<script type="text/javascript" src="../js/jquery-2.1.1.min.js"></script>
<script type="text/javascript">
	$('#bin').hide();
	
	$('#productid').on('change', function() {
		$('#bin').show().focus();

		$('#loaded-info').removeClass();
		$('#loaded-info').load("includes/inc-load-product-info.asp", {productid: $('#productid').val()});		
	});
	
	// Submit form
	$("#frm-discontinued").submit(function(e) {
		$('#bin').hide();
		$('#productid').focus();
		var bin = $('#bin').val();
	
		$.ajax({
		method: "post",
		url: "includes/inc-assign-discontinued.asp",
		data: {productid: $('#productid').val(), bin: bin}
		})
		.done(function(msg) {
			$("#loaded-info").addClass("notice-green").html("Assigned to bin # " + bin).show();
			
		})
		.fail(function(msg) {
			$("#loaded-info").addClass("notice-red").html("ERROR").show();
		})
		$('#bin, #productid').val('');
		e.preventDefault();
        return false;
	});

</script>
</body>
</html>