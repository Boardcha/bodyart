<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objcmd = Server.CreateObject("ADODB.command")
objcmd.ActiveConnection = DataConn
objcmd.CommandText = "SELECT * FROM tbl_shipping_notice"
Set rsGetNotice = objcmd.Execute()
%>
<html>
    <head>
            <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<link rel="stylesheet" type="text/css" href="/CSS/redactor.css" />
<script type="text/javascript" src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="/js/redactor.js"></script>
<script type="text/javascript" src="/js/redactor-plugin-source.js"></script>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
        <form id="frm-update-notice">
<h6 class="font-weight-bold mb-3">Type shipping notice below</h6>
<textarea name="description" id="description"><%= rsGetNotice.Fields.Item("shipping_notice").Value %></textarea>


<div class="mt-4">To disable notice, select "None" below</div>

<select class="form-control my-2" name="country">
       <option value="<%= rsGetNotice.Fields.Item("country").Value %>" selected><%= rsGetNotice.Fields.Item("country").Value %></option> 
       <option value="">NONE - No notice</option>
   <option value="USA">USA</option>
    <option value="Canada">Canada</option>
</select>   
   <button type="submit" class="btn btn-purple my-4">Save</button>
       </form>


<div id="message"></div>
 <p></p>   
 <p></p>  
 <p></p>  
 <p></p>  
 <p></p>  
 <p></p>      
</div>



</body>
</html>
<script type="text/javascript">
    $("#description").redactor({
        buttons: ['html', 'bold', 'italic', 'underline'],
        formatting: ['p', 'h3', 'h4', 'h5'],
        plugins: ['source']
    });

    	// START update account profile name
		$('#frm-update-notice').submit(function (e) {
            $('#message').html('<i class="fa fa-spinner fa-2x fa-spin"></i>').show();

            $.ajax({
            method: "post",
            url: "administrative/ajax-shipping-notice.asp",
            data: $("#frm-update-notice").serialize()
            })
            .done(function(msg) {
                $('#message').html('<div class="alert alert-success">Shipping notice saved</div>');
            })
            .fail(function(msg) {
                $('#message').html('<div class="alert alert-danger">Site error</div>');
            })

       e.preventDefault();
    });  // END update account profile name	
</script>
<%
DataConn.Close()
%>