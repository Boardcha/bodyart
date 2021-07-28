<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!DOCTYPE html> 
<html>
<head>
<title>Instagram Orders</title>
</head>
<body>
<!--#include file="admin_header.asp"-->
<div class="p-3">
        <h5 class="mb-5">Manage Instagram Orders NOT WORKING AS OF 10/9/20</h5>
              
              <form id="frm-instagram" method="post" enctype="multipart/form-data">
                <input class="d-block mb-3" name="file1" id="file1" type="file" accept=".csv">
                <button type="submit" class="btn btn-primary" id="btn-upload-instagram">Upload files<i class="fa fa-spin fa-lg fa-spinner ml-3" style="display:none" id="btn-spinner"></i></button>
              </form> 
              <div class="mt-5" id="html-message"></div>            
</div>
</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript">
  
  
    $("#btn-upload-instagram").click(function(e) {
      var myForm = document.getElementById('frm-instagram');
      var form_data = new FormData($(myForm)[0]);
      var file1 = $('#file1').val();
  
          //$('#btn-upload-instagram').prop('disabled', true);
        $('#btn-spinner').show();
      $('#html-message').html('');
   
      $.ajax({   
      method: "post",
      url: "instagram/instagram-upload-csv.asp",
      data: form_data,
      dataType: "json",
    //  cache: false,
      processData: false,
      contentType: false
      })
      .done(function(json, msg) {
        if(json.status == 'success') {
          $('#btn-spinner').hide();
          $('#html-message').html('<span class="alert alert-success p-2">Upload success</span>');
        }
        if(json.status == 'fail') {
          $('#btn-spinner').hide();
          $('#html-message').html('<span class="alert alert-danger p-2">Upload failed</span>');
        }
      })
      .fail(function(json, msg) {
          $('#btn-spinner').hide();
          $('#html-message').html('<span class="alert alert-danger p-2">Upload failed</span>');
      })

      e.preventDefault();
    });
    
</script>
<%
DataConn.Close()
%>