<%@LANGUAGE="VBSCRIPT" %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Approve photos</title>
<link rel="stylesheet" href="/CSS/jquery.fancybox.min.css" />
</head>

<body class="review-body">
<!--#include file="admin_header.asp"-->
<%
' decrypt customer ID cookie

Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

password = "3uBRUbrat77V"
data = request.Cookies("ID")

If len(data) > 5 then ' if
	decrypted = objCrypt.Decrypt(password, data)
end if

  if data <> decrypted then
	  CustomerID = decrypted
  else
	  CustomerID = 0
  end if

Set objCrypt = Nothing

set DataConn = Server.CreateObject("ADODB.connection")
DataConn.Open MM_bodyartforms_sql_STRING

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandType = 4
objCmd.CommandText = "SP_inc_GetUser_byID"
objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustomerID))
Set rsGetUser = objCmd.Execute()
rsGetUser_numRows = 0


Set rsGetPhotoReviews = Server.CreateObject("ADODB.Recordset")
rsGetPhotoReviews.ActiveConnection = MM_bodyartforms_sql_STRING
rsGetPhotoReviews.Source = "SELECT jewelry.title, TBL_PhotoGallery.PhotoID, TBL_PhotoGallery.DateSubmitted, TBL_PhotoGallery.filename, TBL_PhotoGallery.type, TBL_PhotoGallery.description, TBL_PhotoGallery.status, TBL_PhotoGallery.ProductID, TBL_PhotoGallery.customerID, TBL_PhotoGallery.Name, TBL_PhotoGallery.Email, TBL_PhotoGallery.DetailID, ProductDetails.ProductDetail1, ProductDetails.Gauge, ProductDetails.Length, customers.customer_first,  customers.customer_last FROM jewelry INNER JOIN TBL_PhotoGallery ON jewelry.ProductID = TBL_PhotoGallery.ProductID INNER JOIN ProductDetails ON TBL_PhotoGallery.DetailID = ProductDetails.ProductDetailID LEFT OUTER JOIN customers ON TBL_PhotoGallery.customerID = customers.customer_ID WHERE (TBL_PhotoGallery.status = 0) ORDER BY PhotoID ASC"
 '    AND TBL_PhotoGallery.DateSubmitted < getdate()-1
rsGetPhotoReviews.CursorLocation = 3 'adUseClient
rsGetPhotoReviews.LockType = 1 'Read-only records
rsGetPhotoReviews.Open()

%>

<div class="d-flex flex-row flex-wrap text-dark">

<% While NOT rsGetPhotoReviews.EOF %>

<div class="col-12 col-xl-3 col-break1600-3 col-break1900-2 my-3 px-1 px-md-2 group photoid_<%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>">
	<div class="card bg-light">
		<div class="card-heaader p-2">
	<a data-fancybox="customer-images" title="<%= Server.HTMLEncode((rsGetPhotoReviews.Fields.Item("title").Value) & " " & rsGetPhotoReviews.Fields.Item("description").Value)%>" href="http://bodyartforms-gallery.bodyartforms.com/<%=(rsGetPhotoReviews.Fields.Item("filename").Value)%>"><img src='http://bodyartforms-gallery.bodyartforms.com/thumb_<%=(rsGetPhotoReviews.Fields.Item("filename").Value)%>' alt="Enlarge photo" class="review-image-s3" style="width: 200px;height: 200px" /></a>

	<img src="../gallery/uploads/thumb_<%=(rsGetPhotoReviews.Fields.Item("filename").Value)%>"  class="review-image-local" style="top: 0;left: 0;width: 70px;height: 70px;border: 2px solid grey;" />
</div>


	<div class="card-body p-2">
		<form class="frm-review" name="frm-vote-photo" id="frm-vote-photo-<%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>" data-photoid="<%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>">
		<div class="title">
			<a href="../productdetails.asp?ProductID=<%=(rsGetPhotoReviews.Fields.Item("ProductID").Value)%>" target="_blank"><%=(rsGetPhotoReviews.Fields.Item("title").Value)%>&nbsp;<%=(rsGetPhotoReviews.Fields.Item("ProductDetail1").Value)%> &nbsp;<%=(rsGetPhotoReviews.Fields.Item("Gauge").Value)%> &nbsp;<%=(rsGetPhotoReviews.Fields.Item("Length").Value)%></a>
		</div>
			<span class="input-productid-label">Product ID:</span> 
			<input  class="form-control form-control-sm mb-2"  name="ProductID" type="text" class="input-productid" value="<%=(rsGetPhotoReviews.Fields.Item("ProductID").Value)%>"/>
      
        <select  class="form-control form-control-sm"  name="photo_status" id="photo_status">
          <option value="1" selected="selected">Accept</option>
          <option value="the jewelry is blurry or pixelated. Please make sure the camera is focused on the actual jewelry and not other things around it">Reject (Blurry)</option>
          <option value="it is too distorted on the sizing and stretches out the jewelry">Reject (Distorted)</option>
          <option value="it is too dark and hard to see">Reject (Too dark)</option>
          <option value="it is too light and hard to see">Reject (Too bright)</option>
          <option value="the glare on the jewelry makes it hard to see">Reject (Light glare)</option>
          <option value="it needs to be more clear. The photo seems pixelated and was a bit blurry. Please make sure the camera is focused on the actual jewelry and not other things around it">Reject (Low pixels/Hard to see)</option>
          <option value="it is too difficult to see any detail in the item from the direction the photo was taken">Reject (Hard to see detail from direction)</option>
          <option value="the images did not upload correctly">Reject (Broken images)</option>
		  <option value="the images did not upload correctly">Reject (Mis-matched images)</option>
          <option value="the picture was submitted to the wrong product">Reject (Submitted to wrong product)</option>
          <option value="the size was too small to see any detail">Reject (Photo too small)</option>
          <option value="the jewelry must be worn in order to be approved">Reject (Not wearing jewelry)</option>
          <option value=" we do not give points for stickers and gauge cards">Reject (Sticker/gauge card)</option>
          <option value="we already have this photo in the gallery">Reject (Duplicate)</option>
          <option value="too small to see the jewelry close up">Reject (Too small to see)</option>
          <option value="too far away to see jewelry close-up">Reject (Too far away)</option>
          <option value="does not accept photos with nudity">Reject (Nudity)</option>
          <option value="the photo has personal information attached to it">Reject (Personal information)</option>
          <option value="the photo contains advertising">Reject (Advertising)</option>
          <option value="0">Reject (Don't send email)</option>
        </select>

        <button class="btn btn-sm btn-purple my-2" type="submit" formmethod="post" >Submit</button>
      <input name="PhotoID" type="hidden" id="PhotoID" value="<%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>" />
      <input name="photo_filename" type="hidden" id="photo_filename" value="<%=(rsGetPhotoReviews.Fields.Item("filename").Value)%>" />
      <input name="customerID" type="hidden" id="customerID" value="<%=(rsGetPhotoReviews.Fields.Item("customerID").Value)%>" />
      <input name="Name" type="hidden" id="Name" value="<%=(rsGetPhotoReviews.Fields.Item("Name").Value)%>" />
      <input name="Email" type="hidden" id="Email" value="<%=(rsGetPhotoReviews.Fields.Item("Email").Value)%>" />
      <input name="title" type="hidden" id="title" value="<%=(rsGetPhotoReviews.Fields.Item("title").Value)%>" />
	</form>
	</div><!-- card body-->
	  <div class="card-footer p-2">
		<span class="photoid_message_<%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>"></span>
		Customer # <a href="customer_edit.asp?ID=<%=(rsGetPhotoReviews.Fields.Item("customerID").Value)%>" target="_blank"><%=(rsGetPhotoReviews.Fields.Item("customerID").Value)%></a>
		<br/>
		Photo ID # <%= rsGetPhotoReviews.Fields.Item("PhotoID").Value %>
		<br/>
		Submitted <%= rsGetPhotoReviews.Fields.Item("DateSubmitted").Value %>
	</div>


        
    </div><!-- card -->
    </div><!-- column -->   
  
<% 
  rsGetPhotoReviews.MoveNext()
Wend
%>
</div><!-- flex -->   
</body>

<script src="/js/jquery-3.3.1.min.js"></script>

<script src="/js/jquery.fancybox.min.js"></script>
<script type="text/javascript">
	// Submit approval or rejection
	$('.frm-review').submit(function (event) {
		event.preventDefault(event); // Do not reload page
		var photoid = $(this).attr('data-photoid');
	
		$('.photoid_' + photoid).addClass('highlight-green');
	console.log(photoid);
		$.ajax({
		method: "post",
		url: "vote_photos.asp",
		data: $('#frm-vote-photo-' + photoid).serialize()
		})
		.done(function(msg) {
			$('.photoid_' + photoid).fadeOut('slow');			
		})
		.fail(function(msg) {
			$('.photoid_message_' + photoid).html('<div class="notice-red">Website error.</div>').show();
		})
		
		return false;
		
	});  // END Submit approval or rejection
</script>
</html>
<%
rsGetPhotoReviews.Close()
DataConn.Close()
Set rsGetUser = Nothing
Set rsGetPhotoReviews = Nothing
%>