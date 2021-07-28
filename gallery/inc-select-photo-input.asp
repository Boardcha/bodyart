<% @LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->

<%
	var_order_detailid = 0
if request.form("var_order_detailid") <> "" then 
	var_order_detailid = request.form("var_order_detailid")
end if

' Pull the customer information from a cookie
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
	Set rsGetUser = objCmd.Execute()
	
	var_email = ""
	var_name = ""
If Not rsGetUser.EOF then
	var_email = rsGetUser.Fields.Item("email").Value
	var_name = rsGetUser.Fields.Item("customer_first").Value
end if

if var_order_detailid = 0 then
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT ProductDetailID, ProductDetail1, Gauge, Length FROM dbo.ProductDetails WHERE ProductID = ?"
	objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10,request.querystring("productid")))
	Set rsGetProductDetails = objCmd.Execute()
end if
%>

<form name="frm-submit-photo" id="frm-submit-photo" method="post" enctype="multipart/form-data" data-orderitemid="<%= var_order_detailid %>">

<div class="close-link2 toggle-photo">Close <i class="fa fa-times fa-lg"></i></div>

<div class="photo-requirements">
  <span class="bold">Requirements:</span>
          <ul>
            <li>You must own the photo and must be wearing the jewelry.</li>
            <li>File type: .jpg only</li>
            <li>Image size no greater than 1.5MB. Make sure you resize your photo.</li>
            <li>File quality: Must be clear, close-up, and easy to see. No blurry, dark, or non quality photos will be approved.</li>
           <li>Not allowed: Nudity or &quot;private&quot; areas<br>
            </li>
      </ul>
</div>
<% if var_order_detailid = 0 then %>
	<div class="control-group">
	<label for="photo-detailid">Select item that matches your photo <span class="required-field">*</span></label>
	<select name="photo-detailid" required>
		<option value="" selected>Select gauge/product</option>
			<% 
			While NOT rsGetProductDetails.EOF 
			%>
				<option value="<%=(rsGetProductDetails.Fields.Item("ProductDetailID").Value)%>"><%=(rsGetProductDetails.Fields.Item("ProductDetail1").Value)%>&nbsp;&nbsp;<%=(rsGetProductDetails.Fields.Item("Gauge").Value)%>&nbsp;&nbsp;<%=(rsGetProductDetails.Fields.Item("Length").Value)%></option>
			<% 
			rsGetProductDetails.MoveNext()
			Wend
			%>
	</select>
	</div>         

	<div class="control-group">
	<label for="photo-name">Your name <span class="required-field">*</span></label>
	<input name="photo-name" type="text" value="<%= var_name %>" required >
	</div>

	<div class="control-group">
	<label for="photo-email">Your e-mail (will not be sold/spammed) <span class="required-field">*</span></label>
	<input name="photo-email" type="email" value="<%= var_email %>" required>
	</div>	
	<br/>
<% else ' if user IS submitting from the order history page then prefill out fields %>
	<input name="photo-detailid" type="hidden" value="<%= request.form("detailid") %>">
	<input name="photo-email" type="hidden" value="<%= var_email %>">
	<input name="photo-name" type="hidden" value="<%= var_name %>" >
<% end if ' var_order_detailid = 0 %>

	<div class="control-group">
	<span class="required-field">*</span> <input name="photo-filename" type="file" accept="image/jpg, image/jpeg" required>
	</div>
	
	<input name="productid" type="hidden" value="<%= request("productid") %>">
	<input name="custid" type="hidden" value="<%= CustID_Cookie %>">
	<input name="order-detailid" type="hidden"  value="<%= var_order_detailid %>">        
     
<div class="photo-release">
	By submitting your photo, you are authorizing Bodyartforms to use your  photo publicly on the Bodyartforms website, social media sites, or for any other Bodyartforms advertising purposes. All of your personal information will be kept confidential and only your submitted photo could potentially be used.
</div>

	<button id="submit-photo" name="submit-photo" class="btn_purple" type="submit">Upload picture
	</button>
	
	<div class="load-upload-message hide"><i class="fa fa-spinner fa-2x fa-spin" id="upload-spinner"></i> Uploading ... this could take 1-3 minutes depending on the size of your photo.</div>
</form>