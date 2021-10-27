<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"
%>
<!DOCTYPE HTML>
<html>
<head>
<title>Approve reviews</title>
<link rel="stylesheet" href="plugins/tooltip/wenk.css">
</head>

<body>
<!--#include file="admin_header.asp"-->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.Prepared = true
objCmd.CommandText = "SELECT TOP (100) PERCENT TBLReviews.ReviewID, TBLReviews.review_edited, TBLReviews.customer_edit_date, review_rating, TBLReviews.email, jewelry.title, jewelry.ProductID, TBLReviews.review, TBLReviews.name, TBLReviews.status, TBLReviews.customer_ID, TBLReviews.ReviewOrderDetailID, jewelry.picture, TBLReviews.DetailID, ProductDetails.Gauge, ProductDetails.Length, ProductDetails.ProductDetail1, customers.customer_first, customers.customer_last, tblreviews.date_submitted, cs_flagged, InvoiceID FROM jewelry INNER JOIN TBLReviews ON jewelry.ProductID = TBLReviews.ProductID INNER JOIN ProductDetails ON TBLReviews.DetailID = ProductDetails.ProductDetailID INNER JOIN customers ON TBLReviews.customer_ID = customers.customer_ID INNER JOIN TBL_OrderSummary ON TBLReviews.ReviewOrderDetailID = TBL_OrderSummary.OrderDetailID WHERE (TBLReviews.status = N'pending') ORDER BY ReviewID ASC"
set rsGetReviews = Server.CreateObject("ADODB.Recordset")
rsGetReviews.CursorLocation = 3 'adUseClient
rsGetReviews.Open objCmd
rsGetReviews.PageSize = 25
total_records = rsGetReviews.RecordCount
intPageCount = rsGetReviews.PageCount

' Variables for paging
Select Case Request("Action")
    case "<<"
        intpage = 1
    case "<"
        intpage = Request("intpage")-1
        if intpage < 1 then intpage = 1
    case ">"
        intpage = Request("intpage")+1
        if intpage > intPageCount then intpage = IntPageCount
    Case ">>"
        intpage = intPageCount
    case else
        intpage = 1
end select
%>

<div class="my-3"></div>
<!--#include file="includes/inc-paging.asp"-->

<% If Not rsGetReviews.EOF Then %>

<div class="d-flex flex-row flex-wrap text-dark">

<%

rsGetReviews.AbsolutePage = intPage '======== PAGING
For intRecord = 1 To rsGetReviews.PageSize 
	If Not IsNull(rsGetReviews("customer_edit_date")) Then review_edited = true Else review_edited = false
%>

    <div class="col-12 col-xl-3 col-break1600-3 col-break1900-2 my-3 px-1 px-md-2" id="column_<%= rsGetReviews.Fields.Item("ReviewID").Value %>">
   <div class="card bg-light">
    <div class="card-heaader p-2">
        <a href="/productdetails.asp?ProductID=<%=(rsGetReviews.Fields.Item("ProductID").Value)%>" target="_blank"><img class="pull-left mr-2" style="height:50px;width:50px" src='http://bodyartforms-products.bodyartforms.com/<%=(rsGetReviews.Fields.Item("picture").Value)%>'   /></a> 
        <div><%=(rsGetReviews.Fields.Item("title").Value)%></div>
        <div class="font-weight-bold">
        <%=(rsGetReviews.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetReviews.Fields.Item("Length").Value)%>&nbsp;<%=(rsGetReviews.Fields.Item("ProductDetail1").Value)%>
    </div>
        <span class="text-warning">
            <% 	for i = 1 to rsGetReviews.Fields.Item("review_rating").Value %>
                 <i class="fa fa-star fa-lg"></i>
             <% next %>
           </span> 
		   - <%
		   If review_edited Then 
				Response.Write rsGetReviews("customer_edit_date") %>
				- <a class="text-danger" href="#" data-wenk="<%=Server.HTMLEncode(rsGetReviews("review"))%>">Edited!</a>
		   <%Else 
				Response.Write rsGetReviews("date_submitted") 
		   End If
		   %>
		   
			
    </div> 
    <div class="card-body p-2">
        <form class="frm-reviews" name="frm-reviews" data-id="<%= rsGetReviews.Fields.Item("ReviewID").Value %>" data-edited="<%If review_edited Then Response.Write "yes"%>">
            
              <textarea name="Review" cols="40" rows="5" class="form-control form-control-sm mb-2" id="Review">
			  <%If review_edited Then Response.Write rsGetReviews("review_edited") Else Response.Write rsGetReviews("review") %>
              </textarea>
            
              <select name="vote" class="form-control form-control-sm" >
                         <option value="accepted" selected="selected">Accept</option>
                         <option value="rejected">Reject - No e-mail sent</option>
                   <option value="We saw that there was an issue with the <%=(rsGetReviews.Fields.Item("title").Value)%> and we sent the info to our customer service department who will be contacting you shortly to make this right. The review was not currently approved, because we are hoping to enhance the experience for you. However, you are welcome to resubmit the review with any update after contact, or as it was originally submitted. <br><br>Thank you for submitting a review, the feedback really helps us to keep improving on our products and services and is always grately appreciated!">Reject - CS will be contacting</option>
                         <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because there were not enough words. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Not enough words</option>
                         <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because the review was a duplicate. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Duplicate review</option>
                         <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because the text review was empty / blank. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Blank review</option>
                         <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because it was submitted to the wrong product. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Submitted to wrong product</option>
                         <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because there was personal information in your review. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Personal info</option>
                        <option value="Unfortunately your review of <%=(rsGetReviews.Fields.Item("title").Value)%> was not approved because there was advertising. However, you are welcome to revise & resubmit the review. <br><br>Thank you!">Reject - Ads/Spam</option>
                       </select>        
                     
                       <input class="btn btn-sm btn-purple my-2" type="submit" name="SubmitReview" id="SubmitReview" value="Submit" />

                       
                       <input name="comments" type="text" class="form-control form-control-sm" placeholder="Our public comments (if needed)"/>
                        
                       <input type="hidden" name="review-id" value="<%= rsGetReviews.Fields.Item("ReviewID").Value %>" />
                       <input name="customer-id" type="hidden" value="<%=(rsGetReviews.Fields.Item("customer_ID").Value)%>" />
                       <input name="name" type="hidden" id="name" value="<%=(rsGetReviews.Fields.Item("name").Value)%>" />
                       <input name="email" type="hidden" id="email" value="<%=(rsGetReviews.Fields.Item("email").Value)%>" />
                       <input name="title" type="hidden" id="title" value="<%=(rsGetReviews.Fields.Item("title").Value)%>" />
                       <input name="OrderDetailID" type="hidden" id="OrderDetailID" value="<%=(rsGetReviews.Fields.Item("ReviewOrderDetailID").Value)%>" />
                     </form>
     </div>
     <div class="card-footer p-2">
        <div class="container-fluid p-0">
            <div class="row">
        <div class="col">
                Written by <a href="order history.asp?var_first=<%= rsGetReviews.Fields.Item("customer_first").Value %>&var_last=<%= rsGetReviews.Fields.Item("customer_last").Value %>" target="_blank"><%=(rsGetReviews.Fields.Item("name").Value)%></a>
        </div>
        <div class="col">
                Customer #<a href="customer_edit.asp?ID=<%=(rsGetReviews.Fields.Item("customer_ID").Value)%>" target="_blank"><%=(rsGetReviews.Fields.Item("customer_ID").Value)%></a>
        </div>
            </div>
        </div>

        <div class="container-fluid p-0 mt-2">
            <div class="row">
                <div class="col">
                    <% if rsGetReviews.Fields.Item("cs_flagged").Value = 0 then %>
                        <span class="text-primary pointer notify" data-dept="cs" data-productid="<%= rsGetReviews.Fields.Item("productid").Value %>"  data-invoiceid="<%= rsGetReviews.Fields.Item("InvoiceID").Value %>" data-reviewid="<%= rsGetReviews.Fields.Item("ReviewID").Value %>" data-title="<%= Server.HTMLEncode(rsGetReviews.Fields.Item("title").Value)%>" data-email="<%= rsGetReviews.Fields.Item("email").Value %>" data-details="<% if rsGetReviews.Fields.Item("Gauge").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("Gauge").Value)%><% end if %> - <% if rsGetReviews.Fields.Item("Length").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("Length").Value)%><% end if %> - <% if rsGetReviews.Fields.Item("ProductDetail1").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("ProductDetail1").Value)%><% end if %>" data-review="<%= Server.HTMLEncode(rsGetReviews.Fields.Item("review").Value) %>">Alert customer service</span>
                    <% else %>
                    <span class="text-success"><i class="fa fa-check mr-2"></i>Customer service alerted</span>
                    <% end if %>
                </div>
                <div class="col">
                        <span class="text-primary pointer notify" data-dept="photography" data-productid="<%= rsGetReviews.Fields.Item("productid").Value %>"   data-invoiceid="<%= rsGetReviews.Fields.Item("InvoiceID").Value %>" data-reviewid="<%= rsGetReviews.Fields.Item("ReviewID").Value %>"  data-title="<%= Server.HTMLEncode(rsGetReviews.Fields.Item("title").Value)%>" data-details="<% if rsGetReviews.Fields.Item("Gauge").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("Gauge").Value)%><% end if %> - <% if rsGetReviews.Fields.Item("Length").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("Length").Value)%><% end if %> - <% if rsGetReviews.Fields.Item("ProductDetail1").Value <> "" then %><%= Server.HTMLEncode(rsGetReviews.Fields.Item("ProductDetail1").Value)%><% end if %>"  data-review="<%= Server.HTMLEncode(rsGetReviews.Fields.Item("review").Value) %>">Alert Rebekah to fix photo</span>
                </div>
            
           
                
            </div>
        </div>
     </div>
   </div>
      
    </div><!-- column -->
  <% 
  rsGetReviews.MoveNext()
  If rsGetReviews.EOF Then Exit For  ' ====== PAGING
  Next ' ====== PAGING
%>
</div><!-- flex wrap -->

<%
  end if 'NOT rsGetReviews.eof
%>
<div class="my-5"></div>
<!--#include file="includes/inc-paging.asp"-->
</body>
</html>

<%
Set rsGetReviews = Nothing
DataConn.Close()
%>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">
	
	$('.frm-reviews').submit(function(e) {
        var reviewid = $(this).attr('data-id');
		var review_edited = $(this).attr('data-edited');
        console.log(reviewid);
		$.ajax({
		method: "post",
		url: "customers/ajax-accept-jewelry-review.asp",
		data: $(this).serialize() + "&review-edited=" + review_edited
		})
		.done(function(msg) {
			$('#column_' + reviewid ).fadeOut( "slow", function() {
                // Animation complete.
            });
		})
		.fail(function(msg) {
			alert("Website error");
		})
		
		e.preventDefault();
		return false;
	});

	$('.notify').click(function(e) {
        var department = $(this).attr('data-dept');
        var title = $(this).attr('data-title');
        var productid = $(this).attr('data-productid');
        var invoiceid = $(this).attr('data-invoiceid');
        var details = $(this).attr('data-details');
        var email = $(this).attr('data-email');
        var review = $(this).attr('data-review');
        var reviewid = $(this).attr('data-reviewid');
        var element = this;

        
		$.ajax({
		method: "post",
		url: "customers/ajax-reviews-notify-emails.asp",
		data: {department: department, title:title, productid: productid, details:details, email:email, review:review, reviewid:reviewid, invoiceid:invoiceid}
		})
		.done(function(msg) {
            $(element).prepend('<span class="text-success mr-2"><i class="fa fa-check"></i></span>');
		})
		.fail(function(msg) {
			alert("Website error");
		})
	});

	
</script>