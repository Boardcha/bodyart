
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
    <title>Pull orders</title>
</head>
<body>
 <!--#include file="includes/scanners-header.asp" -->
 <nav class="navbar navbar-expand navbar-light small py-1" style="background-color: #D8D8D8"> 
        <div class="navbar-nav mr-auto">
                 <div class="form-inline">
                  <button class="btn btn-sm btn-info mr-2" id="btn-clear" type="button"><i class="fa fa-eraser fa-lg"></i></button> <button class="btn btn-sm btn-info ml-1 mr-2" id="btn-reset" type="button"><i class="fa fa-refresh fa-lg"></i></button><input  class="form-control form-control-sm ml-1 mr-1" type="text" id="scan-invoice" placeholder="Scan INVOICE" style="max-width:120px">
                 <div class="ml-2" style="line-height: 100%">
                    #<span id="display-invoice"></span>
                </div>
                <div class="ml-2">
                    <span id="display-conserve"></span>
                </div>
                </div>
        </div>           
    </nav>

<div id="load-message"></div>
<div id="load-body"></div>

<!-- Process backorder Modal -->
<div class="modal fade" id="modal-submit-backorder" tabindex="-1" role="dialog"  aria-labelledby="modal-submit-backorder" >
    <div class="modal-dialog" role="document">
      <div class="modal-content">
        <div class="modal-header">
          <h5 class="modal-title">Submit Backorder</h5>
          <button type="button" class="close close-bo" data-dismiss="modal" aria-label="Close">
            <span aria-hidden="true">&times;</span>
          </button>
        </div>
        <div class="modal-body">
            <div id="new-bo-message"></div>
            <form id="frm-backorder">
            Is the entire order on backorder?<br/>
            <div class="custom-control custom-radio custom-control-inline">
                <input type="radio" id="entireorder_no" name="entireorder" class="custom-control-input" value="" checked>
                <label class="custom-control-label" for="entireorder_no">No</label>
            </div>
            <div class="custom-control custom-radio custom-control-inline">
                <input type="radio" id="entireorder_yes" name="entireorder" class="custom-control-input" value="Entire order on BO">
                <label class="custom-control-label" for="entireorder_yes">Yes</label>
            </div>
<br/><br/>
    		<select class="form-control form-control-sm" name="partial" id="partial">
                <option value="" selected>If partial items sent (select qty) ...</option>
                <option value="">None sent</option>
                <option value="1 sent">1 sent</option>
                <option value="2 sent">2 sent</option>
                <option value="3 sent">3 sent</option>
                <option value="4 sent">4 sent</option>
                <option value="5 sent">5 sent</option>
                <option value="6 sent">6 sent</option>
                <option value="7 sent">7 sent</option>
                <option value="8 sent">8 sent</option>
                <option value="9 sent">9 sent</option>
                <option value="10 sent">10 sent</option>
            </select>
			<br />
			<h6>Reason for backorder:</h6>
      <h6>Reason for backorder:</h6>
      <div class="custom-control custom-radio">
        <input class="custom-control-input" name="BOReason" type="radio" id="radio" value="there were not enough items left in stock" checked>
        <label class="custom-control-label" for="radio">Not enough items left in stock</label>
      </div>
      <div class="custom-control custom-radio">
        <input class="custom-control-input" type="radio" name="BOReason" id="radio2" value="the last pair did not match">
        <label class="custom-control-label" for="radio2">Last pair did not match</label>
      </div>
      <div class="custom-control custom-radio">
        <input class="custom-control-input" type="radio" name="BOReason" id="radio3" value="the last ones were not the right size">
        <label class="custom-control-label" for="radio3">Last ones were not the right size</label>
      </div>
      <div class="custom-control custom-radio">
        <input class="custom-control-input" type="radio" name="BOReason" id="radio5" value="it was broken">
        <label class="custom-control-label" for="radio5">It was broken</label>
      </div>			
            </form>
        </div>
        <div class="modal-footer">
            <button type="button" class="btn btn-primary" id="btn-submit-bo" data-orderdetailid="">Submit</button>
          <button type="button" class="btn btn-secondary close-bo" data-dismiss="modal">Close</button>
        </div>
      </div>
    </div>
</div>
<!-- End Process backorder Modal -->


<!-- Process Error Alert Modal -->
<div class="modal fade" id="modal-submit-error" tabindex="-1" role="dialog"  aria-labelledby="modal-submit-error" >
  <div class="modal-dialog" role="document">
    <div class="modal-content">
      <div class="modal-body">
          <div id="message-error"></div>
			<form id="frm-error">
            Select item issue<br/>
            <select class="form-control form-control-sm mb-2" name="item_issue" id="item_issue">
              <option value="Print new scanning label">Print new scanning label</option>
              <option value="Wrong items in bin">Wrong items in bin</option>
              <option value="Broken item">Broken item</option>
              <option value="Quantity is off">Quantity is off</option>
              <option value="Combine bags">Combine bags</option>
              <option value="Needs to be paired">Needs to be paired</option>
              <option value="Singles are bad match">Singles are bad match</option>
              <option value="Needs new location">Needs new location</option>
            </select>
			Additional info (optional):<br/>
			<textarea class="form-control form-control-sm" id="error_description"></textarea>	  
            </form>
      </div>
      <div class="modal-footer">
          <button type="button" class="btn btn-primary" id="btn-submit-error" data-orderdetailid="">Submit</button>
        <button type="button" class="btn btn-secondary close-bo" data-dismiss="modal">Close</button>
      </div>
    </div>
  </div>
</div>
<!-- End Process backorder Modal -->


</body>
</html>
<script src="/js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
<script type="text/javascript" src="scripts/pull-orders.js?v=040422"></script>