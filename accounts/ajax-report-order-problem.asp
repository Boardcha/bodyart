<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->

<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID, qty, title, ProductDetail1, price, OrderDetailID, gauge, length, picture FROM QRY_OrderDetails WHERE InvoiceID = ? ORDER BY OrderDetailID ASC"
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10, request.form("id")))
Set rsGetItems = objCmd.Execute()

%>
<h6 id="header-select-item">Select the item that has an issue...</h6>
<% While NOT rsGetItems.EOF %>
			<div class="custom-control custom-radio my-1 problem-select-item" data-qty="<%= rsGetItems.Fields.Item("qty").Value %>">
				<input type="radio" class="custom-control-input ml-3 mr-2" id="item-<%= rsGetItems.Fields.Item("OrderDetailID").Value %>" name="report-item" value="<%= rsGetItems.Fields.Item("OrderDetailID").Value %>" data-qty="<%= rsGetItems.Fields.Item("qty").Value %>" required>
				<label class="custom-control-label" for="item-<%= rsGetItems.Fields.Item("OrderDetailID").Value %>">
					<img src="https://bodyartforms-products.bodyartforms.com/<%= rsGetItems.Fields.Item("picture").Value %>" style="width:40px;height:auto">
						<%=(rsGetItems.Fields.Item("title").Value)%> &nbsp;&nbsp;<%=(rsGetItems.Fields.Item("gauge").Value)%>&nbsp;<%=(rsGetItems.Fields.Item("length").Value)%>&nbsp;&nbsp;<%=(rsGetItems.Fields.Item("ProductDetail1").Value)%>
				</label>
			</div>
<% 
rsGetItems.MoveNext()
Wend
rsGetItems.requery()
%>

<div id="block-select-problem" style="display:none">
	<h6>What was wrong with your item?</h6>
		<div class="custom-control custom-radio my-1">
			<input class="custom-control-input ml-3 mr-2" name="status" id="missing" type="radio" value="Missing" required>
			<label class="custom-control-label" for="missing">Missing</label>
		</div>
		<div class="custom-control custom-radio my-1">
			<input class="custom-control-input ml-3 mr-2" name="status" id="broken" type="radio" value="Broken" required>
			<label class="custom-control-label" for="broken">Broken</label>
		</div>	
		<div class="custom-control custom-radio my-1">
				<input class="custom-control-input ml-3 mr-2" name="status" id="wrong" type="radio" value="Wrong" required>
				<label class="custom-control-label" for="wrong">Wrong</label>
		</div>
		<div class="form-group qty-missing" style="display:none">
			<label for="qty-missing">How many items were <span id="error-type"></span>?</label>
			<input class="form-control" name="qty-missing" id="qty-missing" type="tel" value="0" />
	</div>
		<div class="custom-control custom-radio my-1">
				<input class="custom-control-input ml-3 mr-2" name="status" id="mismatched" type="radio" value="Mis-matched" required>
				<label class="custom-control-label" for="mismatched">Mis-matched</label>
		</div>
		<div class="custom-control custom-radio my-1">
				<input class="custom-control-input ml-3 mr-2" name="status" id="other" type="radio" value="Misc" required>
				<label class="custom-control-label" for="other">Other</label>
		</div>

          
      <h6>Please give us more details...</h6>
	 	 <textarea class="form-control" name="description" maxlength="200" rows="4" required></textarea>

</div><!-- select problem block -->		  
<%
DataConn.Close()
Set DataConn = Nothing
%>
