<%
If rsProduct("standard_anodization") = True Then
	sqlWhere = "WHERE standard_anodization = 1"
Else
	sqlWhere = "WHERE standard_anodization = 0"
End IF	
If rsProduct("high_voltage_anodization") = True Then
	sqlWhere = sqlWhere & " OR high_voltage = 1"
Else
	sqlWhere = sqlWhere & " OR high_voltage = 0"
End IF	
'Response.Write sqlWhere & "aaa"


set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
	
objCmd.CommandText = "SELECT * FROM TBL_Anodization_Colors_Pricing " & sqlWhere & "  ORDER BY color_name ASC"

set rsGetItems = Server.CreateObject("ADODB.Recordset")
rsGetItems.CursorLocation = 3 'adUseClient
rsGetItems.Open objCmd
var_totalitems = rsGetItems.RecordCount

if not rsGetItems.eof then
%>
<div class="alert alert-warning">	
<div class="dropdown my-2" id="add-anodization-menu" style="<%= var_cart_modal %>">
<button class="btn btn-light rounded-0 bg-white text-left dropdown-toggle font-weight-bold  py-2 w-100" style="border:1px solid #ced4da" type="button" id="dropdownAddAnodization" data-flip="false" data-toggle="dropdown" aria-haspopup="true" aria-expanded="false" style="<%= var_cart_modal %>"><span id="selected-anodization">Optional color add-on:</span>
</button>
<div class="dropdown-menu modal-scroll-long rounded-0 w-100" style="border:2px solid #ced4da" aria-labelledby="dropdownAddAnodization">
		<div id="msg-filtered-dropdown"></div>
<div class="dropdown-item bg-white btn-group-vertical btn-group-toggle m-0 p-0 " data-toggle="buttons">

			<label class="btn rounded-0 py-3 py-lg-2 border-bottom text-left btn-select-menu">
				<input class="add-anodization" type="radio" name="add-anodization" value="0" data-anod-id="0" data-base-price="0" data-title="No custom color wanted" dropdown-title="No custom color wanted"
				data-variant="No custom color wanted">
				No custom color wanted / Plain polish
			</label> 
			<%
			i_count = 0
			While NOT rsGetItems.EOF	
			
				base_price = formatnumber(rsGetItems("base_price"),2)
				anodId = rsGetItems("anodID")
				%>
				<label class="btn rounded-0 py-3 py-lg-2 border-bottom text-left btn-select-menu">
					<input class="add-anodization" type="radio" name="add-anodization" value="<%=(rsGetItems("anodID").Value)%>" data-anod-id="<%= anodId %>" data-base-price="<%= base_price %>"  data-title="<%= replace(rsGetItems("color_name").Value, """", "") %>" dropdown-title="<%= exchange_symbol %><%= base_price %>
					&nbsp;&nbsp;&nbsp;&nbsp;<%= server.htmlencode(rsGetItems("color_name").Value) %>" data-variant="<%= trim(server.htmlencode(rsGetItems("color_name").Value)) %>">
					<%= exchange_symbol %><%= base_price %>&nbsp;&nbsp;&nbsp;&nbsp; <%=rsGetItems("color_name").Value%>&nbsp;&nbsp;&nbsp;&nbsp
				</label> 
				<%
				
				i_count = i_count + 1
				rsGetItems.MoveNext()
			Wend 
			rsGetItems.Requery() 
			%>

		</div><!-- button group -->
		</div><!-- drop down menu -->
		</div><!-- drop down -->
<!-- for cart page update -->
<input type="hidden" name="cartid" value="<%= request("cartid") %>">
<a class="text-info" href="/images/color-charts/anodization-colors.jpg" data-fancybox="anodization-chart" >View color chart</a>
	<div class="mt-2">
	<i class="fa fa-alert fa-lg mr-2"></i>Adding custom coloring service will delay the shipment of your order up to 1 week or less.
	</div>
</div>
<%
end if 	' not rsGetItems.eof 
%>