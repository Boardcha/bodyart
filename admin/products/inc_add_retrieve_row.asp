<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT *, TBL_Barcodes_SortOrder.ID_Description FROM dbo.ProductDetails INNER JOIN dbo.TBL_Barcodes_SortOrder ON dbo.ProductDetails.DetailCode = dbo.TBL_Barcodes_SortOrder.ID_Number WHERE ProductID = ? ORDER BY ProductDetailID DESC"
	objCmd.Parameters.Append(objCmd.CreateParameter("productid",3,1,10,request.querystring("productid")))
	Set rs_getdetails = objCmd.Execute()
	
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT GaugeShow FROM TBL_GaugeOrder ORDER BY GaugeOrder ASC" 
	Set rsGetGauges = objCmd.Execute()
	
	Set objCmd = Server.CreateObject ("ADODB.Command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "SELECT * FROM dbo.TBL_Barcodes_SortOrder" 
	Set rs_getsections = objCmd.Execute()

%>
<%' response.write request. %>
<tbody class="row-group ajax-update">
	<tr class="show-less <%= inactive_class %>">
		<td>
			<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>
			<% if rs_getdetails.Fields.Item("BinNumber_Detail").Value <> 0 then %>
			BIN <%=(rs_getdetails.Fields.Item("BinNumber_Detail").Value)%>
			<% end if %>
		</td>
		<td>
			<input class="form-control form-control-sm" name="sort_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= (rs_getdetails.Fields.Item("item_order").Value)%>"  data-column="item_order" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<select class="form-control form-control-sm" name="section_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"  data-column="DetailCode" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
				<option value="<%=(rs_getdetails.Fields.Item("DetailCode").Value)%>" selected>
					<%=(rs_getdetails.Fields.Item("ID_Description").Value)%>
				</option>
				<% While NOT rs_getsections.EOF %>
				<option value="<%=(rs_getsections.Fields.Item("ID_Number").Value)%>">
					<%=(rs_getsections.Fields.Item("ID_Description").Value)%>
				</option>
				<% 
				rs_getsections.MoveNext()
				Wend
				rs_getsections.MoveFirst()
				%> 
			</select>
		</td>
		<td>
			<input class="form-control form-control-sm" name="location_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= (rs_getdetails.Fields.Item("location").Value)%>" data-column="location" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="qty-onhand_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rs_getdetails.Fields.Item("qty").Value)%>" data-column="qty" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="max_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text"value="<% if Not isNull(rs_getdetails.Fields.Item("stock_qty").Value) then%><%=(rs_getdetails.Fields.Item("stock_qty").Value)%><% else %>0<% end if %>" data-column="stock_qty" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="thresh_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%=(rs_getdetails.Fields.Item("restock_threshold").Value)%>" data-column="restock_threshold" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<select class="form-control form-control-sm" name="gauge_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="Gauge" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
				<option value="<% If (rs_getdetails.Fields.Item("Gauge").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Gauge").Value)%><% end if %>" selected><% If (rs_getdetails.Fields.Item("Gauge").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Gauge").Value)%><% end if %></option>  
				<% While NOT rsGetGauges.EOF %>
				<option value="<%= Server.HTMLEncode(rsGetGauges.Fields.Item("GaugeShow").Value) %>"><%= rsGetGauges.Fields.Item("GaugeShow").Value %></option>
				<% rsGetGauges.MoveNext()
				Wend 
				rsGetGauges.ReQuery() %>
				<option value="">None</option>
				<option value="n/a">n/a</option>
			</select>
		</td>
		<td>
			<select class="form-control form-control-sm" name="length_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-column="Length" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
				<option value="<% If (rs_getdetails.Fields.Item("Length").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Length").Value)%><% end if %>" selected><% If (rs_getdetails.Fields.Item("Length").Value) <> "" Then %><%= Server.HtmlEncode(rs_getdetails.Fields.Item("Length").Value)%><% end if %></option>  
				<option value="">None</option>
				<option value="3/16&quot;">3/16&quot;</option>
				<option value="1/4&quot;">1/4&quot;</option>
				<option value="5/16&quot;">5/16&quot;</option>
				<option value="3/8&quot;">3/8&quot;</option>
				<option value="10mm">10mm</option>
				<option value="7/16&quot;">7/16&quot;</option>
				<option value="11mm">11mm</option>
				<option value="12mm">12mm</option>
				<option value="1/2&quot;">1/2&quot;</option>
				<option value="9/16&quot;">9/16&quot;</option>
				<option value="5/8&quot;">5/8&quot;</option>
				<option value="11/16&quot;">11/16&quot;</option>
				<option value="3/4&quot;">3/4&quot;</option>
				<option value="13/16&quot;">13/16&quot;</option>
				<option value="7/8&quot;">7/8&quot;</option>
				<option value="15/16&quot;">15/16&quot;</option>
				<option value="1&quot;">1&quot;</option>
				<option value="1-1/16&quot;">1-1/16&quot;</option>
				<option value="1-1/8&quot;">1-1/8&quot;</option>
				<option value="1-3/16&quot;">1-3/16&quot;</option>
				<option value="1-1/4&quot;">1-1/4&quot;</option>
				<option value="1-5/16&quot;">1-5/16&quot;</option>
				<option value="1-3/8&quot;">1-3/8&quot;</option>
				<option value="1-7/16&quot;">1-7/16&quot;</option>
				<option value="1-1/2&quot;">1-1/2&quot;</option>
				<option value="1-9/16&quot;">1-9/16&quot;</option>
				<option value="1-5/8&quot;">1-5/8&quot;</option>
				<option value="1-11/16&quot;">1-11/16&quot;</option>
				<option value="1-3/4&quot;">1-3/4&quot;</option>
				<option value="1-7/8&quot;">1-7/8&quot;</option>
				<option value="2&quot;">2&quot;</option>
				<option value="2-1/4&quot;">2-1/4&quot;</option>
				<option value="2-1/2&quot;">2-1/2&quot;</option>
				<option value="3&quot;">3&quot;</option>
			</select>
		</td>
		<td>
			<input class="form-control form-control-sm" name="details_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("ProductDetail1").value <> "" then%>value="<%= Server.HTMLEncode(rs_getdetails.Fields.Item("ProductDetail1").Value)%>"<% end if %> data-column="ProductDetail1" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="retail_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= FormatNumber(rs_getdetails.Fields.Item("price").Value,2)%>" data-column="price" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="wholesale_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" value="<%= FormatNumber(rs_getdetails.Fields.Item("wlsl_price").Value,2)%>" data-column="wlsl_price" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<input class="form-control form-control-sm" name="vendor-sku_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="text" <% if rs_getdetails.fields.item("detail_code").value <> "" then%>value="<%=(rs_getdetails.Fields.Item("detail_code").Value)%>"<% else %>value=" "<% end if %> data-column="detail_code" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">
		</td>
		<td>
			<% 	if (rs_getdetails.Fields.Item("active").Value) = 1 then
					var_checked = "checked"
				else
					var_checked = ""
				end if

			%>
			<input name="active_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" type="checkbox" value="1" <%= var_checked %>  data-column="active" data-detailid="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-unchecked="0">
		</td>
		<td>
			&nbsp;
		</td>
		<td>
			<span class="input_move" name="input_move_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"><span><span class="btn btn-sm btn-secondary font-weight-bold copyid" name="copy_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>"><i class="fa fa-copy-far fa-lg"></i></span>
			<span class="btn btn-sm btn-secondary font-weight-bold moveid" name="move_<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>" data-id="<%= rs_getdetails.Fields.Item("ProductDetailID").Value %>">M</span>
		</td>
	</tr>
</tbody>

<%DataConn.Close() %>