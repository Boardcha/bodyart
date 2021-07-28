<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
if request("gauge") <> "" and request("gauge") <> "All" then
	sql_gauge = "AND ProductDetails.Gauge = ?"
	filter_gauge = replace(request("gauge"),"+", " ")
end if

if request("color") <> "" and request("color") <> "All" then
	sql_color = "AND ProductDetails.ProductDetail1 = ?"
	filter_color = replace(request("color"),"+", " ")
end if

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn

if request("this_value") <> "All" then

if request("type") = "reviews" then
    if request("filter") = "gauge" then
    var_type = "reviews-gauge"
    var_select_type = "reviews"
    var_select_by = "color"
        
      objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID 	WHERE (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.ProductDetail1 IS NOT NULL) " & sql_gauge & " GROUP BY ProductDetails.ProductDetail1"
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))
        if request("gauge") <> "" then
            objCmd.Parameters.Append(objCmd.CreateParameter("filter_gauge",200,1,15,filter_gauge))
        end if
    end if

    if request("filter") = "color" then
    var_type = "reviews-color"
    var_select_type = "reviews"
    var_select_by = "gauge"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total, TBL_GaugeOrder.GaugeOrder FROM  TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID WHERE (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.Gauge IS NOT NULL) " & sql_color & " GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder ORDER BY TBL_GaugeOrder.GaugeOrder"

        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))
        if request("color") <> "" then
            objCmd.Parameters.Append(objCmd.CreateParameter("filter_color",200,1,75,filter_color))
        end if
    end if

end if ' if reviews

if request("type") = "photos" then
    if request("filter") = "gauge" then
        var_type = "photos-gauge"
        var_select_type = "photos"
        var_select_by = "color"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE (ProductDetails.ProductDetail1 IS NOT NULL) AND (TBL_PhotoGallery.ProductID = ?) AND (TBL_PhotoGallery.status = 1)  " & sql_gauge & " GROUP BY ProductDetails.ProductDetail1"

        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))
        if request("gauge") <> "" then
            objCmd.Parameters.Append(objCmd.CreateParameter("filter_gauge",200,1,15,filter_gauge))
        end if
    end if

    if request("filter") = "color" then
        var_type = "photos-color"
        var_select_type = "photos"
        var_select_by = "gauge"

        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total FROM TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE (ProductDetails.Gauge IS NOT NULL) AND TBL_PhotoGallery.ProductID = ? AND TBL_PhotoGallery.status = 1 " & sql_color & " GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder, TBL_PhotoGallery.ProductID, TBL_PhotoGallery.status ORDER BY TBL_GaugeOrder.GaugeOrder"

        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))
        if request("color") <> "" then
            objCmd.Parameters.Append(objCmd.CreateParameter("filter_color",200,1,75,filter_color))
        end if
    end if
end if ' if photos

else ' if showing all
    filter_color = "All"
    filter_gauge = "All"


 if request("type") = "reviews" then
    if request("filter") = "gauge" then
        var_type = "reviews-color"
        var_select_type = "reviews"
        var_select_by = "gauge"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total, TBL_GaugeOrder.GaugeOrder FROM  TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID WHERE active = 1 and (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.Gauge IS NOT NULL) GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder ORDER BY TBL_GaugeOrder.GaugeOrder"
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))
    end if

    if request("filter") = "color" then
        var_type = "reviews-gauge"
        var_select_type = "reviews"
        var_select_by = "color"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBLReviews ON ProductDetails.ProductDetailID = TBLReviews.DetailID 	WHERE  active = 1 and (TBLReviews.ProductID = ?) AND (TBLReviews.status = N'accepted') AND (ProductDetails.ProductDetail1 IS NOT NULL) GROUP BY ProductDetails.ProductDetail1"
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))

    end if
end if ' if photos

if request("type") = "photos" then
    if request("filter") = "gauge" then
        var_type = "photos-color"
        var_select_type = "photos"
        var_select_by = "gauge"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.Gauge AS gauge, COUNT(*) AS gauge_total FROM TBL_GaugeOrder INNER JOIN ProductDetails ON TBL_GaugeOrder.GaugeShow = ProductDetails.Gauge RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE active = 1 and (ProductDetails.Gauge IS NOT NULL) AND TBL_PhotoGallery.ProductID = ? AND TBL_PhotoGallery.status = 1 GROUP BY ProductDetails.Gauge, TBL_GaugeOrder.GaugeOrder, TBL_PhotoGallery.ProductID, TBL_PhotoGallery.status ORDER BY TBL_GaugeOrder.GaugeOrder"
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))

    end if

    if request("filter") = "color" then
        var_type = "photos-gauge"
        var_select_type = "photos"
        var_select_by = "color"
        objCmd.CommandText = "SELECT TOP (100) PERCENT ProductDetails.ProductDetail1 AS color, COUNT(*) AS color_total FROM ProductDetails RIGHT OUTER JOIN TBL_PhotoGallery ON ProductDetails.ProductDetailID = TBL_PhotoGallery.DetailID WHERE active = 1 and  (ProductDetails.ProductDetail1 IS NOT NULL) AND (TBL_PhotoGallery.ProductID = ?) AND (TBL_PhotoGallery.status = 1)  GROUP BY ProductDetails.ProductDetail1"
        objCmd.Parameters.Append(objCmd.CreateParameter("ProductID",3,1,12,request("productid")))

    end if
end if ' if photos

end if ' showing all

'response.write "productid: -" & request("productid") & "-  "
'response.write "Gauge: -" & filter_gauge & "-  "
'response.write "Color: -" & server.htmlencode(filter_color) & "-  "
'response.write "Filter gauge: -" & server.htmlencode(filter_gauge) & "-  "
'response.write "sql color: -" & server.htmlencode(sql_color) & "-  "
'response.write "sql gauge: -" & server.htmlencode(sql_gauge) & "-  "
'response.write "Gauge from request: -" & request("guage") & "-  "
'response.write "type: -" & request("type") & "-  "
'response.write "filter: -" & request("filter") & "-  "

Set getCounts = objCmd.Execute()
%>
<option value="All">Filter <%= var_select_type %> by <%= var_select_by %></option>
<option value="All">Show all</option>
<optgroup label=" "></optgroup>
<%
if not getCounts.eof then 
i = 0
while not getCounts.eof

if var_type = "photos-color" or var_type = "reviews-color" then
if i = 0 then %>
<optgroup label="<%= filter_color %> sizes:"></optgroup>
<%
end if
%>
<option value="<%= Server.URLEncode(getCounts.Fields.Item("gauge").Value) %>">
    <%=getCounts.Fields.Item("gauge").Value%>&nbsp;&nbsp;&nbsp;(<%=getCounts.Fields.Item("gauge_total").Value%>&nbsp;<%= var_select_type %>)
</option>
<%
end if ' = photos-color

if var_type = "photos-gauge" or var_type = "reviews-gauge" then
if i = 0 then %>
<optgroup label="<%= filter_gauge %> colors:"></optgroup>
<%
end if
%>
<option value="<%= Server.URLEncode(getCounts.Fields.Item("color").Value) %>">
    <%=getCounts.Fields.Item("color").Value%>&nbsp;&nbsp;&nbsp;(<%=getCounts.Fields.Item("color_total").Value%>&nbsp;<%= var_select_type %>)
</option>
<%
end if ' = photos-color

i = i + 1
getCounts.movenext() 
wend

end if ' recordset not empty
    
DataConn.Close()
Set DataConn = Nothing
%>