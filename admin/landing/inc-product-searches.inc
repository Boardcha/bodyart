<form class="form-inline" name="invoice_search" action="/admin/product-edit.asp" method="get">
    <input class="form-control form-control" name="ProductID" type="text" placeholder="Product ID #">
    <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
</form>
<form class="form-inline" name="detailid_search" action="/admin/SearchDetailID.asp" method="post">
    <input class="form-control form-control" name="DetailID" type="text" placeholder="Detail ID #">
    <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
</form>
<form class="form-inline" name="location_search" action="/admin/location_search.asp" method="post">
    <select class="form-control mb-1 mr-1" name="section" id="search-section">
        <% While NOT rs_getsections.EOF %>                          
        <option value="<%=(rs_getsections.Fields.Item("ID_Description").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>
      <% 
      rs_getsections.MoveNext()
      Wend
      %> 
  </select>
    <input class="form-control form-control mb-1 mr-2" name="location" type="text" placeholder="Location #">
    <button class="btn btn-sm btn-secondary" type="submit">Search</button>
</form>
<form class="form-inline" name="sku_search" action="/admin/SearchDetailID.asp" method="post">
    <input class="form-control form-control" name="sku" type="text" placeholder="SKU #">
    <button class="btn btn-sm btn-secondary ml-2" type="submit">Search</button>
</form>