<%
Set objCmd = Server.CreateObject ("ADODB.Command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM TBL_Barcodes_SortOrder" 
Set rs_getsections = objCmd.Execute()
%>
<form action="/admin/barcodes_modifyviews.asp" method="post">
  <div class="form-inline mb-3">
<select class="form-control form-control-sm mr-4" name="type" id="type">
  <option value="-">None</option>
  <% While NOT rs_getsections.EOF %>                          
    <option value="<%=(rs_getsections.Fields.Item("ID_Number").Value)%>"><%=(rs_getsections.Fields.Item("ID_Description").Value)%></option>
  <% 
  rs_getsections.MoveNext()
  Wend
  %> 
</select>

<input class="mr-1" type="radio" name="DetailSort" value="Equal" id="type_0" checked="checked"> =
<input class="ml-4 mr-1" type="radio" name="DetailSort" value="Greater" id="type_0"> >
<input class="ml-4 mr-1" type="radio" name="DetailSort" value="GreaterLess" id="type_2"> &lt; &gt; 
<input class="ml-5 mr-1" type="checkbox" name="new" value="yes">Only return details added in the last 60 days
</div>
<div class="my-2">
  Location numbers:
  <div class="form-inline">
    <input class="form-control form-control-sm w-25" name="Details" type="text" id="Details" placeholder= "Example: 123, 456, 789" /> 
    <span class="mx-3">through</span>
    <input class="form-control form-control-sm w-25" name="Details2" type="text" id="Details2" maxlength="6">
  </div>
</div>

Product IDs:
<input class="form-control form-control-sm w-50 mb-3" name="Products" type="text" id="Products" placeholder= "Example: 12345, 67890, 57891">

<button class="btn btn-purple" type="submit">Update query</button>
</form>