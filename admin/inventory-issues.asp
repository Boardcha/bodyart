<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

set objCmd = Server.CreateObject("ADODB.Command")
objCmd.ActiveConnection = DataConn

objCmd.CommandText = "SELECT ProductDetail1, location, Gauge, Length, title, ProductDetails.ProductDetailID, ProductDetails.ProductID, BinNumber_Detail, TBL_Barcodes_SortOrder.ID_Description, issue_fixed, issue_description, issue_reported_by_who, issue_id FROM ProductDetails INNER JOIN TBL_Barcodes_SortOrder ON ProductDetails.DetailCode = TBL_Barcodes_SortOrder.ID_Number INNER JOIN tbl_product_issues ON ProductDetails.ProductDetailID = tbl_product_issues.issue_detailid INNER JOIN jewelry ON ProductDetails.ProductID = jewelry.ProductID WHERE tbl_product_issues.issue_fixed = 0 ORDER BY issue_date_reported ASC"
set rsGetRecords = objCmd.Execute()

set rsGetRecords = Server.CreateObject("ADODB.Recordset")
rsGetRecords.CursorLocation = 3 'adUseClient
rsGetRecords.Open objCmd
%>

<html>
<head>
<title>Review reported inventory issues</title>
<script type="text/javascript" src="../js/jquery-3.3.1.min.js"></script>
<script type="text/javascript" src="../js/bootstrap-v4.min.js"></script>
</head>
<body>

<!--#include file="admin_header.asp"-->
<div class="mx-2">
<% If Session("SubAccess") <> "N" then ' DISPLAY ONLY TO PEOPLE WHO HAVE ACCESS TO THIS PAGE %>

	<% If NOT rsGetRecords.EOF Then %>
		<h5 class="mt-3 mb-2"><%= rsGetRecords.RecordCount %> reported issues</h5>
		<button class="btn btn-sm btn-secondary d-inline-block mb-3" id="update_query_labels" title="Update barcode query for label printing"><i class="fa fa-label fa-lg mr-1"></i> Print requested replacement labels</button>
		<span class="mb-3 ml-1" id="msg-query-update"></span>

		<table class="table table-striped table-hover">
		<thead class="thead-dark">
			<tr>
				<th><button class="btn btn-secondary btn-sm toggle-off" data-issue_id="all">All done</button></th>
				<th width="20%">Item</th>
				<th>Location</th>
				<th>Reported by</th>
				<th width="40%">Reported issue</th>
			</tr>
		</thead>
		<tbody id="row_all"> 
		<% 
		While NOT rsGetRecords.EOF 

		if instr(rsGetRecords("issue_description"), "Print new scanning label") > 0 then
		  detailids = detailids & " OR ProductDetailID = " & rsGetRecords("ProductDetailID")
		end if
		%>
			   
				<tr id="row_<%= rsGetRecords("issue_id") %>">
					<td>
						<button class="btn btn-primary btn-sm toggle-off" data-issue_id="<%= rsGetRecords("issue_id") %>">Done</button>
					</td>
					<td>
						<a class="text-secondary" href="product-edit.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%>&info=less"><%=(rsGetRecords.Fields.Item("title").Value)%></a>&nbsp; <%=(rsGetRecords.Fields.Item("Gauge").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("Length").Value)%></td>
						<td>
							<%=(rsGetRecords.Fields.Item("ID_Description").Value)%>&nbsp;<%=(rsGetRecords.Fields.Item("location").Value)%>&nbsp;
						<% if (rsGetRecords.Fields.Item("BinNumber_Detail").Value) <> 0 then %>
							(BIN <%=(rsGetRecords.Fields.Item("BinNumber_Detail").Value)%>)
						<% end if %>
					</td>
					<td><%=(rsGetRecords.Fields.Item("issue_reported_by_who").Value)%></td>
					<td><span class="text-danger"><%=(rsGetRecords.Fields.Item("issue_description").Value)%></span></td>
				</tr>
			
			<% 
			rsGetRecords.MoveNext()
		  
		Wend
		%>
		</tbody>
		</table>
		<input type="hidden" id="detailids" value="<%= replace(detailids, "OR", "AND",1 , 1) %>">
	<% else ' if there are no records to review %>
		<h5 class="mt-3 mb-2">
			No reported issues
		</h5>
	<% End If ' end rsGetRecords.EOF And rsGetRecords.BOF %>

<% else ' unathorized access error %>
	Not accessible
<% end if ' END ACCESS TO PAGE FOR ONLY USERS WHO SHOULD BE ABLE TO SEE IT %>

</div>
<!--#include file="includes/inc_scripts.asp"-->
<script type="text/javascript">
    // Clear error
    $(document).on("click", ".toggle-off", function(){
        var issue_id = $(this).attr('data-issue_id');
		if (issue_id == "all"){
			if (confirm("Are you sure to set all done?") == true)
				toggleOff(issue_id);
		}else {
			toggleOff(issue_id);
		}
    })

	function toggleOff(issue_id){
        $.ajax({
            method: "post",
            url: "inventory/toggle-off-inventory-issue.asp",
            data: {issue_id: issue_id}
            })
            .done(function(msg) {
                $('#row_' + issue_id).hide();
            })
            .fail(function(msg) {
                
        })	
	}
	
    // BEGIN Alter barcode query for item labels
    $(document).on("click", '#update_query_labels', function() { 
		$.ajax({
			method: "post",
			url: "/admin/barcodes_modifyviews.asp?type=labels_by_detailid",
			data: {detailids: $('#detailids').val()}
		})
		.done(function() {
			$('#msg-query-update').html('<span class="alert alert-success px-2 py-0"><i class="fa fa-check"></i></span>').show().delay(2500).fadeOut("slow");
		});
    });	// END Alter barcode query for item labels

</script>
</body>
</html>
<%
rsGetRecords.Close()
%>
<%
rsGetUser.Close()
Set rsGetUser = Nothing
%>
