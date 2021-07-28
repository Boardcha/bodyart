$(document).ready(function(){
 
	function auto_update() {
		$(document).on("change", ".ajax-update input:not('#customer_credit, .tender, #ReturnMailer, .charge_amount, .checkbox-return'), .ajax-update textarea:not('.charge_description, #private_notes'), .ajax-update select:not('.tender')", function() {
	//	$(".ajax-update input:not('#customer_credit, .tender, .charge_amount, .checkbox-return'), .ajax-update textarea:not('.charge_description, #private_notes'), .ajax-update select:not('.tender')").change(function(){
			var column_name = $(this).attr("data-column");
			var column_val = $(this).val();
			var id = $("#main-id").val();
			var detail_id = $(this).attr("data-detailid");
			var productdetailid = $(this).attr("data-productdetailid");
			var productid  = $(this).attr("data-productid");
			var field_name = $(this).attr("name");
			var friendly_name = $(this).attr("data-friendly");		
			
			if ($(this).is(':checkbox')) {
				if ($(this).prop("checked")) { // Get values if it's a checkbox
					column_val = $(this).val();
				} else {
					column_val = $(this).attr("data-unchecked");
				}
			}
			
			var $this = $(this);
			if ($this.is("input")) {
				var field_type = "input"
			} else if ($this.is("select")) {
				var field_type = "select"
			} else if ($this.is("textarea")) {
				var field_type = "textarea"
			}
				
			$.ajax({
				method: "POST",
				url: "invoices/ajax_edit_invoices.asp",
				data: {id: id, column: column_name, value: column_val, detailid: detail_id, friendly_name:friendly_name, productdetailid: productdetailid, productid: productid}
				})
				.done(function( msg ) {
					// Highlight field green for success
					
					$(field_type + "[name='"+ field_name +"']").addClass("alert-success");
				//	$("#updated-text").prependTo(field_type + "[name='"+ field_name +"']");
					setTimeout(function(){
						$(field_type + "[name='"+ field_name +"']").removeClass("alert-success");}, 4000);					
				
				//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
				})
				.fail(function(msg) {
				// Highlight field red for failure
				
					$(field_type + "[name='"+ field_name +"']").addClass("alert-danger");
					setTimeout(function(){
						$(field_type + "[name='"+ field_name +"']").removeClass("alert-danger");
						}, 4000);					
				
				//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
				});
		});
	} // end auto update function	
	
	auto_update();


});