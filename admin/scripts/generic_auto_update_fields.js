// auto_url -- variable that each page will set to tie into this feature for load

function auto_update() {
	// Auto-update form fields
		$(document).on("change", ".ajax-update input, .ajax-update textarea, .ajax-update select:not([name='password'])", function() {

		var column_name = $(this).attr("data-column");
		var column_val = $(this).val();
		var id = $(this).attr("data-id");
		var field_name = $(this).attr("name");
		var friendly_name = $(this).attr("data-friendly");
		var int_string = $(this).attr("data-int_string");
		var tempid = $("#tempid").val();
		
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
			url: auto_url,
			data: {id: id, column: column_name, value: column_val,  friendly_name:friendly_name, tempid: tempid, int_string: int_string}
			})
			.done(function( msg ) {
				// Highlight field green for success
				$(field_type + "[name='"+ field_name +"']").addClass("alert-success");
				
				setTimeout(function(){					
					$(field_type + "[name='"+ field_name +"']").removeClass("alert-success");
				}, 3000);

				//	console.log("Success");
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			})
			.fail(function(msg) {
			// Highlight field red for failure
			
				$(field_type + "[name='"+ field_name +"']").addClass("alert-danger");
				setTimeout(function(){
					$(field_type + "[name='"+ field_name +"']").removeClass("alert-danger");
				}, 3000);					
				alert("The field did not save. Try again or contact Amanda.");
			//	console.log( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			});
	});
} // end auto update function