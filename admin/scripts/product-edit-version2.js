// Create redactor underline button plug-in
 $.Redactor.prototype.iconic = function()
    {
        return {
            init: function ()
            {
                var icons = {
					'bold': '<i class="fa fa-format-bold"></i>',
					'italic': '<i class="fa fa-format-italic"></i>',
                    'link': '<i class="fa fa-link"></i>',
					'lists': '<i class="fa fa-format-list-bulleted"></i>'
                };
 
                $.each(this.button.all(), $.proxy(function(i,s)
                {
                    var key = $(s).attr('rel');
 
                    if (typeof icons[key] !== 'undefined')
                    {
                        var icon = icons[key];
                        var button = this.button.get(key);
                        this.button.setIcon(button, icon);
                    }
 
                }, this));
            }
        };
    };

	// Initialize bootstrap popovers
	$(function () {
		$('[data-toggle="popover"]').popover()
	  })

	  // Close popover if clicking outside of it
	  $('body').on('click', function () {
		$('.popover').popover('hide');
	});

// WYSIWYG editor for description box
$("#description").redactor({
	tabKey: false, // allows tab key to move to next form field
    buttons: ['html', 'bold', 'italic', 'underline', 'link', 'lists'],
// OLD	buttons: ['html', 'formatting', 'bold', 'italic', 'link', 'unorderedlist', 'outdent', 'indent'],
	formatting: ['p', 'h3', 'h4', 'h5'],
	plugins: ['iconic', 'source'],

	
	
	
	callbacks: {
	keydown: function(e) // Current bug in redcator II as of Oct 2016 that does not run ajax on blur. Had to add this in. They added it as a bug and will fix it in a future release.
        {
            // tab key
            if (e.which === 9)
            {
                var new_html = (this.code.get());
		var column_name = "description"
		var column_val = new_html
		var id = $(".ajax-update :input[name='product-id']").val();
		var friendly_name = "Product description"
		
	//	console.log("Value: " + new_html + " Product id: " + id)
		
			$.ajax({
			method: "POST",
			url: "products/inc_edit_product.asp",
			data: {id: id, column: column_name, value: column_val, friendly_name:friendly_name}
			})
			.done(function( msg ) {
				// Highlight field green for success
				$('.redactor-layer, .redactor-toolbar').css('background-color', '#c3e6cb  !important');
				setTimeout(function(){
					$('.redactor-layer, .redactor-toolbar').css('background-color', '#fff  !important');
				}, 3000);					
				//	console.log("Success");
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			})
			.fail(function(msg) {
			// Highlight field red for failure
			
                $('.redactor-layer, .redactor-toolbar').css('background-color', '#F6cece  !important');
				setTimeout(function(){
					$('.redactor-layer, .redactor-toolbar').css('background-color', '#fff  !important');
				}, 3000);					
				console.log("Failed");
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			});
            }
        },
	
	blur: function(e)
    {
		var new_html = (this.code.get());
		var column_name = "description"
		var column_val = new_html
		var id = $(".ajax-update :input[name='product-id']").val();
		var friendly_name = "Product description"
		
	//	console.log("Value: " + new_html + " Product id: " + id)
		
			$.ajax({
			method: "POST",
			url: "products/inc_edit_product.asp",
			data: {id: id, column: column_name, value: column_val, friendly_name:friendly_name}
			})
			.done(function( msg ) {
				// Highlight field green for success
				$('.redactor-layer, .redactor-toolbar').addClass("ajax_input_success");
				setTimeout(function(){
					$('.redactor-layer, .redactor-toolbar').addClass("ajax_input_fadeout");
					$('.redactor-layer, .redactor-toolbar').removeClass("ajax_input_success");
					}, 3000);					
					$('.redactor-layer, .redactor-toolbar').removeClass("ajax_input_fadeout");

				//	console.log("Success");
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			})
			.fail(function(msg) {
			// Highlight field red for failure
			
				$('.redactor-editor, .redactor-toolbar').addClass("ajax_input_fail");
				setTimeout(function(){
					$('.redactor-editor, .redactor-toolbar').addClass("ajax_input_fadeout");
					$('.redactor-editor, .redactor-toolbar').removeClass("ajax_input_fail");
				}, 3000);					
				$('.redactor-editor, .redactor-toolbar').removeClass("ajax_input_fadeout");
				
			
				console.log("Failed");
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			});
		
    }	// end blur
	} // end callback
}); 

	// Combine reviews
	$('#combine_reviews').click(function(){
		$('#combine_div').toggle();
	});
	
	$('#combine-submit').click(function(){
		var new_productid = $('#productid').val();
		var old_productid = $('#id-transfer-reviews').val();
	$.ajax({
		method: "POST",
		url: "products/ajax_move_reviews_and_photos.asp",
		data: {new_productid: new_productid, old_productid: old_productid}
		})
		.done(function( msg ) {
            $('#combine_success').show();
		})
		.fail(function(msg) {
			alert('Transfer FAILED');
		});
	});
	
	// delete image
	$("#edit_images_link").click(function(){
		$(this).hide();	
		$('#input-img-description').hide();
		$('.img-thumb-clone').remove();
		var imgid = $(this).attr("data-imgid");
		var filename = $(this).attr("data-filename");
		$("#img_remove").load("products/ajax_delete_image.asp?imgid=" + imgid);	
		load_img_footer();
	});
	
	$("#mainimage").click(function(){
		if ($(this).is(':checked')) {
			$("#img_description").hide();
			//$("#video").hide();
			$(".dz-button").html('Drop <span>MAIN</span> images here or click to upload.');
		}   
	});		
	
	$("#additionalimage").click(function(){
		if ($(this).is(':checked')) {
			$("#img_description").removeClass("d-none").show();
			//$("#video").hide();
			$(".dz-button").html('Drop <span>ADDITIONAL</span> images here or click to upload.');
		}   
	});	
	/*
	$("#video").click(function(){
		if ($(this).is(':checked')) {
			$("#video").removeClass("d-none").show();
			$("#img_description").hide();
		}   
	});	 
	*/
	
	// Dropzone options	
	//$( document ).ready(function() {
		Dropzone.options.frmUpload = {
			autoProcessQueue: false,
			parallelUploads : 3,
			uploadMultiple: true,
			init: function () {
				var myDropzone = this;
		
				// Update selector to match your button
				$("#add_new_image").click(function (e) {
					if (myDropzone.getQueuedFiles().length == 3) { 
						e.preventDefault();
						myDropzone.processQueue();
					}
					else {
						alert("(3) images must be uploaded together!");
					}				
				});
				
				$("#clear_dropzone").click(function (e) {
					myDropzone.removeAllFiles(true);
				});
				
				this.on("sendingmultiple", function(file) {					
					for (var i = 0; i < file.length; i++) {
						 for (var j = 0; j < i; j++) {
							if (file[j].width === file[i].width) { 
								alert('Gotcha! It seems that you forgot to resize one of images. You can not upload images in the same dimensions.');
								myDropzone.removeAllFiles(true);
							}
						 }
					}
				});	
	
				myDropzone.on("processing", function(file) {
				  	if ($("#mainimage").is(':checked')) this.options.url = "products/ajax_add_main_image.asp";
					if ($("#additionalimage").is(':checked')) this.options.url = "products/ajax_add_new_image.asp";	
				});

  				myDropzone.on('sending', function(file, xhr, formData) {
					// Append all form inputs to the formData Dropzone will POST
					var data = $('#frmUpload').serializeArray();
					$.each(data, function(key, el) {
						formData.append(el.name, el.value);
					});
					var productid = $("#add_new_image").attr("data-productid");
					var add_img_description = $('#add_img_description').val();
					formData.append('productid', productid);
					formData.append('add_img_description', add_img_description);
				
				});
				
				myDropzone.on("complete", function(file, response) {
					myDropzone.removeFile(file);
				});	
				
				myDropzone.on('success', function(file, response) {
					if ($("#mainimage").is(':checked')) $("#main_img").attr("src", response);
					if ($("#additionalimage").is(':checked')) load_img_footer();
				});					
			}
		}
	//});

	function load_img_footer() {
		$('#detail_images').text('Loading images...');
		setTimeout(function(){
		$('#detail_images').load('products/ajax_get_image_thumbnails_row_bootstrapped.asp?productid=' + $('#productid').val());
		$('#detail_images').text('');
		$('#add_image_spinner').hide();
		},5000); };
	
	load_img_footer();
	
	// Display larger images on hover over thumbnails in footer
	$(document).on("mouseenter", '.mini-thumb', function() { 
		var img_name = $(this).attr("data-name");
		$('#enlarge_footer_image').show();
		$('#enlarge_footer_image').html('<img style="height:240px;width:auto" src="http://bodyartforms-products.bodyartforms.com/'+ img_name +'">')
	});
	$(document).on("mouseleave", '.mini-thumb', function() { 
		$('#enlarge_footer_image').hide();
	});	
	
	// Assign image to detail (click on add img icon)
	$('.assign_img').click(function(){
	
		if (localStorage.getItem('img_detailid') != $(this).attr('data-id') || localStorage.getItem('img_detailid') === null) {
			localStorage.setItem("img_detailid", $(this).attr('data-id'));
			$('.thumb-activate').addClass('thumb-selection');
			$('.assign_img').not(this).removeClass("btn-success");
			$(this).toggleClass("btn-success");
		} else { // 
			localStorage.removeItem("img_detailid");
			$(this).removeClass("btn-success");
			$('.thumb-activate').removeClass('thumb-selection');			
		}
	});
	
	// Assign image to detail (select thumbnail)
	$(document).on("click", '.thumb-selection', function(event) { 
		var imgid = $(this).attr("data-imgid");
		var detailid = localStorage.getItem('img_detailid');

		$("#img_remove").load("products/ajax_img_assignto_detail.asp?imgid=" + imgid + "&detailid=" + detailid + "");

		if (imgid === '0') {
			$('#img_' + detailid).html('');
		} else {
			$('#img_' + detailid).html($(this).clone().removeClass('thumb-selection'));
		}
		
		localStorage.removeItem("img_detailid");
		$('.assign_img').removeClass("btn-success");
		$('.thumb-activate').removeClass('thumb-selection');
	});
	
	// Display description update field when clicking an image thumbnail
	$(document).on("click", '.mini-thumb', function() { 
		$('.mini-thumb').css('border', "0");  
		$(this).css('border', "2px dotted red");  
		$("#edit_images_link").show();
		var imgid = $(this).attr("data-imgid");
		var img_description = $(this).attr("data-description");

		$('#input-img-description').show();
		$('.img-thumb-clone').html('');
		$(this).clone().appendTo('.img-thumb-clone');
		$('#input-img-description').attr('data-imgid', imgid);
		$('#edit_images_link').attr('data-imgid', imgid);
		var filename = $(this).attr('src').split('/');
		filename = filename[filename.length-1];
		$('#edit_images_link').attr('data-filename', filename);
		$('#input-img-description').val(img_description);
	});	

	// Display description update field when clicking an image thumbnailUpdate mini thumbnail image description
	$(document).on("change", '#input-img-description', function() { 
		var imgid = $(this).attr("data-imgid");
		var img_description = $(this).val();
		
		$.ajax({
		method: "POST",
		url: "products/ajax_update_image_description.asp",
		data: {imgid: imgid, img_description: img_description}
		})
		.done(function( msg ) {
			// Highlight field green for success
				$('#input-img-description').addClass("ajax_input_success");
				setTimeout(function(){
					$('#input-img-description').addClass("ajax_input_fadeout");
					$('#input-img-description').removeClass("ajax_input_success");
					}, 3000);					
					$('#input-img-description').removeClass("ajax_input_fadeout");
		})
		.fail(function(msg) {
			// Highlight field red for failure
			
				$('#input-img-description').addClass("ajax_input_fail");
				setTimeout(function(){
					$('#input-img-description').addClass("ajax_input_fadeout");
					$('#input-img-description').removeClass("ajax_input_fail");
				}, 3000);					
				$('#input-img-description').removeClass("ajax_input_fadeout");
		});

	});	
	
	
	$("select[name='tags']").chosen({
		placeholder_text_multiple: "Select tags...",
		inherit_select_classes: true
	});

	$("select[name='category']").chosen({
		placeholder_text_multiple: "Select categories...",
		inherit_select_classes: true
	});

	$("select[name='piercing_type']").chosen({
		placeholder_text_multiple: "Select piercing types...",
		inherit_select_classes: true
	});

	$("select[name*='materials']").chosen({
		placeholder_text_multiple: "Select materials...",
		inherit_select_classes: true
	});

	$("select[name*='colors']").chosen({
		placeholder_text_multiple: "Select colors...",
		inherit_select_classes: true
	});

	$("select[name='threading']").chosen({
		placeholder_text_multiple: "Select threading...",
		inherit_select_classes: true
	});

	$("select[name='flares']").chosen({
		placeholder_text_multiple: "Select flare types...",
		inherit_select_classes: true
	});


	// Change active / inactive drop down select colors
	$("select[name=active]").change(function(){
		if ($(this).val() == '1') {
			$(this).addClass('alert-success');
			$(this).removeClass('alert-danger');			
		} else {
			$(this).addClass('alert-danger');
			$(this).removeClass('alert-success');
		}
	}); // end active selector colors	
	
	// Apply to all materials button
	$("#apply_all_material").click(function(){
		
		var arr_selected_ones_main = $("#materials_main option:selected").map(function() {
			return $(this).text().trim();
		}).get();
		console.log(arr_selected_ones_main);
	
		
		$('.variant-materials').each(function() {
			var arr_selected_ones_variant = $(this).val();
		//	console.log(arr_selected_ones_variant);
			var arrUniqueCompound = arrRemoveDuplicates(arr_selected_ones_main.concat(arr_selected_ones_variant));
		//	console.log(arrUniqueCompound);	
			$(this).val('');
			$(this).val('').trigger("chosen:updated");
			$(this).val(arrUniqueCompound).trigger("chosen:updated");	
			$(this).change();
		//	console.log('--------- next item below');
		});

		$("select[name*='materials']").trigger('chosen:updated');
	
			
	}); // Apply to all materials button
	

	// Apply to all COLORS button
	$("#apply_all_colors").click(function(){
		
		var arr_selected_ones_main = $("#colors_main option:selected").map(function() {
			return $(this).text().trim();
		}).get();
			
		$("select[name*='colors']:not('#colors_main')").each(function() {		

			var arr_selected_ones_variant = $(this).val();
			var arrUniqueCompound = arrRemoveDuplicates(arr_selected_ones_main.concat(arr_selected_ones_variant));
			
			//console.log($(this).val());
			$(this).val('').trigger("chosen:updated");
			$(this).val(arrUniqueCompound).trigger("chosen:updated");
			$(this).change();
		});

	}); // Apply to all COLORS button

	// Remove duplicates from an array
	function arrRemoveDuplicates(arr) {
		var a = [];
		for (var i=0, l=arr.length; i<l; i++)
			if (a.indexOf(arr[i].trim()) === -1 && arr[i].trim() !== '')
				a.push(arr[i].trim());
		return a;
	}
	
	// Apply to all WEARABLE button
		$("#apply_all_wearable_materials").click(function(){
			
			$("select[name*='wearable']:not(#wearable_main) option").removeAttr("selected");
			$( "#wearable_main option:selected" ).each(function() {
				$("select[name*='wearable']:not(#wearable_main) option[value='" + $.trim($(this).text()) + "']").attr("selected", "selected").change();
			});	
	}); // Apply to all WEARABLE button


	
	// Toggle grey row change for active/inactive
	$("input[name^=active_]").change(function(){
		var id = $(this).attr("data-detailid");
			if ($(this).prop("checked")) { // If checked remove inactive styles
                    $('#tbody-' + id).removeClass('table-secondary');
                    $('#tbody-' + id).css({'background-color': ''})
                    $('#tbody-' + id + ' :input').not(':input[type=button]').css({'background-color': 'white', 'border':'solid 1px #ced4da'})
				} else { // if checked style as inactive row
                	$('#tbody-' + id).addClass('table-secondary');
                    $('#tbody-' + id).css({'background-color': '#d6d8db'})
                    $('#tbody-' + id + ' :input').not(':input[type=button]').css({'background-color': '#d6d8db', 'border':'solid 1px #A4A4A4'})
				}	
	}); // Toggle grey row change for active/inactive
	
	
	// START last sold date expand and load ----------------------------------------
    $(".date_expand").click(function(){		
        var detailid = $(this).attr("data-detailid");
       $('.popover').popover('hide');
     //    console.log(detailid);	


		$('.loader-div').load("products/ajax_last_sold_dates.asp", {detailid: detailid}, function() {
            $("#last_sold_" + detailid).attr('data-content', $('.loader-div').html());
            $("#last_sold_" + detailid).popover('show');
        });	
        
        
        
	}); // END last sold date expand and load ----------------------------------------
	
localStorage["move_details"] = "" // set initial value to nothing
localStorage["copy_details"] = "" // set initial value to nothing
	
	// Copy button
	$(".copyid").click(function(){		
		var id = $(this).attr("data-id");
		$('#move-copy-productid').show();
		$('#move-copy-text').html("Copy");		
		$('#move-copy-productid').removeClass("move-detail");
		$('#move-copy-productid').addClass("copy-detail");
		$("span[name=copy_" + id + "]").toggleClass('bg-success');
		$("span[name^=move_]").removeClass('bg-success');
		
		if ($(".bg-success")[0]){} // if class if found do nothing 
		else { // if it's not found then hide span input box
			$('#move-copy-productid').hide();
		}
	
		// check for duplicates and dont' allow them
		if ($(this).hasClass('bg-success')) {
			// check for duplicates and dont' allow them
			if (localStorage.copy_details.indexOf(id) === -1) {
				localStorage["copy_details"] = localStorage.copy_details + id + ","
			}
		} else { // if it's inactive then remove the id from storage
					localStorage["copy_details"] = localStorage["copy_details"].replace(id + ',','');
		}	

	}); // Copy button
	
	// Move button
	$(".moveid").click(function(){		
		var id = $(this).attr("data-id");
		$('#move-copy-productid').show();
		$('#move-copy-text').html("Move");
		$('#move-copy-productid').removeClass("copy-detail");
		$('#move-copy-productid').addClass("move-detail");
		$("span[name=move_" + id + "]").toggleClass('bg-success');		
		$("span[name^=copy_]").removeClass('bg-success');
		
		if ($(".bg-success")[0]){} // if class if found do nothing 		
		else { // if it's not found then hide span input box
			$('#move-copy-productid').hide();
		}
		
		if ($(this).hasClass('bg-success')) {
			// check for duplicates and dont' allow them
			if (localStorage.move_details.indexOf(id) === -1) {
				localStorage["move_details"] = localStorage.move_details + id + ","
			}
		} else { // if it's inactive then remove the id from storage
					localStorage["move_details"] = localStorage["move_details"].replace(id + ',','');
		}	
	}); // Move button
	
	// ------------ When product # is inputted for copy/move then load ajax and redirect to new page
	$("#toggle-productid").change(function(){
		var productid = $("#toggle-productid").val();
		if ($('#move-copy-productid').hasClass('copy-detail')) {
			var toggle_type = "copy" 
			var details = localStorage["copy_details"]
		} else {
			var toggle_type = "move"
			var details = localStorage["move_details"]
		}	

		var section = $("#add-section").val();
				
		//console.log(productid + ", " + toggle_type + ", " + details);
		//console.log("materials: " + materials + " // colors: " + colors)

		$.ajax({
		method: "POST",
		dataType: "json",
		url: "products/ajax_duplicate_move_product.asp",
		data: {move_to_id: productid, details: details, toggle_type: toggle_type, section: section}
		})
		.done(function( json, msg ) {
			window.location.replace("?productid=" + json.productid);
		})
		.fail(function(msg) {
			alert(toggle_type + " FAILED");
		});
	}); // END transferring copied & moved items ---------------------------------------
	
	// Expand all link click
	$(".expand-all").click(function(){	
        $(".expanded-details").toggle();
        $("#btn-expand-all").toggleClass('fa-angle-double-down fa-angle-double-up');
        $('.row-group').toggleClass('border-toggle-on');
	}); // end expand all click
	
	function expand_one() {
		// Expand ONE click
		$(".expand-one").click(function(){	
			var id = $(this).attr("data-id");
            $("." + id).toggle();
            $("#expand_" + id).toggleClass('fa-angle-double-down fa-angle-double-up');
            $('#tbody-' + id).toggleClass('border-toggle-on');
		}); // end expand ONE click
	}
	
	expand_one();
	

	
	// Toggle new
	$('#new-toggle').click(function(){	
		var productid = $(this).attr("data-id");
	
		$('#new-toggle').text(function(i, v){
		   return v === 'Add to new' ? 'Remove from new' : 'Add to new'
		})

	//	var a = $('#new-toggle').attr('class'); 
	//	console.log(a);
		
		// if item is in the new section then we need to remove it
		if ($(this).hasClass('btn-primary')) {
		//	console.log("ACTIVE");
				$.ajax({
				method: "POST",
				dataType: "json",
				url: "products/ajax_new_section_toggle.asp",
				data: {productid: productid, active: "no"}
				})
				.done(function( msg ) {
					$('#new-toggle').removeClass("btn-primary");
					$('#new-toggle').addClass("btn-secondary");
					})
				.fail(function(msg) {
					alert("FAILED");
				});
		} // if item is not in the new section then we need to add it
		if ($(this).hasClass('btn-secondary')) {
		//	console.log("INACTIVE");
				$.ajax({
				method: "POST",
				dataType: "json",
				url: "products/ajax_new_section_toggle.asp",
				data: {productid: productid, active: "yes"}
				})
				.done(function( msg ) {
					$('#new-toggle').removeClass("btn-secondary");
					$('#new-toggle').addClass("btn-primary");
				})
				.fail(function(msg) {
					alert("FAILED");
				});
		}			
	}); // End new button toggle
	
	// Duplicate product
	$("#duplicate").click(function(){
		$("#duplicate-show-buttons").toggle();
	});
	
	$("#duplicate-product").click(function(){	
		var productid = $(this).attr("data-id");
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "products/ajax_duplicate_move_product.asp",
		data: {productid: productid, duplicate: "product-only"}
		})
		.done(function( json, msg ) {
			window.location.replace("?productid=" + json.productid);
		})
		.fail(function(msg) {
			alert("COPY FAILED");
		});
	}); // end Duplicate product
	
	// ------ START duplicate product + details ---------------------
	$("#duplicate-all").click(function(){	
		var productid = $(this).attr("data-id");
		$.ajax({
		method: "POST",
		dataType: "json",
		url: "products/ajax_duplicate_move_product.asp",
		data: {productid: productid, duplicate: "all"}
		})
		.done(function( json, msg ) {
			window.location.replace("?productid=" + json.productid);
		})
		.fail(function(msg) {
			alert("COPY FAILED");
		});
	}); // ------ END duplicate product + details ---------------------


	
	function auto_update() {
		// Auto-update form fields
           $(".ajax-update input:not('.origqty, #img_90, #img_400, #img_1000, #additionalimage, #mainimage, #add_img_thumb, #add_img_description, #id-transfer-reviews, .no_update, #combine_productid, #combine_detailinfo'), .ajax-update textarea:not('#combine_comments'), .ajax-update select:not('#colors_main, #wearable_main')").change(function(){
			var column_name = $(this).attr("data-column");
			var column_val = $(this).val();
			var id = $(".ajax-update :input[name='product-id']").val();
			var detail_id = $(this).attr("data-detailid");
			var field_name = $(this).attr("name");
			var friendly_name = $(this).attr("data-friendly");
			
            //console.log("THIS: " + $(this));
           // console.log("VALUE: " + $(this).val());
           
			// break items out if they are using the tagging select menus
			if (column_name == 'jewelry') {
				var chosen_values = $("select[name='category']").chosen().val();
				// console.log("column: " + column_name + "  Values: " + chosen_values + "  ID: " + id);
				column_val = '' + chosen_values + '';
			}

			// break items out if they are using the tagging select menus
			if (column_name == 'tags') {
				var chosen_values = $("select[name='tags']").chosen().val();
				// console.log("column: " + column_name + "  Values: " + chosen_values + "  ID: " + id);
				column_val = '' + chosen_values + '';
			}
			
			// break items out if they are using the tagging select menus
			if (column_name == 'piercing_type') {
				var piercing_type_values = $("select[name='piercing_type']").chosen().val();
			//	console.log("column: " + column_name + "  Values: " + piercing_type_values + "  ID: " + id);
				column_val = '' + piercing_type_values + '';
			}

			// break items out if they are using the tagging select menus
			if (column_name == 'internal') {
				var threading_values = $("select[name='threading']").chosen().val();
			//	console.log("column: " + column_name + "  Values: " + threading_values + "  ID: " + id);
				column_val = '' + threading_values + '';
			}
			
			// break items out if they are using the tagging select menus
			if (column_name == 'flare_type') {
				var flare_values = $("select[name='flares']").chosen().val();
			//	console.log("column: " + column_name + "  Values: " + threading_values + "  ID: " + id);
				column_val = '' + flare_values + '';
			}						
			
			// break items out if they are using the tagging select menus
			if (column_name == 'material') {
				var materials_main_values = $("select[name='materials_main']").chosen().val();
			//	console.log("column: " + column_name + "  Values: " + materials_main_values + "  ID: " + id);
				column_val = '' + materials_main_values + '';
			}				

			// break items out if they are using the tagging select menus
			if (column_name == 'colors') {
				var colors_values = $("#colors_" + detail_id).chosen().val();
			//	console.log("column: " + column_name + "  Values: " + colors_values + "  ID: " + id);
				column_val = '' + colors_values + '';
			}	

			// break items out if they are using the tagging select menus
			if (column_name == 'detail_materials') {
				var detail_materials_values = $("#materials_" + detail_id).chosen().val();
			//	console.log("column: " + column_name + "  Values: " + detail_materials_values + "  ID: " + id);
				column_val = '' + detail_materials_values + '';
			}				
			
			
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
			//console.log( " PRE AJAX Column: " + column_name + " Value: " + column_val + " ID: " + id + " Detail ID: " + detail_id + " Field name: " + field_name  );	
			
			$.ajax({
				method: "POST",
				url: "products/inc_edit_product.asp",
				data: {id: id, column: column_name, value: column_val, detailid: detail_id, friendly_name:friendly_name}
				})
				.done(function( msg ) {
                    //console.log('SUCCESS - General field update ... success message' + msg);

                    // Highlight field green for success
					$(field_type + "[name='"+ field_name +"'], .select-" + field_name + " .chosen-choices").addClass("alert-success");

                    // If it's a checkbox add the is-valid bootstrap class that makes the label green
                    if ($(field_type + "[name='"+ field_name +"']").is(':checkbox')) {
                        $(field_type + "[name='"+ field_name +"']").addClass("is-valid");
                        console.log("checkbox");
                    }

					setTimeout(function(){					
                        $(field_type + "[name='"+ field_name +"']").removeClass("is-valid");
						$(field_type + "[name='"+ field_name +"'], .select-" + field_name + " .chosen-choices").removeClass("alert-success");
                        }, 4000);			
				})
				.fail(function(msg) {
                    //console.log('FAILED - General field update ... error message' + msg);
				    // Highlight field red for failure
                    $(field_type + "[name='"+ field_name +"'], .select-" + field_name + " .chosen-choices").addClass("alert-danger");

                    // If it's a checkbox add the is-valid bootstrap class that makes the label red
                    if ($(field_type + "[name='"+ field_name +"']").is(':checkbox')) {
                        $(field_type + "[name='"+ field_name +"']").addClass("is-invalid");
                        console.log("checkbox");
                    }

					setTimeout(function(){
                        $(field_type + "[name='"+ field_name +"']").removeClass("is-invalid");
                        $(field_type + "[name='"+ field_name +"'], .select-" + field_name + " .chosen-choices").removeClass("alert-danger");
						}, 4000);					
				});
		});
	} // end auto update function
	
	function update_qty() {
	// START auto update qty field ---------------------------------
	$('.origqty').change(function(){
        console.log("Updating qty");
		var detailid = $(this).attr("data-detailid");
		var origqty = $(this).attr("data-origqty");
		var qty = $(this).val();
		var id = $(".ajax-update :input[name='product-id']").val();		

		$.ajax({
			method: "POST",
			dataType: "json",
			url: "products/ajax_update_qty.asp",
			data: {detailid: detailid, qty: qty, origqty: origqty, id: id}
			})
			.done(function(json,msg) {
				// Highlight field green for success		
				$("input[name='qty-onhand_" + detailid + "']").addClass("alert-success");
				$("input[name='qty-onhand_" + detailid + "']").attr('data-origqty', qty); // update origqty to new qty value entered
				setTimeout(function(){
					$("input[name='qty-onhand_" + detailid + "']").removeClass("alert-success");
					}, 3000);					
					
				if (json.status != "") {
				//	console.log(json.status)
					
					$("input[name='qty-onhand_" + detailid + "']").val(json.difference);
					$("input[name='qty-onhand_" + detailid + "']").attr('data-origqty', json.difference); // update origqty to new qty value entered
					alert(json.status);
				} else { // if console status is good then
				
				}
			//	alert( "error" + msg + "Column: " + column_name + " Value: " + column_val + " ID: " + id + "Detail ID: " + detail_id  );
			})
			.fail(function(msg) {
				// Highlight field red for failure			
				$("input[name='qty-onhand_" + detailid + "']").addClass("alert-danger");
				setTimeout(function(){
					$("input[name='qty-onhand_" + detailid + "']").removeClass("alert-danger");
                    }, 3000);
				alert("QTY UPDATE FAILED");
			});  
		});
	} // END auto update qty field ---------------------------------
	
	auto_update();
	update_qty();

	// Button press to clear variant fields
	$(".btn-clear-fields").click(function () {
		update_field = $(this).attr("id");

		if (update_field == 'clear-variant-materials') {
			$('#materials_main').val('').trigger('chosen:updated');
			$('#materials_main').val('').change();
			$(".variant-materials").val('').trigger("chosen:updated");
			$(".variant-materials").val('').change();
		}
		if (update_field == 'clear-variant-wearable') {
			$('#wearable_main').val('').trigger('chosen:updated');
			$('#wearable_main').val('').change();
			$(".variant-wearable").val('').trigger("chosen:updated");
			$(".variant-wearable").val('').change();
		}
		if (update_field == 'clear-variant-colors') {
			$('#colors_main').val('').trigger('chosen:updated');
			$('#colors_main').val('').change();
			$(".variant-colors").val('').trigger("chosen:updated");
			$(".variant-colors").val('').change();
		}
	});

	// Add a new item
	$("#add-detail").submit(function (event) {
        console.log("Submitting new item")
	//	alert($("#add-detail :input[name='productid']").val());
		$.ajax({
			method: "POST",
			url: "products/inc_add_product.asp",
			data: $("#add-detail").serialize()
		})
		.done(function( msg ) {
		//		var showme = $("#add-detail").serialize();
		//console.log(showme);
			$('#materials_add, #colors_add').val('').trigger('chosen:updated');
			$('#materials_add, #colors_add').val('').trigger('chosen:updated');
			add_id =	$('#productid').val();
			$(".loader-div").load("products/inc_add_retrieve_row.asp?productid=" + add_id, function() {
			$($(this).html()).insertAfter('#display-new-row');
			
			auto_update();
			expand_one();
			
			$('#details-table > tbody:eq(2)').addClass("ajax_input_success");
			setTimeout(function(){			
				$('#details-table > tbody:eq(2)').addClass("ajax_input_fadeout");
				$('#details-table > tbody:eq(2)').removeClass("ajax_input_success");}, 3000);					
			$('#details-table > tbody:eq(2)').removeClass("ajax_input_fadeout");
		});
		
			$("#add-detail")[0].reset();
		})
			.fail(function( msg ) {
			alert( "FAILED: " + msg );
		});
		
		event.preventDefault(); // Prevent form from submitting
	}); // Add a new item
	
	
	// Combine now button clicked
	$("#combine_now").click(function () {
		var new_product_id = $('#productid').val();
		var old_productid = $("#combine_productid").val();
		var detailinfo = $("#combine_detailinfo").val(); 

		$.ajax({
		method: "POST",
		url: "products/ajax_combine_products.asp",
		data: {new_product_id: new_product_id, old_productid: old_productid, detailinfo: detailinfo}
		})
		.done(function( msg ) {
			location.reload();
		//	window.open("product-edit.asp?ProductID=" + old_productid);
		})
		.fail(function(msg) {
			alert("COMBINE FAILED");
		});
	});	
	
	/*
	// Load more
	$("#load_more").click(function () {
		product_id =	$('#productid').val();
		$.get("products/inc_retrieve_rows_load_more.asp?action=>&productid=" + product_id, function(data) {
			$(data).appendTo("#load_more_content").fadeIn("slow");
		});
	}); // Load more
	
	*/

	// Hide all select options where value = &nbsp;
	$("select > option[value='&nbsp;']").remove();
	
	// Auto submit filter form after changing select menus
	$('#filter_active, #filter_gauge').change(function() {
        $('#frm_filters').submit();
	});
	
	//	Sorting image function
	$(function () {
    $("#detail_images").sortable({
		items: 'img',
        update: function (event, ui) {
			var sort_array = $("#detail_images").sortable("toArray");

			$.ajax({
			url: "products/ajax-update-image-sorting.asp",
			method: "POST",
			data: {id_array:JSON.stringify(sort_array)},
			})
			.done(function( msg ) {
				$('#sort-message').html("Sort saved").show().delay(3000).fadeOut( "slow");
			})
			.fail(function(msg) {
				alert("Sort failed");
			});	
        }
    }).disableSelection();
	}); // End sorting function
	
	
	// Apply all button & save
	$(".applyall").click(function () {
		var column = $(this).attr("data-column");
		var column_field = $(this).attr("data-field");
		var column_value = $('#' + column_field).val();
		var productid = $('#productid').val();

		
		// only do something if field is not empty
		if (column_value != '') {
		$.ajax({
		method: "POST",
		url: "products/ajax-applyall-updates.asp",
		data: {column: column, column_value: column_value, productid: productid}
		})
		.done(function( msg ) {
			location.reload();
		})
		.fail(function(msg) {
			alert("ERROR");
		});
		} // end only do something if field is not empty
	});		// End apply all button & save
	
	
	
	// Throw a notice if the retail price is less than double wholesale
	$('.check-wholesale').change(function() {
        var item = $(this).attr("data-pricecheck");
        console.log($('.pricecheck_retail_' + item).val() + " , " + ($('.pricecheck_wlsl_' + item).val() * 2));
		if ($('.pricecheck_retail_' + item).val() < ($('.pricecheck_wlsl_' + item).val() * 2 - .05) ) {
			$('<div class="alert alert-danger bold">Retail is less than double wholesale</div>').insertBefore(this).delay(10000).fadeOut();
			console.log($('.pricecheck_retail_' + item).val() + " , " + ($('.pricecheck_wlsl_' + item).val() * 2));
		}

		
	});
	
	// If ispreorder field checked then show extra field options
	$('.ispreorder').change(function() {
        if (this.checked) {
			$('.preorder-fields').show();	
		} else {
			$('.preorder-fields').hide();
		}
	});	

// Change SEO title field after title field update
$("#title").blur(function () {
    var title = $('#title').val();
    var seo_title = $("#seo_meta_title").val();

    $("#seo_meta_title").val(title);
    $("#seo_meta_title").trigger('change');
});	

// Duplicate title into SEO title field and then check to see if it's a duplicate
$("#seo_meta_title").change(function () {
    var seo_title = $("#seo_meta_title").val();
    var productid = $("#productid").val();

    $.ajax({
    method: "POST",
    dataType: "json",
    url: "products/ajax-check-duplicate-title.asp",
    data: {seo_title: seo_title, id: productid}
    })
    .done(function( json, msg ) {
        if (json.status === "fail") {
            $("#msg-seo-title").html('Duplicate title found... Needs to be updated')
            $('#msg-seo-title').addClass("notice-red");
        } else {
            $("#msg-seo-title").html('')
            $('#msg-seo-title').removeClass("notice-red");
        }
    })
    .fail(function(msg) {
        $("#msg-seo-title").html('Error checking duplicate')
        $('#msg-seo-title').addClass("notice-red");
    });

});	

// View edits log
$(document).on("click", '.view-edits-log', function(event) { 
	var detailid = $(this).attr("data-detailid");
	$("#show-edits").load("products/ajax-edits-log.asp?detailid=" + detailid);
});

// Popup opener for managing tags
$("#manage_tags").click(function () {
	$("#show-tags").load("products/manage_tags.asp");
});

// Delete product tag
$(document).on("click", '.delete-tag', function(event) { 
	$.ajax({
		method: "GET",
		url: "products/ajax_manage_tags.asp",
		data: {tagID: $(this).attr("data-tag-id"), deleteTag:"yes"}
	})
	.done(function( data ) {
		$("#show-tags").load("products/manage_tags.asp");
		// Preserve tags already selected and retrieve updated tags from the table TBL_Products_Table
		$("#select-tags").find('option').not(':selected').remove();
		$("#select-tags").append(data);
		$("#select-tags").trigger("chosen:updated");
	})
});		

// Add new product tag
$(document).on("click", '#add_new_tag', function(event) { 
	$.ajax({
		method: "GET",
		url: "products/ajax_manage_tags.asp",
		data: {tag: $("#tag").val(), addTag:"yes"}
	})
	.done(function( data ) {
		$("#show-tags").load("products/manage_tags.asp");
		// Preserve tags already selected and retrieve updated tags from the table TBL_Products_Table
		$("#select-tags").find('option').not(':selected').remove();
		$("#select-tags").append(data);
		$("#select-tags").trigger("chosen:updated");
	});
});	

// Review button
$(document).on("click", '#reviewed', function() { 
    var productid = $("#productid").val();

	$.ajax({
		method: "post",
		url: "products/ajax-reviewed-product.asp",
		data: {productid: productid}
	})
	.done(function() {
		$('#reviewed').hide();
		$('#reviewed-msg').html("Review complete");
	});
});	



