<style>
	.popover{ z-index: 9999; }
</style>
<div class="loader-div" style="display:none"></div>

<script type="text/javascript">
	// Initialize bootstrap popovers
	$(function () {
		$('[data-toggle="popover"]').popover()
	  })

	 // Close popover if clicking outside of it
	  $('body').on('click', function () {
		$('.popover').popover('hide');
	});

	$(document).on("click", '.date_expand', function() {
		var detailid = $(this).attr("data-detailid");
		$('.popover').popover('hide');

		$('.loader-div').load("/admin/products/ajax_last_sold_dates.asp", {detailid: detailid}, function() {
			$("#last_sold_" + detailid).attr('data-content', $('.loader-div').html());
			$("#last_sold_" + detailid).popover('show');
		});	 
	});
</script>
