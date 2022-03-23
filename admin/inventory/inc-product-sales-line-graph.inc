<script type="text/javascript">
// Line Graph
	$(document).on("click", '.toggle-product-detail', function() {
		var detailID = $(this).attr("data-detailID");
		var row_id = $(this).attr("id");
		$('.product-detail-expand' + row_id).toggle();
		$('.td-expand' + row_id).fadeToggle('fast');
		$('.load' + row_id).load('/admin/inventory/ajax-product-sales-line-graph.asp?detailID=' + detailID);
	});
</script>
<script type="text/javascript" src="https://cdnjs.cloudflare.com/ajax/libs/mdbootstrap/4.5.4/js/mdb.min.js"></script>


<script>
// Tooltip
$(document).ready(function(){
	// Bootstrap Tooltip
	$('[data-toggle="tooltip"]').tooltip();
});
</script>
