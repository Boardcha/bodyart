<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
'==== PAGE HAS BEEN BOOSTRAPPED =======
bootstrapped = "yes"

%>
<!DOCTYPE html>
<html>
<head>
<title>Products that haven't sold in two years or greater</title>
<style type="text/css">
	.sale-field select {
		margin-right: 10px;
	}
	.image-container {
		padding-right: 10px;
	}
	.variant-container {
		
		overflow: auto;
		max-height: 100px;
	}
	.product-notes {
		
	}
</style>
</head>
<body>
<!--#include file="admin_header.asp"-->
<script src="https://cdnjs.cloudflare.com/ajax/libs/handlebars.js/4.7.7/handlebars.min.js" integrity="sha512-RNLkV3d+aLtfcpEyFG8jRbnWHxUqVZozacROI4J2F1sTaDqo1dPQYs01OMi1t1w9Y2FdbSCDSQ2ZVdAC8bzgAg==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
<script type="text/javascript">
	$(document).ready(function() {
		// initial variables
		var page = 0;
		var totalPages = 0;
		var rowTemplate = $('#row-template tbody').html();
		var compiledTemplate = Handlebars.compile(rowTemplate);
		var moreToFetch = true;
		var fetching = false;

		var scrollHandler = function (e) {
			var $target = window;

			var scrollTop = $(window).scrollTop();
			var height = $(window).height();
			var documentHeight = $(document).height();

			// load more if hit the bottom of the page
			if (Math.ceil($(window).scrollTop() + $(window).height()) >= $(document).height() && moreToFetch) {
				page += 1;
				fetchPage();
			}
		};

		// $(window).scroll(scrollHandler);

		var $filterSelect = $('#filter-select');
		var $pageSelect = $('.page-select');

		var updatePagination = function () {
			$pageSelect.children().removeClass(['disabled', 'active']);
			// do we disable previous?
			if (page <= 1) {
				$pageSelect.find('.previous-btn').addClass('disabled');
			}

			if (page == totalPages) {
				$pageSelect.find('next-btn').addClass('disabled');
			}

			$pageSelect.find('li[data-page=' + page + ']').addClass('active');
		};

		var changePage = function (e) {
			e.preventDefault();
			var $target = $(e.target);
			if ($target.hasClass('disabled')) {
				return;
			}

			if ($target.hasClass('previous') || $target.parent().hasClass('previous')) {
				if (page > 0) {
					page -= 1;
					fetchPage();
				}
				return;
			}

			if ($target.hasClass('next') || $target.parent().hasClass('next')) {
				if (page < totalPages) {
					page += 1;
					fetchPage();
				}
				return;
			}

			page = parseInt($target.text());
			fetchPage();
		};

		var fetchCount = function () {
			$('.data-container').children().remove();
			$.ajax({
				url: './ajax/ajax-legacy.asp',
				method: 'POST',
				data: {
					action: 'count',
					filter: $filterSelect.val()
				},
				error: function (data) {
					console.error('Could not fetch total count', data);
				},
				success: function (data) {
					$pageSelect.children().remove();

					totalPages = Math.ceil(data.count / 100);

					var $pageTemplate = $('.page-template');

		  			page = 0;		
					$pageTemplate.find('.previous-btn')
						.clone()
						.click(changePage)
						.appendTo($pageSelect);

					for (var i = 1; i <= totalPages; i++) {
						var $el = $('<li />').attr('data-page', i).addClass('page-item');
						$('<a />', { href: '#', text: i}).addClass('page-link').appendTo($el);
						$el.click(changePage).appendTo($pageSelect);
					}

					$pageTemplate.find('.next-btn')
						.clone()
						.click(changePage)
						.appendTo($pageSelect);
				}
			});
		};

		var fetchPage = function () {
			if (fetching) return;
			$('.data-container').html('<tr><td colspan="4" class="text-center"><i class="fa fa-spinner fa-2x fa-spin"></i></td></tr>');
			
			// state flag for blocking multiple requests
			fetching = true;
			updatePagination();
			$.ajax({
				url: './ajax/ajax-legacy.asp',
				method: 'POST',
				data: {
					page: page,
					filter: $filterSelect.val()
				},
				error: function (data) {
					fetching = false;
				},
				success: function (data) {
					fetching = false;

					// if the spinner is there, we clear the entire element
					if ($('.data-container .fa-spinner').length) {
						$('.data-container').children().remove();
					}

					if (data.length < 100) {
						// we didn't get a full page (or we got nothing)
						// so there's no more to fetch
						moreToFetch = false;
					}
					var productIds = [];
					data.forEach(function (item) {
						item.rowClass = 'table-danger';
						if (item.onsale == 'Y') {
							item.rowClass = 'table-success';
						}
						productIds.push(item.ProductID);
						item.saledisplay = (item.onsale == 'Y' ? 'Yes' : 'No') + ' (' + item.salediscount + '%)';
						var $templateItem = $(compiledTemplate(item));

						var productId = item.ProductID;
						$templateItem.find('.status-field').data('status', item.type);
						$templateItem.find('.status-field').click(function (e) {
							if ($(this).data('editing')) {
								return;
							}

							var $target = $(e.currentTarget);
							var $statusSelect = $('.status-template #status-select').clone();
							$statusSelect.find('[value=' + item.type + ']').attr('selected', 'true');

							var $statusField = $(this);

							var $statusSaveBtn = $('<div />').addClass('btn btn-primary').text('Save').click(function(e) {
								var newStatus = $statusSelect.val();
								$.ajax({
									url: './ajax/ajax-legacy.asp',
									method: 'POST',
									data: {
										action: 'status',
										productId: productId,
										status: newStatus,
										user: '<%=(rsGetUser.Fields.Item("name").Value)%>'
									},
									error: function (data) {
										// TODO: error, probably bootstrap toast
										console.error(data);
									},
									success: function (data) {
										if (data.status == "OK") {
											// we successfully updated
											$statusField.text(newStatus);
											$statusField.data('status', newStatus);
											$statusField.removeData('editing');
											
										}
									}
								});
							});

							$target.html('');
							$target.append($statusSelect, $statusSaveBtn);
							$(this).data('editing', true);

						});
						$templateItem.find('.sale-field').data('sale', item.salediscount);
						$templateItem.find('.sale-field').click(function (e) {
							if ($(this).data('editing')) {
								return;
							}
							var $target = $(e.currentTarget);
							var $saleSelect = $('<select />').addClass('form-control-sm');

							for (i = 0; i <= 90; i += 5) {
								$('<option />', {value: i, text: i + '%'}).appendTo($saleSelect);
							}

							var $saleField = $(this);
							$saleSelect.find('[value=' + $saleField.data('sale') + ']').attr('selected', 'true');

							

							var $saveBtn = $('<div />').addClass('btn btn-primary').text('Save').click(function(e) {
								// document.location.href = 'legacy.asp?product=' + productId + '&sale=' + $saleSelect.val()
								var saleAmount = $saleSelect.val();
								$.ajax({
									url: './ajax/ajax-legacy.asp',
									method: 'POST',
									data: {
										action: 'sale',
										productId: productId,
										amount: saleAmount,
										user: '<%=(rsGetUser.Fields.Item("name").Value)%>'
									},
									error: function (data) {
										// TODO: error, probably bootstrap toast
										console.error(data);
									},
									success: function (data) {
										if (data.status == "OK") {
											// we successfully updated
											$saleField.text(((saleAmount > 0) ? 'Yes' : 'No') + ' (' + saleAmount + '%)');
											$saleField.data('sale', saleAmount);
											$saleField.removeData('editing');
											$saleField.parent().removeClass('table-danger').removeClass('table-success').addClass((saleAmount > 0) ? 'table-success' : 'table-danger');
										}
									}
								});
							});;
							$target.html('');
							$target.append($saleSelect, $saveBtn);
							$(this).data('editing', true);


						});
						$('.data-container').append($templateItem);
					});
					// fetch variants for product Id's
					$.ajax({
						url: './ajax/ajax-legacy.asp',
						method: 'POST',
						data: {
							action: 'variants',
							productIds: productIds.join('|')
						},
						error: function (data) {
							// TODO: error, probably bootstrap toast
							console.error(data);
						},
						success: function (data) {

							data.forEach(function (variant) {
								var $product = $('[data-productid=' + variant.ProductID + ']');

								if (!$product.length) {
									console.error('Could not find product ', variant.ProductID);
									return;
								}

								var $variantContainer = $product.find('.variant-container');

								// if spinner is still there, clear it
								if ($variantContainer.find('.fa-spinner')) {
									$variantContainer.find('.fa-spinner').parent().remove();
								}

								var variantDetail = variant.Gauge + ' ' + variant.Length + ' ' + variant.ProductDetail1 + ' - ' + variant.DateLastPurchased;
								$('<li />', {text: variantDetail}).appendTo($variantContainer);
							});

							var remainingLoaders = $('.data-container').find('.fa-spinner');
							if (remainingLoaders.length) {
								remainingLoaders.parent().remove();
							}
						}
					})
				}
			});
		};

		fetchCount();
		$filterSelect.on('change', fetchCount);
		
	});
</script>
<div class="page-template" style="display: none">
	<li class="page-item disabled previous-btn previous">
		<a class="page-link previous" href="#" aria-label="Previous">
		  <span aria-hidden="true">&laquo;</span>
		  <span class="sr-only">Previous</span>
		</a>
	</li>
	<li class="page-item disabled next-btn next">
		<a class="page-link next" href="#" aria-label="Next">
		  <span aria-hidden="true">&raquo;</span>
		  <span class="sr-only">Next</span>
		</a>
	</li>
</div>
<div class="status-template" style="display: none">
	<select id="status-select" class="form-control-sm">
		<option value="None">None</option>
		<option value="limited">limited</option>
		<option value="clearance">clearance</option>
		<option value="One time buy">One time buy</option>
		<option value="Discontinued">Discontinued</option>
	</select>
</div>
<table id="row-template" style="display: none">
	<tr data-productid="{{ProductID}}" class="{{rowClass}}">
		<td scope="row"  style="width: 50%">
			<div class="image-container float-left">
				<a class="font-weight-bold" href="product-edit.asp?ProductID={{ProductID}}" target="_blank"><img src="https://bodyartforms-products.bodyartforms.com/{{picture}}" /></a>
			</div>
			
			<div>
				<div class="font-weight-bold product-notes">{{ProductNotes}}</div>
				<ul class="variant-container">
					<li style="list-style-type: none"><i class="fa fa-spinner fa-spin"></i></li>
				</ul>
			</div>
			
		</td>
		<td>Last purchase: {{LastPurchaseDate}}
			<div class="mt-3">Oldest purchase: {{OldestPurchaseDate}}</div></td>
		<td class='status-field' data-status="{{type}}">{{type}}</td>
		<td class='sale-field'>
		{{saledisplay}} 
		</td>
		
	</tr>
</table>
<div class="p-3">
	<div class="form-group form-inline">
		<label for="filter-select">Show products that haven't sold in</label>
		<select class="ml-3 form-control" id="filter-select">
		  <option selected value="6">6 months</option>
		  <option value="9">9 months</option>
		  <option value="12">1 year</option>
		  <option value="24">2 years</option>
		  <option value="36">3 years</option>
		  <option value="never">Never</option>
		</select>
	</div>
	<nav aria-label="Page navigation">
		<ul class="page-select pagination justify-content-center">
		  
		</ul>
	</nav>
</div>

<div class="p-3">
	<div class="container-fluid p-0 mt-4">
		<div class="row">
		  <div class="col">
			<table class="table table-striped table-hover table-sm table-bordered">
				<thead class="thead-dark">
				  <tr>
					<th scope="col">Product</th>
					<th scope="col">Purchase Info</th>
					<th scope="col">Status</th>
					<th scope="col">Sale</th>
				  </tr>
				</thead>
				<tbody class='data-container'>
					
				</tbody>
			  </table>
		  </div>
	</div>
</div>
<div class="p-3">
	<nav aria-label="Page navigation">
		<ul class="page-select pagination justify-content-center">
		  
		</ul>
	</nav>
</div>

</body>
</html>