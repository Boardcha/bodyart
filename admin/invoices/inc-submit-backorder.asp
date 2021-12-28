<form name="frm-submit-backorder" id="frm-submit-backorder">
					<div class="alert alert-warning">We should have <span class="font-weight-bold" id="bo-qty-instock"></span> <span id="bo-item-title"></span> in stock</div>
					
					<div class="form-group">
					<label class="h6" for="qty">How many do we have on hand:</label>
					<select class="form-control form-control-sm" name="bo-qty" id="bo-qty">
						  <option value="0" selected>0</option>
						  <option value="1">1</option>
						  <option value="2">2</option>
						  <option value="3">3</option>
						  <option value="4">4</option>
						  <option value="5">5</option>
						  <option value="6">6</option>
						  <option value="7">7</option>
						  <option value="8">8</option>
						  <option value="9">9</option>
						  <option value="10">10</option>
					  </select>
					</div>
					
					<h6>Reason for backorder:</h6>
					<div class="custom-control custom-radio">
						<input class="custom-control-input" name="BOReason" type="radio" id="radio" value="our inventory was off and we have none left" checked>
						<label class="custom-control-label" for="radio">Not enough items left in stock</label>
					</div>
					<div class="custom-control custom-radio">
						<input class="custom-control-input" type="radio" name="BOReason" id="radio2" value="the last ones we had did not match well enough to send out">
						<label class="custom-control-label" for="radio2">Last pair did not match</label>
					</div>
					<div class="custom-control custom-radio">
						<input class="custom-control-input" type="radio" name="BOReason" id="radio3" value="the last ones we had were not the right size">
						<label class="custom-control-label" for="radio3">Last ones were not the right size</label>
					</div>
					<div class="custom-control custom-radio">
						<input class="custom-control-input" type="radio" name="BOReason" id="radio5" value="it was broken">
						<label class="custom-control-label" for="radio5">It was broken </label>
					</div>
				  </form>