	
		<div class="card-fields">
		<div class="card-images">
			<i class="fa fa-cc-visa fa-2x"></i>
			<i class="fa fa-cc-mastercard fa-2x"></i>
			<i class="fa fa-cc-amex fa-2x"></i>
			<i class="fa fa-cc-discover fa-2x"></i>
		</div>
		
		<div class="form-group">
		<label for="cardNumber">Card number <span class="text-danger">*</span></label>
		<input class="form-control" required type="tel" name="card_number" id="cardNumber" value="<%= session("card_number") %>" placeholder="&#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;  &#8226;&#8226;&#8226;&#8226;" data-validation="length alphanumeric"  data-validation-length="12-19" data-validation-allowing=" " autocomplete="cc-number" /><span id="current_card_number"></span>
		<div class="invalid-feedback">
				Valid credit card # is required
		</div>
		</div>
		<div class="form-row">
			<div class="col">
		<label for="creditCardMonth">Exp month <span class="text-danger">*</span></label>
		<select class="form-control" required name="billing-month" id="creditCardMonth" name="billing-month" autocomplete="cc-exp-month" >
				  <option value="">Select month</option>
				  <option value="01">01 - January</option>
				  <option value="02">02 - February</option>
				  <option value="03">03 - March</option>
				  <option value="04">04 - April</option>
				  <option value="05">05 - May</option>
				  <option value="06">06 - June</option>
				  <option value="07">07 - July</option>
				  <option value="08">08 - August</option>
				  <option value="09">09 - September</option>
				  <option value="10">10 - October</option>
				  <option value="11">11 - November</option>
				  <option value="12">12 - December</option>
				</select>
				<div class="invalid-feedback">
						Exp month is required
				</div>
			</div>
			<div class="col">
		<label for="creditCardYear">Exp year <span class="text-danger">*</span></label>
		<select class="form-control" required name="billing-year" id="creditCardYear" autocomplete="cc-exp-year">
				  <option value="">Select year</option>
				<% for i = 0 to 10 %>
				  <option value="<%= year(now) + i %>"><%= year(now) + i %></option>
				<% next %>
		</select>
		<div class="invalid-feedback">
				Exp year is required
		</div>
	</div>
		</div>
	</div><!-- card fields -->	
	<div class="address-fields">
		<div class="form-row my-3">
			<div class="col">
		<label for="first">First name <span class="text-danger">*</span></label>
		<input class="form-control" required name="first" id="first" type="text" value="" autocomplete="shipping given-name" />
		<div class="invalid-feedback">
				First name is required
		</div>
		</div>
		<div class="col">

		<label for="last">Last name <span class="text-danger">*</span></label>
		<input class="form-control" required name="last" id="last" type="text" value="" autocomplete="shipping family-name" />
		<div class="invalid-feedback">
				Last name is required
		</div>

		</div>
		</div>
		
		<section data-pg-verify>
		<div id="shipping-address-autocomplete">
			<div class="form-group position-relative">
				<label for="shipping-full-address">Address<span class="text-danger">*</span></label>
				<input type="text" id="shipping-full-address" data-pg-full-address  class="form-control" placeholder="Start typing an a&#8203;ddress..."  autocomplete="off" />
			</div>
		 </div>
		 <div class="form-group position-relative" id="chk-shipping-manual-address-input-container">
			<div class="custom-control custom-checkbox">
				<input type="checkbox" class="custom-control-input" name="chk-shipping-manual-address-input" id="chk-shipping-manual-address-input">
				<label class="custom-control-label" for="chk-shipping-manual-address-input">You can't find the address? Enter the address manually.</label>
			</div>
		</div>		 

		<div id="selected-shipping-address" class="mt-3" style="display:none"></div>

		<div id="shipping-address-container" style="display:none">			
	
			<div class="form-group">
			<label for="address">Address (Line 1) <span class="text-danger">*</span></label>
			<input data-pg-address-line1 class="form-control" required name="address" id="address" type="text" value="<% if var_add_only = "" then %><%= session("address") %><% end if %>" autocomplete="shipping address-line1" />
			<div class="invalid-feedback">
					Address is required
			</div>
			</div>
			<div class="form-group">
			<label for="address2">Apt #, Dorm, Suite&nbsp;&nbsp;</label>
			<input data-pg-address-line2 class="form-control" name="address2" id="address2" type="text" value="<% if var_add_only = "" then %><%= session("address2") %><% end if %>" autocomplete="shipping address-line2" />
			</div>
			<div class="form-group">
			<label for="city">City <span class="text-danger">*</span></label>
			<input data-pg-city class="form-control" required name="city" id="city" type="text" autocomplete="shipping address-level2" />
			<div class="invalid-feedback">
					City is required
			</div>
			</div>
			
			<div class="form-group">                     
			<label for="country">Country <span class="text-danger">*</span></label>
			<select data-pg-country class="form-control" required name="country" id="country" data-validation="required" autocomplete="shipping country" <% if var_update_order_address = "yes" then %>disabled<% end if %>>
			<% 
			if session("country") <> "" and var_add_only = "" then %>
			<option value="<%= session("country") %>" ><%= session("country") %></option>
			<% end if %>
			<option value="USA">USA</option>
			<% 
			While NOT rsGetCountrySelect.EOF 
			%>
			<option value="<%=(rsGetCountrySelect.Fields.Item("Country").Value)%>"><%=(rsGetCountrySelect.Fields.Item("Country").Value)%></option>
			<% 
			rsGetCountrySelect.MoveNext()
			Wend
			rsGetCountrySelect.Requery
			%>
			</select>
			<div class="invalid-feedback">
					Country is required
			</div>
			</div>
			<div class="form-row">
			<div class="col state">
			<label for="state">State (USA)</label>
			<select data-pg-state class="form-control" required name="state" id="state" autocomplete="shipping address-level1">

			<!--#include virtual="/includes/inc_states_select.asp"-->
			</select>
			<div class="invalid-feedback">
					State is required
			</div>
			</div>
			<div class="col province">
			<label for="province">Province / State</label>
			<input class="form-control" name="province" id="province" type="text" value="<% if var_add_only = "" then %><%= session("province") %><% end if %>" autocomplete="shipping address-level2" />
			<div class="invalid-feedback">
					Province is required
			</div>
			</div>
			<div class="col province-canada">
			<label for="province-canada">Province <span class="text-danger">*</span></label>
			<select data-pg-prov class="form-control" name="province-canada" id="province-canada" autocomplete="shipping address-level2">
			<!--#include virtual="/includes/inc_province_canada_select.asp"-->
				  </select>
				  <div class="invalid-feedback">
						Province is required
				</div>
			</div>
			<div class="col">
			<label for="zip">Zip/Postal code <span class="text-danger">*</span></label>
			<input data-pg-pc class="form-control" required name="zip" id="zip" type="text" value="<% if var_add_only = "" then %><%= session("zip") %><% end if %>" autocomplete="shipping postal-code" />
			<div class="invalid-feedback">
					Zip/Postal code is required
			</div>
		</div>
		
	</div>
	</div>
	</section>
	</div><!-- address fields -->
		<input type="hidden" name="type" id="type" value="">
		<input type="hidden" name="cim-id" id="cim-id" value="">
				
		