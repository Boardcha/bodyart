<%
'Variables that need to be set on page(s) so mailer can send
' mailer_type | mail_to_email | mail_to_name | mail_subject | mail_body 

	'===== SETTING DEFAULT VARIABLES ===============
	'===== THIS NEEDS TO STAY AS BAFSERVICE1 TO MATCH OUR SEND EMAIL FROM AMAZON SES OF BAFSERVICE1@
	mail_reply_email = "bafservice1@bodyartforms.com"
	mail_reply_name = "Bodyartforms"
	display_interface = "yes"


	mail_questions_text = "If you have any questions about your account or any other matter, please feel free to contact us <a href=""mailto:bafservice1@bodyartforms.com"" style=""color: #475F8D !important; text-decoration: none;"">via e-mail</a> or by phone at (877) 223-5005."
	
	if (payment_approved = "yes" or mailer_type = "addons approved" or mailer_cash_order = "yes") and done_mailing_certs = "yes" then ' set variable for order details for any approved payment (cc, paypal, cash)
	
		if TotalSpent > 275 AND Session("CouponCode") = "" then
			email_preffered_total = "Your 10% discount: &#8722;" + FormatCurrency(total_preferred_discount, -1, -2, -2, -2) + "<br/>"
		end if ' if preferred customer 
		
		if session("amount_to_collect") > 0 then
			email_sales_tax = "Tax: " + FormatCurrency(session("amount_to_collect"), -1, -2, -2, -2) + "<br/>"
		end if
		
		if Session("CouponCode") <> "" then
			email_coupon_discount = "Coupon: &#8722;" + FormatCurrency(var_couponTotal, -1, -2, -2, -2) + "<br/>"
		end if 

		if Session("GiftCertAmount") <> 0 then 
			email_gift_cert_used = "Gift certificate: &#8722;" + FormatCurrency(Session("GiftCertAmount"), -1, -2, -2, -2) + "<br/>"
		end if 
		
		if session("usecredit") = "yes" then
			email_store_credit_used = "Store credit: &#8722;" + FormatCurrency(Session("storeCredit_used"), -1, -2, -2, -2) + "<br/>"
		end if 
		
		if session("credit_now") <> "" then 
			if session("credit_now") <> 0 then
				email_use_now_credits = "Credit: &#8722;" + FormatCurrency(session("credit_now"), -1, -2, -2, -2) + "<br/>"
			end if 
		end if ' use now credits
		
		' Totals area at bottom of email receipts
		var_email_totals = "Subtotal: " + FormatCurrency(var_subtotal, -1, -2, -2, -2) + "<br/>" + email_coupon_discount + " " + email_preffered_total + " " + email_use_now_credits + " " + email_sales_tax + " " + " " + email_store_credit_used + " " + email_gift_cert_used + "Shipping (" + session("var_email_shipping_option") + "): " + FormatCurrency(session("shipping_cost"), -1, -2, -2, -2) + "<br/><span style='padding-top:5px;font-weight: bold; font-size: 18px'>TOTAL (Paid with " + strCardType + " " + strBilling_cardnumber + ") " + FormatCurrency(var_grandtotal, -1, -2, -2, -2) + " USD</span>"

	end if ' order details build-out

	if IsArray(array_details_2) = True then

		mail_order_details = ""
		For i = 0 to (ubound(array_details_2, 2) - 1)

			' Do not write to email receipt if it's tax... display it in the totals area above
			if Instr(1, array_details_2(2,i), "Tax") = 0 Then

				'	https://bafthumbs-400.bodyartforms.com
				'	https://bodyartforms-products.bodyartforms.com

				'======= REMOVE GAUGE CARD, STICKERS, STORE CREDITS FROM ITEMS LIST
				if array_details_2(6,i) <> 1430 AND array_details_2(6,i) <> 3928 AND array_details_2(6,i) <> 2890 then

					mail_order_details = mail_order_details & "<tr style='border-bottom: 1px solid rgb(100, 100, 100);'><td style='padding-top:15px;padding-right:15px;' valign='top'><img style='width:200px' src='https://bafthumbs-400.bodyartforms.com/" & array_details_2(9,i) & "'></td>"
					
					mail_order_details = mail_order_details & "<td style='padding-top:15px' valign='top'>" & array_details_2(2,i)
					'==== IF A GAUGE IS FOUND ==================
					if array_details_2(11,i) <> "" then 
						mail_order_details = mail_order_details & " <br>" & array_details_2(11,i)
					end if
					'==== IF A GAUGE IS FOUND ==================
					if array_details_2(3,i) <> "" then 
						mail_order_details = mail_order_details & " <br><b>Gauge</b> " & array_details_2(3,i)
					end if
					'==== IF A LENGTH IS FOUND ==================
					if array_details_2(10,i) <> "" then 
						mail_order_details = mail_order_details & "&nbsp;&nbsp;&nbsp;&nbsp;<b>Length</b> " & array_details_2(10,i)
					end if
					'==== IF PREORDER TEXT ==================
					if array_details_2(5,i) <> "" then 
						mail_order_details = mail_order_details & " <br>" & array_details_2(5,i)
					end if

					mail_order_details = mail_order_details & "<br><br>Quantity " & array_details_2(1,i)

					'====== ONLY SEND OUT PRICE DATA ON INITIAL ORDER PLACEMENT ==========
					if mailer_type = "cc approved" Then
						mail_order_details = mail_order_details & " @ " & FormatCurrency(array_details_2(4,i)) & "<br><span style='font-weight:bold;font-size:1.1em'>" & FormatCurrency(array_details_2(4,i) * array_details_2(1,i),2)

							'==== IF ANODIZATION FEE FOUND ==================
							if array_details_2(8,i) > 0 then 
								mail_order_details = mail_order_details & " <br>+ " & FormatCurrency(array_details_2(8,i) * array_details_2(1,i),2) & " color add-on fee"
							end if

					end if

					mail_order_details = mail_order_details & "</span></td></tr>"
				end if '=== filter out gauge card, sticker and credits
			end if
		next
	end if '=======  array_details_2(1) <> ""

	'========== LAST UPDATED FEB 2022 =====================================================
	if done_mailing_certs = "no" then ' gift certificate creation / receipt
		google_utmsource = "Gift certificate received"
		mail_to_email = rec_email
		mail_to_name = rec_name
		mail_subject = your_name + " has gifted you a $" & gift_amount & " Bodyartforms gift certificate!"
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-size:30px;font-weight:bold'>YOUR DAY JUST GOT BETTER.</div><a href='https://bodyartforms.com' style='text-decoration:none'><img src='https://bodyartforms.com/images/baf-present.png' style='width:200px;height:auto;margin-top:10px;margin-bottom:10px'></a></div><div style='font-family:Arial;font-size:20px;border: 6px dashed #696887;padding:10px;font-weight:bold;text-decoration:none;text-align:center'>You received $" & gift_amount & "<br><br>CODE: " & var_cert_code & " </div><br/>" & your_name & "'s message to you:<br/>" & message & "<br/><br/>To use your gift card, simply <a href='https://www.bodyartforms.com' style='color: #696887 !important; text-decoration: none'><strong>shop online at Bodyartforms</strong></a>. At checkout there will be a box where you can enter and apply your code.<div style='font-family:Arial;font-size:30px;font-weight:bold;text-align:center;margin-top:20px'>ENJOY</div><div style='text-align: center'><table style='border-collapse: separate; border-spacing: 4px;width:300px;margin-left: auto; margin-right: auto'><tr><td width='50%' style='background-color:#41415a'><a href='https://bodyartforms.com/products.asp?new=Yes' style='display:block;height:100%;padding:10px;color:#fff; text-decoration: none'>Shop new items</a></td><td width='50%'  style='background-color:#41415a'><a href='https://bodyartforms.com/products.asp?feature=top_seller' style='display:block;height:100%;padding:10px;color:#fff;text-decoration: none'>Shop top sellers</a></td></tr></table></div>"
		
		Call baf_sendmail()
		
	end if ' send out gift certificate

	if mailer_type = "new account" then 'newly created customer account
		google_utmsource = "Account registration"
		mail_to_email = email
		mail_to_name = "New Account"
		mail_subject = "Bodyartforms welcome"
		mail_body = "<div class='column2' style='box-sizing: border-box;display:inline-block;vertical-align: top;margin-bottom:3em'><strong>Hello!</strong><br/><br/>Your account is now active!. You can access your account page at the link below:<br/><a href='https://bodyartforms.com/account.asp' style='color: #475F8D !important; text-decoration: none'>https://bodyartforms.com/account.asp</a><br/><br/><strong>Your username is:</strong> " & mail_to_email & "</div><div class='column2' style='box-sizing: border-box;display:inline-block;color:#0c5460;background-color:#d1ecf1;border-color:1px solid #bee5eb;padding:.75rem 1.25rem;border-radius:.25em'><div style='font-weight:bold;font-size:1.2em'>About our company</div><p>We're a sister owned family business based near Austin, Texas. We've been around since 2001! We have a small close knit group that work hard to make sure you have the best experience possible when you order from us.</p><p>Our motto has always been to give the best customer service possible. If we make a mistake, we'll fix it ASAP. If you have questions or concerns, feel free to get in touch with us via email at service@bodyartforms.com or via phone at our toll free # (877) 223-5005. We're happy to help!</p><div style='font-weight:bold;font-size:1.2em'>What we believe in</div><p>We believe in treating our employees fairly by paying a livable wage, providing health care & retirement benefits, & providing paid vacation/sick time. We also keep a very chill work environment... no bosses breathing down your neck over here!</p>And lastly, we give to charity ... In most years we give over $50,000! We proudly support various causes that are making a difference around the world. The more our company makes, the more profits we donate to charity!</p></div>"
		
		Call baf_sendmail()
	end if ' create new account
	
	if mailer_type = "account activation" then
		google_utmsource = "Account activation"
		mail_to_email = email
		mail_to_name = "Account Activation"
		mail_subject = "Bodyartforms Account Activation Link"
		mail_body = "<div class='column2' style='box-sizing: border-box;display:inline-block;vertical-align: top;margin-bottom:3em'><strong>Hello!</strong><br/><br/>Thank you for registering with Bodyartforms.com. Please click on the below link to activate your account.<br/><br/><a href='https://bodyartforms.com/account-activation.asp?email=" & mail_to_email & "&hash=" & activation_hash & "' style='color: #475F8D !important; text-decoration: none; white-space: nowrap;'>https://bodyartforms.com/account-activation.asp?email=" & mail_to_email & "&hash=" & activation_hash & "</a><br/><br/><br/><br/>If you have any questions or need assistance please reply to this e-mail to get in touch with us or call customer service at (877) 223-5005</div>"
		
		Call baf_sendmail()
	end if ' create new account	

	'============ Send out a one time use coupon (for new customer accounts)=========================
	if email_onetime_coupon = "yes" then 
		google_utmsource = "One time use coupon"
		If google_signin_email <> "" Then
			mail_to_email = google_signin_email
		else
			mail_to_email = Request("email")
		end if

		mail_to_name = "New Account Coupon"
		mail_subject = "10% OFF coupon for registering"
		mail_body = "<div style=""text-align:center""><img src=""https://www.bodyartforms.com/images/10-percent-off.png"" width=""250px""></div><br/><br/>To show our appreciation for you registering an account at Bodyartforms, here is a one time 10% OFF coupon code <span style=""color:#74498C;font-weight:bold"">" & var_cert_code & "</span> that you can use through " & FormatDateTime(now()+29,2) & ". Simply use the code  <span style=""color:#74498C;font-weight:bold"">" & var_cert_code & "</span> at checkout to receive your discount :)<br/><br/>If you have any questions or need assistance please reply to this e-mail to get in touch with us or call customer service at (877) 223-5005"
		
		Call baf_sendmail()
	end if

	if email_newsletter_signup_coupon = "yes" then ' Send out a one time use coupon (for newsletter signups)
		google_utmsource = "Newsletter signup coupon"
		mail_to_email = var_email
		mail_to_name = ""
		mail_subject = "15% OFF Bodyartforms Newsletter Welcome!"
		mail_body = "<div style=""text-align:center""><img src=""https://www.bodyartforms.com/images/15-percent-off.png"" width=""250px""></div><br/><br/><strong>Hello!</strong><br/><br/>To show our appreciation for signing up for our newsletter at Bodyartforms, here is a one time 15% OFF coupon code <span style=""color:#74498C;font-weight:bold"">" & var_cert_code & "</span> that you can use through " & FormatDateTime(now()+29,2) & ". Simply use the code  <span style=""color:#74498C;font-weight:bold"">" & var_cert_code & "</span> at checkout to receive your discount :)<br/><br/>If you have any questions or need assistance please reply to this e-mail to get in touch with us or call customer service at (877) 223-5005"
		
		Call baf_sendmail()
	end if
	
	'===== LAST UPDATED NOVEMBER 2021================
	if mailer_type = "cc approved" then ' credit card approved order receipt
		google_utmsource = "Credit card / PayPal order receipt"
		mail_to_email = session("email")
		mail_to_name = session("shipping_first")
		mail_subject = "Bodyartforms order confirmation"

		if pay_method_afterpay <> "yes" then
			add_ons_link = "<a href='https://bodyartforms.com/products.asp?new=Yes&addon=yes&id=" & Session("invoiceid") & "'><span style='display:inline-block;padding:10px;font-weight:bold;border:#696986 1px solid;cursor:pointer;color:#000000'>Forgot something? Click here to add items before your order ships out</span></a><br/><br/>"
		end if
		
	
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>THANKS FOR YOUR ORDER</div>We'll send you another email with your tracking information when your order ships out.<br><b>Invoice #</b> " & Session("invoiceid") & "<br><b>Order date:</b> " & date() & "<br><br>" & add_ons_link & "</div>" & _
		"<div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>DELIVERY DETAILS</div><br/>" + session("shipping_first") + " " + session("shipping_last") + " " + session("shipping_company") + "<br/>" + session("shipping_address1") + " " + session("shipping_address2") + "<br/>" + session("city") + ", " + session("state") + " " + session("shipping_province") + " " + session("shipping_zip") + "<br/>" + session("country") & _
		"<br/><br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"
		mail_body = mail_body & "<div style='text-align:right;background-color:#e6e6e6;border-bottom: 1px solid rgb(100, 100, 100);border-top: 1px solid rgb(100, 100, 100);margin-top:20px;margin-bottom:20px;padding:20px'>" & var_email_totals & "</div>"
		
		Call baf_sendmail()

	end if ' credit card payment approved

	if mailer_type = "addons approved" then ' credit card approved order receipt
		google_utmsource = "Addons order receipt"
		mail_to_email = session("email")
		mail_to_name = session("shipping_first")
		mail_subject = "Bodyartforms add-ons order confirmation"
		
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>ADD ONS ORDER CONFIRMATION</div><br><b>Items added to invoice #</b> " & Session("invoiceid") & "<br><b>Order date:</b> " & date() & "<br><br>" & add_ons_link & "</div>" & _
		"<div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>DELIVERY DETAILS</div><br/>" + session("shipping_first") + " " + session("shipping_last") + " " + session("shipping_company") + "<br/>" + session("shipping_address1") + " " + session("shipping_address2") + "<br/>" + session("city") + ", " + session("state") + " " + session("shipping_province") + " " + session("shipping_zip") + "<br/>" + session("country") + "</td></tr></table>" & _
		"<br/><br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"
		mail_body = mail_body & "<div style='text-align:right;background-color:#e6e6e6;border-bottom: 1px solid rgb(100, 100, 100);border-top: 1px solid rgb(100, 100, 100);margin-top:20px;margin-bottom:20px;padding:20px'>" & var_email_totals & "</div>"
		
		Call baf_sendmail()
	end if ' add-ons payment approved
	
	'====== UPDATED JAN 2022 ================================
	if mailer_cash_order = "yes" then ' cash order receipt
		google_utmsource = "Cash order receipt"
		mail_to_email = var_email
		mail_to_name = var_shipping_first
		mail_subject = "Bodyartforms order confirmation"
		
		if var_shipping_company <> "" then
			var_shipping_company = "<br/>" + var_shipping_company
		end if
		
		var_invoiceid = Session("invoiceid")
			
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>THANKS FOR YOUR ORDER</div>We'll send you another email with your tracking information when your order ships out.<br><b>Invoice #</b> " & Session("invoiceid") & "<br><b>Order date:</b> " & date() & "<br><br>" & add_ons_link & "</div>" & _
		"<br><br/>Your order confirmation is below for your records.<br/><br/><p>For money order and cash payments, only US funds are accepted.  If payment is sent in any currency other than US funds, your payment will be returned to you.  We cannot accept checks, money grams, or wire transfers.</p><br/><p>We recommend adding insurance or tracking on any payments sent.  We are not responsible for payments lost in shipping.</p><br/><p><strong>For money order payments:</strong><br/>The best place to get a money order is at your bank or post office. Please be sure the money order is in US funds and that the payment is made out to Bodyartforms.  Please also include your invoice # on the money orders FOR line. The address to send payment to is below. After we receive payment your order will ship out the following business day.<br/><br/><p><strong>For cash payments:</strong><br/>Please send US funds wrapped in an extra piece of paper to conceal it in the envelope. The address to send payment to is below. After we receive payment your order will ship out the following business day. <i>Please do not send coins in your envelope. If you do, it may tear the envelope and cause the funds to be lost.</i></p><br/><p><strong>Cancellations and order changes:</strong><br/><p>If you need to cancel your order, simply do not send payment in and it will cancel itself out after 60 days. If you need to make a change to your order, you can place a new order and send payment in for the new order only.</p><p><i>When sending in payment, please be sure to put your full name and mailing address, as well as the invoice #, on the envelope.</i></p><br/><strong>Send payment to:</strong><br/>Bodyartforms<br/>1966 S. Austin Ave.<br/>Georgetown, TX  78626</p><br/>" & _
		"<div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>DELIVERY DETAILS</div><br/>" + session("shipping_first") + " " + session("shipping_last") + " " + session("shipping_company") + "<br/>" + session("shipping_address1") + " " + session("shipping_address2") + "<br/>" + session("city") + ", " + session("state") + " " + session("shipping_province") + " " + session("shipping_zip") + "<br/>" + session("country")  & _
		"<br/><br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"
		mail_body = mail_body & "<div style='text-align:right;background-color:#e6e6e6;border-bottom: 1px solid rgb(100, 100, 100);border-top: 1px solid rgb(100, 100, 100);margin-top:20px;margin-bottom:20px;padding:20px'>" & var_email_totals & "</div>"
		
		Call baf_sendmail()

	end if ' cash order
	
	if mailer_type = "admin_reset_user_password" then ' ADMIN
		mail_to_email = request.form("email")
		mail_to_name = request.form("emailname")
		mail_subject = "Reset password for " + request.form("username")
		
		mail_body = "<a href=""http://www.bodyartforms.com/admin/password_reset.asp?token=" + reset_token + """>Click here</a> to reset " + request.form("username") + "'s password."
		
		Call baf_sendmail()
	end if ' ADMIN reset user password
	
	if mailer_type = "front_reset_user_password" then ' FRONT END CUSTOMER
		google_utmsource = "Reset user password"
		mail_to_email = request.form("email")
		mail_to_name = ""
		mail_subject = "Reset your Bodyartforms password"
		
		mail_body = "Hello,<br/><br/>We've sent this message because you have or someone has requested that your Bodyartforms password be reset. To get back into your Bodyartforms account you'll need to create a new password.<br/><br/>To reset your password, <a href=""https://www.bodyartforms.com/password_reset.asp?token=" + reset_token + """><strong>click this link</strong></a> and you'll be taken to our website where you can securely update your password.<br/><br/>Here's the link again in case you can't click the link above: <br/>https://www.bodyartforms.com/password_reset.asp?token=" + reset_token + "<br/><br/><br/><br/><br/>"
		
		Call baf_sendmail()
	end if ' FRONT END CUSTOMER reset user password
	
	if mailer_type = "reject-photo" then ' Photo rejection
		google_utmsource = "Photo rejection"
		mail_to_email = request.form("Email")
		mail_to_name = request.form("Name")
		mail_subject = "Your photo submission at Bodyartforms"
		mail_body = "Unfortunately your photo of " & Request.form("title") & " had to be rejected because " & Request.form("photo_status") & ". You are more than welcome to revise and re-submit your photo :) Thanks!"
		
		Call baf_sendmail()
	end if ' Photo rejection

	if mailer_type = "reject-review" then ' Review rejection
		google_utmsource = "Review rejection"
		mail_to_email = request.form("email")
		mail_to_name = request.form("name")
		mail_subject = "Your jewelry review submission at Bodyartforms"
		mail_body = Request.Form("vote")
		
		Call baf_sendmail()
	end if ' Photo rejection

	'===== LAST UPDATED DECEMBER 2021================
	if mailer_type = "order-shipment-notification" then ' from admin section
		google_utmsource = "Order shipment notification"
		mail_to_email = rsGetInvoice("email")
		mail_to_name =  rsGetInvoice("customer_first")
		mail_subject = "Bodyartforms order shipment notification"
		
		' For orders that are not office pick up
		if var_shipping_type <> "OFFICE PICK UP" then
			
			mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>YOUR ORDER HAS SHIPPED!</div><br>"
			
			mail_body = mail_body & "<div style='text-align:left'>Hello " & mail_to_name & ",<br/><br/>This is an automated email to notify you that your order is being packaged up today and will be in the mail soon. We appreciate your business very much!<br><br>"

			mail_body = mail_body & var_tracking

			mail_body = mail_body & "</div><br><div style='text-align:center'><b>Invoice #</b> " & rsGetInvoice("ID") & "<br><b>Order date:</b> " & FormatDateTime(rsGetInvoice("date_order_placed"),vbLongDate) & _
			"</div><div style='text-align:left'><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>DELIVERY DETAILS</div><br/>"

			mail_body = mail_body & rsGetInvoice("customer_first") & " "
			mail_body = mail_body & rsGetInvoice("customer_last") & " "
			mail_body = mail_body & rsGetInvoice("company") & "<br/>"
			mail_body = mail_body & rsGetInvoice("address") & " "
			mail_body = mail_body & rsGetInvoice("address2") & "<br/>"
			mail_body = mail_body & rsGetInvoice("city") & ", "
			mail_body = mail_body & rsGetInvoice("state") & " "
			mail_body = mail_body & rsGetInvoice("province") & " "
			mail_body = mail_body & rsGetInvoice("zip") & "<br/>"
			mail_body = mail_body & rsGetInvoice("country")

			mail_body = mail_body & "<br/><br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"

			mail_body = mail_body & "</div></div>"

		else ' if it's an office pick up

			mail_body = "Hello " & mail_to_name & ",<br/><br/>This is an automated email to notify you that your order #" & rsGetInvoice("ID") & " is being packaged and will be available for pickup Mon - Fri (9am to 5pm).<br/><br/><strong>Our address is:</strong> <br/>Bodyartforms<br/>1966 S. Austin Ave.<br/>Georgetown, TX  78626<br/><br/><a href='https://www.google.com/maps/place/1966+S+Austin+Ave,+Georgetown,+TX+78626/@30.6257777,-97.6806638,17z/data=!3m1!4b1!4m5!3m4!1s0x8644d66050a433bf:0xa7e710a073726aa2!8m2!3d30.6257777!4d-97.6784751?hl=en'>Link to Google Maps</a><br/><a href='https://www.youtube.com/watch?v=42U6-0VHz5c&feature=youtu.be'>Here's a video of how to find our warehouse</a><br/><br/>We appreciate your business very much!<br/><br/>If you have any questions or need assistance with your order please reply to this e-mail to get in touch with us or call customer service at (877) 223-5005"
		end if
		
		Call baf_sendmail()
	end if ' order-shipment-notification
	
	'=============== LAST UPDATED JAN 2022 ===============================
	if mailer_type = "OUT_FOR_DELIVERY" then 
		google_utmsource = "Out for delivery notification"
		mail_to_email = var_email
		mail_to_name = var_first
		mail_subject = "Your Bodyartforms order is out for delivery today"	
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>YOUR ORDER IS BEING DELIVERED TODAY</div><br>"
				
		mail_body = mail_body & "<div style='text-align:left'>Hello " & mail_to_name & ",<br/><br/>This is an automated email to let you know that your Bodyartforms order is scheduled to be delivered today!<br><br>"

		mail_body = mail_body & "<div style='font-family:Arial;color: #ffffff;;background-color:#696986;padding:20px;border-radius:10px'>Your tracking # is " & rsGetInvoice("USPS_tracking") & "<br>Shipped via " & rsGetInvoice("shipping_type") & "<br><br><a style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#41415a;padding:10px;font-weight:bold;text-decoration:none' href='"

		mail_body = mail_body & var_tracking & "'>TRACK YOUR PACKAGE</a></div>"

		mail_body = mail_body & "</div><br><div style='text-align:center'><b>Invoice #</b> " & rsGetInvoice("ID") & "<br><b>Order date:</b> " & FormatDateTime(rsGetInvoice("date_order_placed"),vbLongDate) & _
			"</div><div style='text-align:left'>"

		mail_body = mail_body & "<br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"

		mail_body = mail_body & "</div></div>"
		
		Call baf_sendmail()
	end if 

	'=============== LAST UPDATED JAN 2022 ===============================
	if mailer_type = "ORDER_DELIVERED" then 

		google_utmsource = "Order delivered notification"
		mail_to_email = var_email
		mail_to_name = var_first
		mail_subject = "Your Bodyartforms order has been delivered"	
		mail_body = "<div style='text-align:center'><div style='font-family:Arial;font-weight:bold;font-size:26px'>YOUR ORDER HAS BEEN DELIVERED</div><br>"
				
		mail_body = mail_body & "<div style='text-align:left'>Hello " & mail_to_name & ",<br/><br/>This is an automated email to let you know that your Bodyartforms order has been delivered! We hope you love everything you received &#10084;<br><br>If there are any issues with your order please reply to this email to get in touch with us. We'll reply as soon as we can!<br><br>"

		mail_body = mail_body & "<div style='font-family:Arial;color: #ffffff;;background-color:#696986;padding:20px;border-radius:10px'>Your tracking # is " & rsGetInvoice("USPS_tracking") & "<br>Shipped via " & rsGetInvoice("shipping_type") & "<br><br><a style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#41415a;padding:10px;font-weight:bold;text-decoration:none' href='"

		mail_body = mail_body & var_tracking & "'>TRACK YOUR PACKAGE</a></div>"

		mail_body = mail_body & "</div><br><div style='text-align:center'><b>Invoice #</b> " & rsGetInvoice("ID") & "<br><b>Order date:</b> " & FormatDateTime(rsGetInvoice("date_order_placed"),vbLongDate) & _
			"</div><div style='text-align:left'>"

		mail_body = mail_body & "<br/><div style='font-family:Arial;font-size:16px;color: #ffffff;;background-color:#696986;padding:10px'>ITEMS</div><table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>"

		mail_body = mail_body & "</div></div>"
		
		Call baf_sendmail()
	end if 
	
	'=============== LAST UPDATED JAN 2022 ===============================
	if mailer_type = "ORDER_DELAYED" then 
		google_utmsource = "Order delayed notification"
		mail_to_email = var_email
		mail_to_name = var_first
		'cc1_name = "Parth"
		'cc1_email = "tanejap652@gmail.com"
		mail_subject = "Notification - Your Bodyartforms shipment is delayed"	
		mail_body = "Hello " & mail_to_name & ",<br/><br/>This is an automated email to notify you that your order #" & var_invoiceid & " is delayed from the original delivery estimate date, " & var_estimated_delivery_date & ".<br/><br/>" & var_tracking & "<br/><br/>We're very sorry for any inconvenience!<br/>If you have any questions or need assistance with your order please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"
		
		Call baf_sendmail()
	end if 
	
	if mailer_type = "cancelled_order" then ' front facing order cancellation
		google_utmsource = "Cancelled order receipt"
		mail_to_email = rsGetOrder.Fields.Item("email").Value
		mail_to_name = rsGetOrder.Fields.Item("customer_first").Value
		mail_subject = "Bodyartforms order cancellation confirmation"
		mail_body = "Hello " & rsGetOrder.Fields.Item("customer_first").Value & ",<br/><br/>This is an automated email confirmation that your order #" & invoiceid & " has been cancelled and will not ship out. <br/><br/>" & FormatCurrency(var_order_total, -1, -2, -0, -2) & " has been applied to your store credit.<br/><br/>If you have any questions or need further assistance please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"

		Call baf_sendmail()
		
	end if ' front facing order cancellation

	' ===== From ajax-reship-items to send out notification
	if mailer_type = "reship_approve" then 
		google_utmsource = "Item reship notification"
		mail_to_email = rsGetInvoice.Fields.Item("email").Value
		mail_to_name = rsGetInvoice.Fields.Item("customer_first").Value
		mail_subject = "Item reship notification"

		mail_body = "Hello " & rsGetInvoice.Fields.Item("customer_first").Value & ",<br/><br/>We are very sorry that there were issues with your order #" & rsGetInvoice.Fields.Item("ID").Value
		
		if email_stocked_items <> "" then
			mail_body = mail_body & "<br/><br/>We have set up a new order (#" & move_to_invoice & ") to reship the following items that we had in stock:<ul>" & email_stocked_items & "</ul> This order will ship out in the next 1-2 business days and you will receive a separate e-mail with the tracking #."
		end if

		if email_outofstock_items <> "" then
			mail_body = mail_body & "<br/><br/>Unfortunately we did not have the items in stock to ship out below:<br/><ul>" & email_outofstock_items & "</ul>"

			if rsGetInvoice.Fields.Item("customer_ID").Value > 0 then
				mail_body = mail_body & "We have issued a $" & var_refund_total & " store credit to your account. You can use this credit at anytime at checkout."
			else
				mail_body = mail_body & "We have issued you a $" & var_refund_total & " gift certificate. You can use the following code <strong>" & var_cert_code & "</strong> at checkout on our site to apply your credit to any future order."
			end if

			mail_body = mail_body & " If you would rather have a refund please <a href='https://bodyartforms.com/refunds.asp?id=" & encrypted_code & "'>visit this page</a>."
		end if
		
		mail_body = mail_body & "<br/><br/>We appreciate your business very much!<br/>If you have any questions or need assistance with your order please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"
	
		Call baf_sendmail()
	end if ' ====== Item reship notification


	' ===== From refunds.asp  Receipt of customers refund request/success
	if mailer_type = "customer_submitted_refund_notification" then 
		google_utmsource = "Bodyartforms refund receipt"
		mail_to_email = rsCheckRefund.Fields.Item("email").Value
		mail_to_name = rsCheckRefund.Fields.Item("customer_first").Value
		mail_subject = "Bodyartforms refund receipt"

		mail_body = "Your refund of $" & var_db_refund_amt & " has been successfully processed and will typically post to your account within 5-7 business days."
				
		mail_body = mail_body & "<br/><br/>We appreciate your business very much!<br/>If you have any questions or need assistance please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"
	
		Call baf_sendmail()
	end if ' ====== customer_submitted_refund_notification
	
	' ===== From refunds.asp  Receipt of customers refund request/success
	if mailer_type = "customer_submitted_refund_as_store_credit_notification" then 
		google_utmsource = "Bodyartforms refund receipt"
		mail_to_email = rsCheckRefund.Fields.Item("email").Value
		mail_to_name = rsCheckRefund.Fields.Item("customer_first").Value
		mail_subject = "Bodyartforms refund receipt"

		mail_body = "Your refund of $" & var_db_refund_amt & " has been successfully processed into your account as a store credit."
				
		mail_body = mail_body & "<br/><br/>We appreciate your business very much!<br/>If you have any questions or need assistance please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"
	
		Call baf_sendmail()
	end if ' ====== customer_submitted_refund_as_store_credit_notification	
	
	'======= RE-WORDED BACKORDER EMAIL DECEMBER 2021 ==================================
	'======= GENERAL BACKORDER EMAIL ==================================
	if mailer_type = "backorder" then 'Backorder email
		google_utmsource = "Backorder notification"
		mail_to_email = var_customer_email
		mail_to_name = var_customer_name
		mail_subject = "Backorder notice - " + var_item_description + " (Invoice #" & var_invoice_number & ")"
		mail_body = "Hey " & var_customer_name &_
		"<br><br>Thank you so much for your order! We are sorry to say that we do not have the item listed below to send out because " & var_bo_reason & ". If you ordered more than one item, then your package has already shipped out with the rest of your items." &_ 
		"<table style='border-collapse:collapse;width: 98%'>" + mail_order_details + "</table>" &_
		"<br>We always try our best to fulfill every order perfectly, but we dropped the ball. That's on us. Here's what we can do:" &_
		"<ul>"
		
		'======== Only display this option if it's a regularly stocked item
		if var_jewelry_status = "None" then
			mail_body = mail_body & "<li>You can leave the item on backorder and we'll ship it when it comes back in stock</li>"
		end if

		If var_refund_total > 0 Then ' Make sure if it is not just a free item
			If var_customer_number > 0 Then
				mail_body = mail_body & "<li><a href='https://bodyartforms.com/refunds-backordered-items.asp?id=" & var_invoice_number & "&hash=" & encrypted_code & "'>You can get in-store credit for the item</a></li>" 
			End If
			mail_body = mail_body & "<li><a href='https://bodyartforms.com/refunds-backordered-items.asp?id=" & var_invoice_number & "&hash=" & encrypted_code & "'>You can get a refund for the item</a></li>"
		End If
		
		mail_body = mail_body & "<li>You can exchange the item for something else (Just reply and let us know which item you want instead)</li>" &_
		"</ul>" &_
		"<div style='font-family:Arial;color:#ffffff;background-color:#696986;padding:20px;border-radius:10px'>We'd also like to extend you this one time coupon code for <strong>15% off any future order</strong> by way of apology.<div style='text-align:center;font-family:Arial;font-size:16px;color: #ffffff;;background-color:#41415a;padding:10px;font-weight:bold;text-decoration:none;margin:15px'>" & var_cert_code & "</div>We take customer service super seriously, and we're always working on improving. If you have any questions or feedback, we'd love to hear from you at <a style='text-decoration:none' href='mailto:help@bodyartforms.com'>help@bodyartforms.com</a></div>" &_
		"<br><br>Thank's again for your support, and we look forward to hearing from you," &_
		"<br><br>The Bodyartforms Team"
		
		Call baf_sendmail()
		
	end if '======== BACKORDERED ITEM

	' ===== BACKORDERS From ajax-backorder-process to send out notification
	if mailer_type = "bo_notification" then 
		google_utmsource = "Backorder status update"
		mail_to_email = rsGetOrder.Fields.Item("email").Value
		mail_to_name = rsGetOrder.Fields.Item("customer_first").Value
		mail_subject = "Backorder status update"

		mail_body = "Hello " & rsGetOrder.Fields.Item("customer_first").Value & ",<br/><br/>We are very sorry about the backordered " & rsGetItemDetails.Fields.Item("item_title").Value & " on your order (Invoice #" & rsGetOrder.Fields.Item("ID").Value & "). " & backorder_email_body
			
		mail_body = mail_body & "<br/><br/>We appreciate your business very much!<br/>If you have any questions or need assistance with your order please reply to this e-mail to get in touch with us. We're here to help Mon - Fri from 9am - 5pm.<br/><br/>Customer service:  (877) 223-5005"
	
		Call baf_sendmail()
	end if ' ====== BACKORDERS From ajax-backorder-process to send out notification

	if mailer_type = "bo-preorder-standard" then 'Backorder custom order email - BASIC
		google_utmsource = "Preorder backorder notice - generic"
		mail_to_email = rsGetItem.Fields.Item("email").Value
		mail_to_name = rsGetItem.Fields.Item("customer_first").Value
		mail_subject = "Custom item backorder notice (#" & rsGetItem.Fields.Item("InvoiceID").Value & ")"
		mail_body = "Unfortunately, the manufacturer has placed your " + rsGetItem.Fields.Item("title").Value + " on back order.  The piece you ordered can still be made, however it may take longer than the estimated 4-6 weeks for us to receive it and then ship it out to you. We are very sorry for the inconvenience, and we would like to give you a few options:<p><ul><li>We can leave the item on back order and when it arrives from the manufacture we will ship out your entire order</li><li>We can ship out the stock items that we have now via the shipping method you chose, and ship out the custom item via basic mail as soon as it comes in</li><li>We can issue you a refund or store credit for the custom item and ship out the rest of your order</li><li>We can cancel the order and issue you a refund or store credit for the entire order</li></ul><p>Please let us know what you would prefer, by either replying to this email, or giving us a call at 512-943-8654 and we will get you taken care of.<p>Thank you!"
		
		Call baf_sendmail()
		
	end if ' Backorder custom item email - BASIC

	if mailer_type = "bo-preorder-specs" then 'Backorder custom item email - SPECS ISSUE
		google_utmsource = "Preorder need more specs"
		mail_to_email = rsGetItem.Fields.Item("email").Value
		mail_to_name = rsGetItem.Fields.Item("customer_first").Value
		mail_subject = "Custom item specs problem (Invoice #" & rsGetItem.Fields.Item("InvoiceID").Value & ")"
		mail_body = "Unfortunately, the manufacturer has contacted us to let us know the your item " + rsGetItem.Fields.Item("title").Value + " cannot be made with the specs that were included. We are very sorry for the inconvenience, and we would like to give you a few options:<p><ul><li>You can contact Melissa at Melissa@bodyartforms.com and she can let you know what other options are available for the item(s) that you ordered.</li><li>We can issue you a refund or store credit for the custom item and ship out the rest of your order</li><li>We can cancel the order and issue you a refund or store credit for the entire order. If there are other custom items on the order, there will be a 15% restocking fee for canceling the other custom items.</li></ul><p>Please let us know what you would prefer, by either replying to this email, or giving us a call at 512-943-8654 and we will get you taken care of.<p>Thank you!"
		
		Call baf_sendmail()
		
		
	end if ' Backorder custom item email	- SPECS ISSUE
	
	if mailer_type = "bo-preorder-discontinued" then 'Backorder custom item email - DISCONTINUED
		google_utmsource = "Preorder discontinued notice"	
		mail_to_email = rsGetItem.Fields.Item("email").Value
		mail_to_name = rsGetItem.Fields.Item("customer_first").Value
		mail_subject = "Custom item discontinued notice (#" & rsGetItem.Fields.Item("InvoiceID").Value & ")"
		mail_body = "Unfortunately, the manufacturer has discontinued " + rsGetItem.Fields.Item("title").Value + "and your custom item can no longer be made. We are very sorry for the inconvenience, and we would like to give you a couple of options:<p><ul><li>We can issue you a refund or store credit for the custom item and ship out the rest of your order</li><li>We can cancel the order and issue you a refund or store credit for the entire order</li></ul><p>Please let us know what you would prefer, by either replying to this email, or giving us a call at 512-943-8654 and we will get you taken care of.<p>Thank you!"
		
		Call baf_sendmail()
		
		
	end if ' Backorder custom item email	- DISCONTINUED

	if mailer_type = "money-request" then 'PayPal money request
		google_utmsource = "PayPal money request"
		' 	Found a sweet page that will generate PayPay payment links here
		'	http://www.itaynoy.com/sites/paypal_button_generator/
		mail_to_email = request.form("email")
		mail_to_name = request.form("first")
		mail_subject = "Bodyartforms PayPal Money Request Link"
		mail_body = "Hello " & request.form("first") & ",<br/><br/>Below you'll find your secure link to pay via PayPal. You'll need to be logged into your PayPal account before starting the payment process.<br/><br/><strong>Description:</strong><br/>" & request.form("description") & "<br/><br/><br/><a href='https://www.paypal.com/cgi-bin/webscr?&cmd=_xclick&business=bafpaypal@bodyartforms.com&currency_code=USD&amount=" & request.form("amount") & "&item_name=Body jewelry for Invoice " & request.form("invoice") & "' style=""background-color: #FFBF00; padding: .5em; text-decoration: none; color: black; font-weight: bold"">Click here to pay via PayPal</a>"
		
		Call baf_sendmail()
		
	end if ' PayPal money request

	' --------  RETURNS -----------------
	if mailer_type = "returns" then
		google_utmsource = "Return notification"
		mail_to_email = rsGetOrder.Fields.Item("email").Value
		mail_to_name = rsGetOrder.Fields.Item("customer_first").Value
		mail_subject = "Bodyartforms returned order #" & rsGetOrder.Fields.Item("ID").Value

		if var_reason = "Damaged" then
			mail_body = "Unfortunately your order was returned to us in a damaged condition. Please reply to this email and let us know whether you'd like it re-sent or refunded."
		elseif var_reason = "Undeliverable address" OR var_reason = "No reason given" then
			mail_body = "Unfortunately your package has been returned to us :( The reason the shipping provider gave was <strong>" & var_reason & "</strong>. <br><br>The address you gave us is below:<br><br>" & var_company & rsGetOrder.Fields.Item("customer_first").Value & " " &    rsGetOrder.Fields.Item("customer_last").Value & "<br>" & rsGetOrder.Fields.Item("address").Value & "<br>" & var_address2 & rsGetOrder.Fields.Item("city").Value & ", " & rsGetOrder.Fields.Item("state").Value &"" & rsGetOrder.Fields.Item("province").Value & " " & rsGetOrder.Fields.Item("zip").Value & "<br>" & rsGetOrder.Fields.Item("country").Value & "<br><br>If the address above is correct and not delivered, please let us know and we'll re-ship it free of charge to the address above. If the address is not correct, we will need your corrected address and PERMISSION TO RE-BILL SHIPPING to re-send the package. If  you'd like you are also more than welcome to cancel the order and get a refund (less shipping).<br><br>Please reply to this email and let us know how you would like to proceed."
		elseif var_reason = "return_refunded" then
			mail_subject = "Bodyartforms returned item(s) for order #" & rsGetOrder.Fields.Item("ID").Value
			mail_body = "We have received your returned item(s) and have issued the funds back as shown below:<br/><br/><strong>Credit Card (or PayPal)</strong> " & FormatCurrency(cc_refund_due,2) & "<br/><strong>Store credit</strong> " & FormatCurrency(storecredit_refund_due,2) & "<br/><strong>Gift certificate</strong> " & FormatCurrency(giftcert_refund_due,2) & "<br/><br/>Refunds to a credit card usually take 5-7 business days to clear.<br/>PayPal refunds usually take 2-3 days.<br/>All other refunds and store credits should be seen immediately.<br/><br/>If you have any questions or concerns please reply to this email and we'll get back to you as soon as possible (Mon - Fri)."
		else
			mail_body = "Unfortunately your package has been returned to us :( The reason the shipping provider gave was <strong>" & var_reason & "</strong>. <br><br>The address you gave us is below:<br><br>" & var_company & rsGetOrder.Fields.Item("customer_first").Value & " " &    rsGetOrder.Fields.Item("customer_last").Value & "<br>" & rsGetOrder.Fields.Item("address").Value & "<br>" & var_address2 & rsGetOrder.Fields.Item("city").Value & ", " & rsGetOrder.Fields.Item("state").Value &"" & rsGetOrder.Fields.Item("province").Value & " " & rsGetOrder.Fields.Item("zip").Value & "<br>" & rsGetOrder.Fields.Item("country").Value & "<br><br>We are more than happy to re-send your order if you cover the shipping cost to have it re-shipped. If not, we can also refund your order (less shipping).<br><br>Please reply to this email and let us know how you would like to proceed."
		end if
	

		Call baf_sendmail()
		
	end if ' --------  RETURNS
	
	'======= BEGIN CONTACT US PAGE =====================================================
	if mailer_type = "contact-us" then 'Contact us page
		display_interface = "no"

		if Request.form("comments") <> "" then
			
			mail_to_email = "bafservice1@bodyartforms.com"
			mail_to_name = "Bodyartforms"
			mail_reply_email = Request.form("email")
			mail_reply_name = Request.form("name")
			
			mail_subject = Request.form("reason")
			mail_body = "<b>Invoice:</b><br/><a href='http://www.bodyartforms.com/admin/invoice.asp?ID=" + Request.form("invoice") + "'>" + Request.form("invoice") + "</a><br/><br /><b>Email:</b><br/><a href='http://www.bodyartforms.com/admin/order%20history.asp?var_email=" + Request.form("email") + "'>" + Request.form("email") + "</a><br/><br/><b>Comments:</b><br/>" + Request.form("comments")

		end if

		Call baf_sendmail()
	end if 
	'======= END CONTACT US PAGE =====================================================

	'======= BEGIN 500 ERROR PAGE =====================================================
	if mailer_type = "500-error" then ' 500 error page notification
		display_interface = "no"

		if request.form("comments") <> "" then
			
			mail_to_email = "amanda3@bodyartforms.com"
			mail_to_name = "Amanda"
			
			mail_subject = "500 Page Error"
			mail_body = "Customer comments<br> " & request.form("comments")

		end if
		Call baf_sendmail()

	end if ' 500 error page notification
	'======= END 500 ERROR PAGE =====================================================

	'======= BEGIN AUTO FLAG EMAIL =====================================================
	if mailer_type = "auto-flag" then 'Auto flag trip from checkout-final.asp
		display_interface = "no"
		
		mail_to_email = "bafservice1@bodyartforms.com"
		mail_to_name = "Bodyartforms"	
		
		mail_subject = "WEBSITE FRAUD AUTO-FLAG TRIPPED"
		mail_body = "<b>Invoice:</b> " & invoice_id

		Call baf_sendmail()
	end if
	'======= END AUTO FLAG EMAIL ======================================================


	'======= BEGIN ORDER SURVEY EMAIL ======================================================
	if mailer_type = "order-survey" then 'Order survey
		display_interface = "no"
		
		mail_to_email = "baf_surveys@bodyartforms.com"
		mail_to_name = "Bodyartforms"
		mail_reply_email = "amanda3@bodyartforms.com"
		mail_reply_name = "Bodyartforms"	
		
		mail_subject = "BAF survey"
		mail_body = "Invoice #<a href='http://www.bodyartforms.com/admin/invoice.asp?ID=" + InvoiceID + "'>" + InvoiceID + "</a><br/>" + rsInvoice.Fields.Item("customer_first").Value + " " + rsInvoice.Fields.Item("customer_last").Value + "<br/>" + rsInvoice.Fields.Item("email").Value + "<br/><b>Packaged by: </b>" & rsInvoice.Fields.Item("PackagedBy").Value & "<br/><b>Shipped: </b>" + Cstr(rsInvoice.Fields.Item("date_sent").Value) + "<br/><br/>" + CS + Selection + Pricing + StockLevels + Experience + Packaging + Presentation + Items + Quality + Delivery + Overall + NewJewelry + Comments + ItemInfo

		Call baf_sendmail()

	end if
	'======= END ORDER SURVEY EMAIL ======================================================

	'======= BEGIN NOTIFY CS OF BAD REVIEW ======================================================
	if mailer_type = "notify-cs" then
		display_interface = "no"
		
		mail_to_email = "bafservice1@bodyartforms.com"
		mail_to_name = "Bodyartforms"
		
		mail_subject = "Website alert: Unhappy customer / Bad jewelry review"
		mail_body = "<strong>E-mail:</strong><br/>" & request.form("email") & "<br/><br/>Invoice # " & request.form("invoiceid") & "<br/><br/><strong>Their review:</strong><br/>" & request.form("review") & "<br/><br/><a href='https://www.bodyartforms.com/productdetails.asp?ProductID=" & request.form("productid") & "'>https://www.bodyartforms.com/productdetails.asp?ProductID=" & request.form("productid") & "</a><br/><br/>Customer is unhappy with " & request.form("title") & " - " & request.form("details")

		Call baf_sendmail()

	end if
	'======= END NOTIFY CS OF BAD REVIEW ======================================================

	'======= BEGIN NOTIFY CS OF BAD PHOTO SUBMISSION ==========================================
	if mailer_type = "notify-photography" then ' notify photography of bad review
		display_interface = "no"
			
		mail_to_email = "rebekah@bodyartforms.com"
		mail_to_name = "Rebekah"	
		
		mail_subject = "Website alert: Item needs photo fix"
		mail_body = "<strong>Their review:</strong><br/>" & request.form("review") & "<br/><br/><a href='https://www.bodyartforms.com/productdetails.asp?ProductID=" & request.form("productid") & "'>https://www.bodyartforms.com/productdetails.asp?ProductID=" & request.form("productid") & "</a><br/><br/>Customer submitted a review for " & request.form("title") & " - " & request.form("details")

		Call baf_sendmail()
		
	end if
	'======= END NOTIFY CS OF BAD PHOTO SUBMISSION ==========================================


	'======= BEGIN SPLIT ORDER DAILY REPORT =============================================
	if mailer_type = "split-orders" then 'Split orders
		
		mail_to_email = rsGetPacker("email")
		mail_to_name = rsGetPacker("name")
		cc1_name = "Andres"
		cc1_email = "andres@bodyartforms.com"
		mail_reply_email = "andres@bodyartforms.com"
		mail_reply_name = "Andres"	
		
		mail_subject = rsGetPacker("name") & " - Daily orders & errors"
		mail_body = "<strong>Total orders: " & rsEmailStats.Fields.Item("Total").Value & "<br/><br/>" & _ 
		"Autoclaves: " & rsEmailStats.Fields.Item("Autoclaves").Value & "</br>" & _
		"UPS: " & rsEmailStats.Fields.Item("UPS").Value & "<br/>" & _
		"Express & Priority: " & rsEmailStats.Fields.Item("priority").Value & "<br/>" & _
		"DHL & First Class: " & rsEmailStats.Fields.Item("DHL").Value & "</strong>"

		mail_body = mail_body & "<br/><br/><div style=""font-size:x-large;font-weight:bold"">REPORTED ERRORS FOR THE LAST 30 DAYS<br>" & _
			var_error_percentage & " accuracy</div>" & _ 
			"<div>Flip-flop: " & Error_flip_total & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mis-matched: " & Error_matching_total & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Broken: " &  Error_broken_total & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Missing: " & Error_missing_total & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Wrong: " & Error_wrong_total & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Misc: " &  Error_misc_total & "</div><br/><br/>"

		While NOT rsGetErrors_Details.EOF 
			if rsGetErrors_Details.Fields.Item("date_sent").Value <> "" then
				date_sent = FormatDateTime(rsGetErrors_Details.Fields.Item("date_sent").Value,1)
			end if
			If rsGetErrors_Details("item_problem") = "Missing" then
				missing_data = rsGetErrors_Details("ErrorQtyMissing") & "&nbsp;&nbsp;Scanned:&nbsp;" & rsGetErrors_Details("TimesScanned")
			end if

			mail_body = mail_body & "<div style='background-color:#e6e6e6;width:100%;padding:10px;margin-top:5px;vertical-align:top'><strong>" & rsGetErrors_Details("item_problem") & missing_data & "<br/>" & _
				"Shipped: " & date_sent & "</strong><br><br/>" & _
				rsGetErrors_Details("ErrorDescription") & _
				rsGetErrors_Details("qty") & " | " & rsGetErrors_Details("title") & "&nbsp; " & _
				rsGetErrors_Details("Gauge") & "&nbsp; " & rsGetErrors_Details("Length") & "&nbsp; " & rsGetErrors_Details("ProductDetail1") & "&nbsp; " & rsGetErrors_Details("notes") & "</div><br/>"
			

			rsGetErrors_Details.MoveNext()
		Wend

	Call baf_sendmail()

	end if   'mailer_type = "split-orders"
	'======= END SPLIT ORDER DAILY REPORT ==================================================


	'======= BEGIN WEBSITE FEEDBACK ==========================================
	if mailer_type = "website-feedback" then
			
		mail_to_email = "amanda3@bodyartforms.com"
		mail_to_name = "Amanda"
		mail_reply_email = "amanda3@bodyartforms.com"

		mail_subject = "Site Feedback"
		mail_body = "Email " & request.form("feedback-email") & "<br/>Comments: " & request.form("feedback-comments")

		Call baf_sendmail()
		
	end if
	'======= END WEBSITE FEEDBACK  ==========================================

	'======= BEGIN NOTIFY ON WAITING LIST THAT ITEM IS BACK IN STOCK ===========================
	if mailer_type = "notify waiting list" then 'notify customers that items are back in stock from the admin
		google_utmsource = "Waiting list notification"
		mail_to_email = rsGetCustomers.Fields.Item("email").Value
		mail_to_name = "Waiting list notification"
		mail_subject = "Back in Stock: " & rsGetCustomers.Fields.Item("item_name").Value
		mail_body = "You asked us to tell you when the " & rsGetCustomers.Fields.Item("item_name").Value & " would be back in stock for ordering.<br><br>We are pleased to tell you it is now available. We only have a limited amount, and this email is not a guarantee you'll get one. So hurry to <a href=http://www.bodyartforms.com/productdetails.asp?ProductID=" & rsGetCustomers.Fields.Item("ProductID").Value & "&referrer=waiting-list>BAF</a> and make sure you get yours now!<br><br><div style=""background-color: #404064; padding: .4em 0; width:120px; font-weight: bold; text-align: center;""><a href=http://www.bodyartforms.com/productdetails.asp?ProductID=" & rsGetCustomers.Fields.Item("ProductID").Value & "&referrer=waiting-list style=""text-decoration: none; color: #B8B8E6 !important;"">BUY NOW</a></div><a href=http://www.bodyartforms.com/productdetails.asp?ProductID=" & rsGetCustomers.Fields.Item("ProductID").Value & "&referrer=waiting-list><img src=http://bafthumbs-400.bodyartforms.com/" & rsGetCustomers.Fields.Item("picture_400").Value & " width=120px height=120px></a><div style=""padding: 1em 0"">Happy Shopping,<br><br>Your friends at Bodyartforms</div><div style=""padding: 2em 0 0 0; font-size: .9em;"">P.S. If you arrive at the site, and the item is already sold out, we are sorry! But, at least you know you have great taste :)</div>"
		
		Call baf_sendmail()
	
	end if ' notify waiting list
	'======= END NOTIFY ON WAITING LIST THAT ITEM IS BACK IN STOCK ===========================


	'======= BEGIN INVENTORY COUNT NOTIFICATION ===========================
	if mailer_type = "inventory-count-notification" then
		display_interface = "no"
			
		mail_to_email = "jackie@bodyartforms.com"
		mail_to_name = "Jackie"
		
		mail_subject = "Inventory issue item #" & rsGetRegular.Fields.Item("ProductDetailID").Value
		mail_body = "<font face='verdana' size='2'>Item # " & rsGetRegular.Fields.Item("ProductDetailID").Value & " seems to have an issue with either the product being inactive, or it's a clearance/limited item that has not been moved over into the correct area."

		Call baf_sendmail()
		
	end if
	'======= END INVENTORY COUNT NOTIFICATION ===========================

	'======= BEGIN REPORT PHOTO =====================================================
	if mailer_type = "reported-photo" then 
		display_interface = "no"

		if Request.form("comments") <> "" then
			
			mail_to_email = "bafservice1@bodyartforms.com"
			mail_to_name = "Bodyartforms"
			mail_reply_email = Request.form("email")
			mail_reply_name = Request.form("name")
			
			mail_subject = "Reported photo"
			mail_body = "<b>Page url: </b><br/><a href='" + Request.form("url") + "'>" + Request.form("url") + "</a><br/><br /><b>Comments:</b><br/>" + Request.form("comments") + "<br/><br /><b>Reported photo:</b> " + Request.form("caption") + "<br/><br/>" + "<img style=""width:50%; max-width:500px"" src=""" + Request.form("img_src") + """ />"

		end if

		Call baf_sendmail()
	end if 
	'======= END REPORT PHOTO =====================================================	
%>
