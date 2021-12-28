<%
function baf_sendmail()

%>
	<!--#include virtual="/Connections/chilkat.asp" -->
	<!--#include virtual="/Connections/aws-email.asp" -->
<%
	set mailman = Server.CreateObject("Chilkat_9_5_0.MailMan")
	' direct submission unauthenticate to baf1:
	'mailman.SmtpHost = "172.30.1.241"

	'========== Direct SES Submission
	mailman.SmtpHost = "email-smtp.us-east-1.amazonaws.com"
	mailman.SmtpPort = 587
	'==========  baf1smtpuser credentials in AWS. it only has DES submission privaledges
	mailman.SmtpUsername = var_mail_aws_access_key
	mailman.SmtpPassword = var_mail_aws_security_key



	set email = Server.CreateObject("Chilkat_9_5_0.Email")

	email.From = "Bodyartforms <bafservice1@bodyartforms.com>"

	email.Subject = mail_subject

	email.Body = email.Body + "<html>"

	if display_interface = "yes" then
		'------ HEADER -------
		email.Body = "<head><style>.expand {text-align: left;padding: 1em;display: inline-block;margin: 0;box-sizing: border-box} @media (max-width: 768px) {.expand, .column2{width:100%}} @media (min-width: 769px) {.expand, .column2{width:50%}}</style></head><body style='font-family:Arial, Helvetica, sans-serif;width:700px;'><font color='#fff'>-- Empowering Your Self Expression --</font><div style='width:100%;background-color:#696986;text-align:center;padding:15px 10px'><a href='https://bodyartforms.com/?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Logo%20click' target='_blank'><img src='https://bodyartforms.com/images/baf-logo-solid-white.png' height='30px'  alt='BAF logo'></a></div><div style='padding:.75em 0;font-size:.9em;font-weight:bold;text-align:center;background-color:#e6e6e6'><a style='color:#000;text-decoration:none' href='https://bodyartforms.com/products.asp?new=Yes&utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Shop%20navigation%20click' target='_blank'>SHOP</a><a style='margin:0 2.5em;color:#000;text-decoration:none' href='https://bodyartforms.com/account.asp?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Your%20account%20click' target='_blank'>ACCOUNT</a><a style='color:#000;text-decoration:none' href='https://bodyartforms.com/contact.asp?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Contact%20page%20click' target='_blank'>CONTACT US</a></div><div style='padding:2em 0 4em 0'>"
	end if

	'------ BODY DYNAMICALLY SET -------
	email.Body = email.Body + mail_body

	if display_interface = "yes" then
	'------ FOOTER -------
		email.Body = email.Body + "</div><div style='text-align:center'><i>Email us at <a href='mailto:help@bodyartforms.com'>help@bodyartforms.com</a> or call us at (877) 223-5005<br>Monday thru Friday 9am - 5pm Central Time</i></div><br><div style='background-color:#696986;color:#fff;width:100%;margin:0;vertical-align:top'><div class='expand' style='vertical-align:top;font-size:1.1em'><a href='https://bodyartforms.com/contact.asp?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Contact%20page%20click' target='_blank' style='color:#fff;text-decoration:none;padding: .25em .5em;display:block'>Contact Us</a><a href='https://bodyartforms.com/faqs.asp?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=FAQs%20click' target='_blank' style='color:#fff;text-decoration:none;padding:.25em .5em;display:block'>FAQs & Support</a><a href='https://bodyartforms.com/returns.asp?utm_source=" & google_utmsource & "&utm_medium=Site%20email&utm_campaign=Return%20policy%20click' target='_blank' style='color:#fff;text-decoration:none;padding:.25em .5em;display:block'>Return Policy</a></div><div  class='expand' ><div style='font-size:1.2em;font-weight:bold'>Get notified about sales!</div><a href='https://manage.kmail-lists.com/subscriptions/subscribe?a=VnRhsk&g=UTEZqk' style='background-color:#ffffff;border-radius:.4em;padding:.3em .5em;text-decoration:none;color:#000;display:inline-block' >Sign up for our newsletter</a><br/><br/><a href='http://instagram.com/bodyartforms' target='_blank'><img src='https://bodyartforms.com/images/icons/instagram-white.png' style='height:30px;padding:0 .2em'/></a><a href='http://www.facebook.com/pages/Bodyartforms/149344708430326' target='_blank'><img src='https://bodyartforms.com/images/icons/facebook-white.png' style='height:30px;padding:0 .2em'/></a></div></div><img src='https://www.google-analytics.com/collect?v=1&tid=UA-32113869-1&cid=555&t=event&ec=Site%20emails&ea=Opened%20" & google_utmsource & "'/></body>"
	end if

	email.Body = email.Body + "</html>"

	success = email.AddTo(mail_to_name,mail_to_email)
	'Add on CCs
	if cc1_email <> "" then
		success = email.AddTo(cc1_name,cc1_email)
	end if
	email.ReplyTo = mail_reply_email

	success = mailman.SendEmail(email)
	If (success <> 1) Then
		'Response.Write "<pre>" & Server.HTMLEncode( mailman.LastErrorText) & "</pre>"
		'Response.End
	End If

end function
%>