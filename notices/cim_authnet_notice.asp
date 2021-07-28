<%
' PERMANENTLY close security notice message box
if request.querystring("message_close") = "yes" then

				set objCmd = Server.CreateObject("ADODB.command")
				objCmd.ActiveConnection = DataConn
				objCmd.CommandText = "UPDATE customers SET notify_flag = 0 WHERE customer_ID = ?"
				objCmd.Parameters.Append(objCmd.CreateParameter("cim_custid",200,1,30,session("custID_account")))
				objCmd.Execute()
				
				showbox = "no"

end if
%>
<% if showbox = "yes" then %>
<div id="notification" class="update_success">
<div align="right"><strong><a href="?message_close=yes">[X] Close</a></strong></div>
Security upgrade & update notice
<span style="font-size: 12px; font-weight:normal;">
<br/><br/>
We've been working hard to improve the site and one of our recent changes is a security update to make storing your shipping and billing information easier and more secure for you.  With the new update:
<br/>
<br/>
- You can now store more than one credit card to your billing profile.<br/>
- You can store and choose between multiple shipping addresses with more ease<br/>
<br/>
To complete the update of your account, your information will need to be re-entered below. We are sorry for any inconvenience and if you have any questions please contact customer service <a href="contact.asp" target="_blank">here</a>.
<br/>
<br/>
<a href="?message_close=yes"><strong>Close this message &amp; do not display again</strong></a>
</span>
</div>
<% end if ' if showbox = yes %>