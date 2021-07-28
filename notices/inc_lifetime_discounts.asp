<%
if request.querystring("message_close") = "yes" then
	response.cookies("lifetime_discount") = "hide"
end if

' Show if cookie has not been set 
if request.cookies("lifetime_discount") = "" then
%>
<div id="notification" class="update_success">
<strong>10% permanent discount notice</strong>
<span style="font-size: 12px; font-weight:normal;">
<p>
Currently, anyone with a store account, who has spent over $275 (not including shipping) qualifies for 10% off every future order they place. However, starting <strong>October 1st</strong>, we will no longer be adding on that 10% discount for people that have not reached the status.
</p> 
<p>
For those of you who are close, but not quite there yet, no worries! You still have time to get your orders in and qualify for that permanent discount before it’s gone forever.
</p>
<a href="?message_close=yes"><strong>Close this message</strong></a>
</span>
</div>
<% end if %>