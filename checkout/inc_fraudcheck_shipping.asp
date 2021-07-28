<%
' Check to make sure a shipping option is selected. If not, then page needs to return back to checkout page so that user is prompted to select an option

if request.form("shipping-option") <> "" then


end if
%>