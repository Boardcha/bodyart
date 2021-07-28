<%
if request.cookies("currency") = "USD" or request.cookies("currency") = "" then
currency_img = "usa.png"
elseif request.cookies("currency") = "CAD" then
currency_img = "canada.png"
elseif request.cookies("currency") = "GBP" then
currency_img = "uk.png"
elseif request.cookies("currency") = "EUR" then
currency_img = "euro.png"
elseif request.cookies("currency") = "AUD" then
currency_img = "australia.png"
elseif request.cookies("currency") = "JPY" then
currency_img = "japan.png"
elseif request.cookies("currency") = "NZD" then
currency_img = "nz.png"
elseif request.cookies("currency") = "DKK" then
currency_img = "denmark.png"
end if
if request.cookies("currency") = "USD" or request.cookies("currency") = "" then
currency_text = "$ USD"
else
currency_text = request.cookies("currency-symbol") & " " & request.cookies("currency")
end if
%>