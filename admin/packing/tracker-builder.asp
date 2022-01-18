<%
if instr(rsGetInvoice.Fields.Item("shipping_type").Value, "DHL") > 0 then
    var_tracking = var_tracking & "https://bodyartforms.com/dhl-tracker.asp?tracking=" & rsGetInvoice.Fields.Item("USPS_tracking").Value
end if 
if instr(rsGetInvoice.Fields.Item("shipping_type").Value, "USPS") > 0 then
    var_tracking = var_tracking & "https://tools.usps.com/go/TrackConfirmAction?qtc_tLabels1=" & rsGetInvoice.Fields.Item("USPS_tracking").Value
end if
if instr(rsGetInvoice.Fields.Item("shipping_type").Value, "UPS") then
    var_tracking = var_tracking & "https://www.ups.com/track?loc=en_US&requester=QUIC&tracknum=" & rsGetInvoice.Fields.Item("UPS_tracking").Value & "/"
end if
%>