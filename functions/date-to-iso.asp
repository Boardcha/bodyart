<%
'======== CONVERTS DATE INTO ISO FORMAT ====================

Function iso8601Date(dt)
    iso_dt = datepart("yyyy",dt) & "-"
    iso_dt = iso_dt & RIGHT("0" & datepart("m",dt),2) & "-"
    iso_dt = iso_dt & RIGHT("0" & datepart("d",dt),2)
    iso_dt = iso_dt & "T"
    iso_dt = iso_dt & RIGHT("0" & datepart("h",dt),2) & ":"
    iso_dt = iso_dt & RIGHT("0" & datepart("n",dt),2) & ":"
    iso_dt = iso_dt & RIGHT("0" & datepart("s",dt),2) & "Z"
    iso8601Date = iso_dt
End Function
%>