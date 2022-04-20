<%
'Remove free gifts from stored cookies depending on order subtotal
	if fraudcheck_freegifts_subtotal < 30 then
		response.cookies("freegift3id") = ""
		response.cookies("freegift3id").expires = DateAdd("d",-1,Now())	
	end if	
	if fraudcheck_freegifts_subtotal < 50 then
		response.cookies("freegift4id") = ""
		response.cookies("freegift4id").expires = DateAdd("d",-1,Now())	
	end if	
	if fraudcheck_freegifts_subtotal < 75 then
		response.cookies("freegift5id") = ""
		response.cookies("freegift5id").expires = DateAdd("d",-1,Now())	
	end if	
	if fraudcheck_freegifts_subtotal < 100 then
		response.cookies("freegift6id") = ""
		response.cookies("freegift6id").expires = DateAdd("d",-1,Now())	
	end if	
	if fraudcheck_freegifts_subtotal < 150 then
		response.cookies("freegift7id") = ""
		response.cookies("freegift7id").expires = DateAdd("d",-1,Now())	
	end if	
%>