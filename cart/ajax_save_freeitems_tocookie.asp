<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="cart/generate_guest_id.asp"-->
<%

	if request.form("freegift1") <> "" then
		response.cookies("freegift1id") = request.form("freegift1")
	end if
	if request.form("freegift2") <> "" then
		response.cookies("freegift2id") = request.form("freegift2")
	end if	
	if request.form("freegift3") <> "" then
		response.cookies("freegift3id") = request.form("freegift3")
	end if	
	if request.form("freegift4") <> "" then
		response.cookies("freegift4id") = request.form("freegift4")
	end if	
	if request.form("freegift5") <> "" then
		response.cookies("freegift5id") = request.form("freegift5")
	end if		


DataConn.Close()
Set DataConn = Nothing
Set rs_getCart = Nothing
%>