<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/functions/encrypt.asp"-->
<!--#include virtual="/Connections/authnet.asp"-->

<%
' Get customer info from database
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM customers  WHERE customer_ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
Set rsGetUser = objCmd.Execute()

cim_custid = rsGetUser.Fields.Item("cim_custid").Value


' Update email address in Auth.net CIM
strReq = "<?xml version=""1.0"" encoding=""utf-8""?>" _
& "<updateCustomerProfileRequest xmlns=""AnetApi/xml/v1/schema/AnetApiSchema.xsd"">" _
& MerchantAuthentication() _
& "<profile>" _
& "  <merchantCustomerId>" & CustID_Cookie & "</merchantCustomerId>" _
& "  <email>" & request.form("email") & "</email>" _
& "  <customerProfileId>" & cim_custid & "</customerProfileId>" _
& "</profile>" _
& "</updateCustomerProfileRequest>"
Set objUpdateProfile = SendApiRequest(strReq)

	If IsApiResponseSuccess(objUpdateProfile) Then
'  strEmail = objUpdateProfile.selectSingleNode("/*/api:profile/api:email").Text
  
  		' Update email in BAF database
		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET email = ? WHERE customer_ID = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("Email",200,1,50, request.form("email")))
		objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10, CustID_Cookie))
		objCmd.Execute()
	
		' Re-write cookies
		Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")
		
		password = "3uBRUbrat77V"
		data = CustID_Cookie
		
		encrypted = objCrypt.Encrypt(password, data)
		Response.Cookies("ID") = encrypted
		Response.Cookies("ID").Expires = DATE + 30
		
%>
{
	"status":"success"
}
<%		
	Else
%>
{
	"status":"fail"
}
<%
	End if ' If successful connection to Auth.Net CIM Profile



DataConn.Close()
Set DataConn = Nothing
%>