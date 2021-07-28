<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% response.Buffer = False %>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include file="../functions/encrypt.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
set DataConn = Server.CreateObject("ADODB.connection")
DataConn.Open = MM_bodyartforms_sql_STRING

Set rsGetNumbers = Server.CreateObject("ADODB.Recordset")
rsGetNumbers.ActiveConnection = DataConn
rsGetNumbers.Source = "SELECT TOP 10000 ID, cc_num, cc_num2 FROM dbo.sent_items WHERE (cc_num2 IS NULL) AND (NOT (cc_num IS NULL)) ORDER BY ID DESC"
rsGetNumbers.Open()


' Move through each record and decrypt the current credit card #
While NOT rsGetNumbers.EOF

ID = rsGetNumbers.Fields.Item("ID").Value

If rsGetNumbers.Fields.Item("cc_num").Value <> "" then

' ASPEncrypt decrypt code
Set CM = Server.CreateObject("Persits.CryptoManager")
Set Context = CM.OpenContext("", True)
Set Exp1Key = Context.CreateExponentOneKey
Set Blob = CM.CreateBlob
CM.LogonUser "www.bodyartforms.com", "ASPEncrypt", "t_zef$arafuge8h8Swecras9e3$67usp@wagasa*h$swasp_swetr&fasTet9ake"
Blob.LoadFromRegistry &H80000002, "Software\Key\AspEncrypt", "EncryptKeyLocation"
Set Key = Context.ImportKeyFromBlob( Exp1Key, Blob, cbtSimpleBlob )

Set EncryptedBlob = CM.CreateBlob
   EncryptedBlob.Base64 = rsGetNumbers.Fields.Item("cc_num").Value
   CCNumber = Key.DecryptText( EncryptedBlob )


' Convert to our built-in ASP encryption
Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

password = "3uBRUbrat77V"
data = CCNumber ' change this out to asp encrypt blob CCNumber
encrypted = objCrypt.Encrypt(password, data)
Set objCrypt = Nothing

' Update table with new credit card information
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "UPDATE sent_items SET cc_num2 = ? WHERE ID = ?"
objCmd.Parameters.Append(objCmd.CreateParameter("cc_num2",200,1,2000,encrypted))
objCmd.Parameters.Append(objCmd.CreateParameter("customerID",3,1,10,ID))
objCmd.Execute()
Set objCmd = Nothing


'Response.Write  ID & " : DONE   OLD # " & CCNumber & "  NEW # " & encrypted & "<br/>"

end if ' if cc_num field is not empty

rsGetNumbers.MoveNext()
Wend

DataConn.Close()
Set DataConn = Nothing
Set rsGetNumbers = Nothing
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
</body>
</html>
