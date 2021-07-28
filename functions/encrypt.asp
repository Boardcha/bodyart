<SCRIPT LANGUAGE="JScript" RUNAT="Server" SRC="md5.js"></SCRIPT>
<%
	Function EncryptString(TheString)
		sResult = hex_md5(TheString) 
		EncryptString = sResult
	End Function
%>