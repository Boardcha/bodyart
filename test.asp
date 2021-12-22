<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/bootstrap-template/header-connection.asp" -->
<!--#include virtual="/bootstrap-template/header-scripts-and-css.asp" -->
<html>
    <html>
        <body>
        <% 
            set objCmd = Server.CreateObject("ADODB.command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "SELECT customer_ID, GETDATE() as 'current_date' FROM customers WHERE customer_ID = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("CustID_Cookie",3,1,10,CustID_Cookie))
            set rsGetDate = objCmd.Execute()

            response.write "<br>CLASSIC ASP NOW() OUTPUT -  " & NOW()
            response.write "<br>MSSQL SELECT GETDATE() AS OUTPUT -  " & rsGetDate("current_date")
        %> 
        </body>
        </html>
</html>