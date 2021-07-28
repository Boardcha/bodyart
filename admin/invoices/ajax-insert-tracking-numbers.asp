<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
' =========================================================================================
' Upload .csv file for Etsy orders and order details

' Link below to install Microsoft Access Database Engine 2016 Redistributable
' https://www.microsoft.com/en-us/download/details.aspx?id=54920
' =========================================================================================

Dim strConn, conn, rs

' --------------------------------------------
' Upload CSV file
' --------------------------------------------
Set Upload = Server.CreateObject("Persits.Upload")
	Upload.OverwriteFiles = True
    Upload.Save("C:\inetpub\wwwroot\bootstrap-svn\admin\uploads") 'LOCALHOST TESTING
    'Upload.Save("C:\inetpub\bootstrap-baf\admin\uploads")  'LIVE SERVER


    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
    Server.MapPath("\admin\uploads\") & ";Extended Properties=""text;HDR=Yes;FMT-Delimited"";"
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.Open strConn

    For Each File in Upload.Files
      ' ----------- RENAME FILE
      File.Copy "C:\inetpub\wwwroot\bootstrap-svn\admin\uploads\tracking-numbers.csv"   'LOCALHOST
      'File.Copy "C:\inetpub\bootstrap-baf\admin\uploads\tracking-numbers.csv"  'LIVE SERVER
      File.Delete
    next ' for each file uploaded
    

    Set rsGetOrders = Server.CreateObject("ADODB.recordset")
    rsGetOrders.open "SELECT * FROM [tracking-numbers.csv]", conn 

    'For Each header In rsGetOrders.Fields
        'Response.Write("column: " & header.Name)
    'Next

    Function TrackStartsWith(string1, string2)
      TrackStartsWith = InStr(1, string1, string2, 1) = 1
    End Function

    while not rsGetOrders.eof
        var_invoice_id = rsGetOrders.Fields.Item("refCustom1").Value    
        var_tracking_number = rsGetOrders.Fields.Item("trackingNumber").Value

        If TrackStartsWith(var_tracking_number, "1Z") Then
          track_column = "UPS_tracking"
        else
          track_column = "USPS_tracking"
        End If
       
        set objCmd = Server.CreateObject("ADODB.Command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE sent_items SET " & track_column & " = ? WHERE ID = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("tracking",200,1,100,var_tracking_number))
        objCmd.Parameters.Append(objCmd.CreateParameter("invoiceid",3,1,15,var_invoice_id))
        objCmd.Execute()
        Set objCmd = Nothing

    rsGetOrders.movenext
    wend
  %>
  {
    "status":"success",
    "reason":""
}
<%

conn.close()
DataConn.Close()
%>