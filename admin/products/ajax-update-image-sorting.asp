<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252" %>
<!--#include file="../../Connections/bodyartforms_sql_ADMIN.asp" -->
<%
item_count = 1

results = replace(replace(replace(request.form("id_array"),"[", ""), "]",""), """", "")
sort_array = split(results, ",")

for each x in sort_array
	
	if x <> 0 then
	
		set objCmd = Server.CreateObject("ADODB.Command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE tbl_images SET img_sort = ? WHERE img_id = ?"
		objCmd.Parameters.Append(objCmd.CreateParameter("img_sort",3,1,5,item_count))
		objCmd.Parameters.Append(objCmd.CreateParameter("img-id",3,1,15,x))
		objCmd.Execute()
	
	end if
	item_count = item_count + 1
	
next

DataConn.Close()
%>