<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM tbl_closeout_forms WHERE provider='USPS' AND CAST(date_created AS date) >= CAST(GETDATE() - 4 AS date) ORDER BY date_created DESC"
Set rsGetManifests = objCmd.Execute()
%>
<html>
<body>
              <table class="table table-striped table-hover table-sm small">
                    <thead class="thead-dark">
                      <tr>
                        <th scope="col">Date</th>
                        <th scope="col">Manifest Link</th>
                      </tr>
                    </thead>
                    <tbody>
                        <% i = 0
                         o = 0
                        While NOT rsGetManifests.EOF 
                         if FormatDateTime(date(),2) = FormatDateTime(rsGetManifests.Fields.Item("date_created").Value,2) AND i = 0 then %>
                        <tr class="table-success">
                                <th colspan="2">TODAY'S MANIFESTS</th>
                            </tr>
                        <% i = i + 1
                        end if %>
                        <%
                        if FormatDateTime(date(),2) <> FormatDateTime(rsGetManifests.Fields.Item("date_created").Value,2) AND o = 0 then %>
                       <tr class="table-warning">
                               <th colspan="2">OLDER MANIFESTS</th>
                           </tr>
                       <% o = o + 1
                       end if %>
                      <tr>
                        <th><%= rsGetManifests.Fields.Item("date_created").Value %></th>
                        <td>
                          <a class="d-block mb-2" href="data:application/pdf;base64, <%= rsGetManifests.Fields.Item("usps_base64_manifest_pdf").Value %>">Open manifest [Right click > Open in a new tab to view]</a>
                        </td>
                      </tr>

                    <% rsGetManifests.movenext()
                    wend %>
                    </tbody>
                  </table>
</body>
</html>
<%
DataConn.Close()
%>