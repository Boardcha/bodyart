<%@LANGUAGE="VBSCRIPT" %>
<!--#include virtual="/Connections/bodyartforms_sql_ADMIN.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = MM_bodyartforms_sql_STRING
objCmd.CommandText = "SELECT * FROM tbl_closeout_forms WHERE provider='DHL' AND CAST(date_created AS date) >= CAST(GETDATE() - 4 AS date) ORDER BY date_created DESC"
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


                        '======== SENT REQUEST TO DHL FOR MANIFEST ===============================
                        set rest = Server.CreateObject("Chilkat_9_5_0.Rest")
                        
                        '  Connect to the REST server.
                        bTls = 1
                        port = 443
                        bAutoReconnect = 1
                        success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
                        success = rest.AddHeader("Content-Type","application/json")
                        
                        ' Set the Authorization property to "Bearer <token>"
                          set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
                          success = sbAuthHeaderVal.Append("Bearer ")
                          success = sbAuthHeaderVal.Append(db_dhl_access_token)
                            rest.Authorization = sbAuthHeaderVal.GetAsString()
                            
                            ResponseRequestManifest = rest.FullRequestNoBody("GET","/shipping/v4/manifest/" & dhl_production_pickup_num & "/" & rsGetManifests.Fields.Item("requestId").Value)
                        
                            set JsonManifest = Server.CreateObject("Chilkat_9_5_0.JsonObject")
                            JsonManifest.EmitCompact = 0
                            JsonManifest.Load(ResponseRequestManifest)
                        %>
                        <%
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
                          <a href="<%= rsGetManifests.Fields.Item("closeout_url").Value %>" target="_blank"><%= rsGetManifests.Fields.Item("closeout_url").Value %></a>

                          <% 'Response.Write "<pre>" & Server.HTMLEncode( JsonManifest.Emit()) & "</pre>" %>

                          <%
                          if JsonManifest.StringOf("manifests") <> "" then ' ONLY LOOP IF THERE ARE MANIFESTS TO SHOW 

                          Set manifestArray = JsonManifest.ArrayOf("manifests")
                          manifestSize = manifestArray.Size
                          i = 0
                          Do While i < manifestSize
                                  
                                  Set manifestObj = manifestArray.ObjectAt(i)

                                  if manifestObj.StringOf("isInternational") = "true" then
                                    manifestType = "international"
                                  else
                                    manifestType = "domestic"
                                  end if

                          %>
                          <a class="d-block mb-2" href="data:application/pdf;base64, <%= manifestObj.StringOf("manifestData") %>">Open <%= manifestType %> manifest [Right click > Open in a new tab to view]</a>
                          <%                              

                                  i = i + 1
                            Loop
                              
                          end if ' only show if there are manifests
                          %>
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