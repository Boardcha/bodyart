<!--#include virtual="/Connections/sql_connection.asp" -->
<!--#include virtual="/Connections/dhl-auth-v4.asp"-->
<script src="https://use.fortawesome.com/dc98f184.js"></script>
<%
if request("tracking") <> "" then
display_international_track_link = ""

'========== GET INVOICE BY TRACKING # TO FIND OUT WHAT DATE SENT WAS TO DETERMINE IF PACKAGE SHOULD BE TRACKED BY APIv2 OR APIv4 ===========================
    set objCmd = Server.CreateObject("ADODB.Command")
    objCmd.ActiveConnection = DataConn
    objCmd.CommandText = "SELECT date_sent, date_order_placed, international_tracking_num, international_tracking_url, country_UPSCode FROM sent_items LEFT OUTER JOIN TBL_Countries ON sent_items.country = TBL_Countries.Country where USPS_tracking = ? AND date_sent <> '' ORDER BY ID DESC"
    objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,request.querystring("tracking") ))
    set rsGetTrackType = objCmd.Execute()

' =================== REQUEST SHIPPING LABEL 
set rest = Server.CreateObject("Chilkat_9_5_0.Rest")

'  Connect to the REST server.
bTls = 1
port = 443
bAutoReconnect = 1
success = rest.Connect(dhl_api_url,port,bTls,bAutoReconnect)
success = rest.ClearAllQueryParams()

if NOT rsGetTrackType.EOF then
    if CDate(rsGetTrackType.Fields.Item("date_sent").Value) >= Cdate("10/6/2020") then '===== AFTER DHL SWITH TO APIv4 ON 10/6/2020 USE DHL PACKAGE ID TO TRACK ======
        success = rest.AddQueryParam("dhlPackageId", request.querystring("tracking"))
    else '===== BEFORE 10/6/2020 USE TRACKING # ==========
        success = rest.AddQueryParam("trackingId", request.querystring("tracking"))
    end if
else '===== if no records are found
    success = rest.AddQueryParam("trackingId", request.querystring("tracking"))
end if

success = rest.AddQueryParam("pickup", "5351961")


' Set the Authorization property to "Bearer <token>"
	set sbAuthHeaderVal = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
	success = sbAuthHeaderVal.Append("Bearer ")
	success = sbAuthHeaderVal.Append(db_dhl_access_token)
	rest.Authorization = sbAuthHeaderVal.GetAsString()

ResponseGetTrack = rest.FullRequestNoBody("GET","/tracking/v4/package")

    set JsonTracking = Server.CreateObject("Chilkat_9_5_0.JsonObject")
    JsonTracking.EmitCompact = 0
    JsonTracking.Load(ResponseGetTrack)
    'Response.Write "<pre>" & Server.HTMLEncode( JsonTracking.Emit()) & "</pre>"

if JsonTracking.IntOf("packages") <> "" then ' SUCCESS 

    if JsonTracking.StringOf("packages[0].package.expectedDelivery") <> "" then
        var_expected_delivery = MonthName(Month(JsonTracking.StringOf("packages[0].package.expectedDelivery")),True) & " " & Day(JsonTracking.StringOf("packages[0].package.expectedDelivery"))

        '======== WRITE EXPECTED DELIVERY DATE TO DB ===================================
        set objCmd = Server.CreateObject("ADODB.command")
        objCmd.ActiveConnection = DataConn
        objCmd.CommandText = "UPDATE sent_items SET estimated_delivery_date = ? WHERE USPS_tracking = ?"
        objCmd.Parameters.Append(objCmd.CreateParameter("estimated_delivery_date",133,1,30,JsonTracking.StringOf("packages[0].package.expectedDelivery") ))
        objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,request.querystring("tracking") ))
        objCmd.Execute()


    end if

if JsonTracking.StringOf("packages[0].events") <> "" then ' ONLY LOOP IF THERE ARE EVENTS TO SHOW 

Set eventsArray = JsonTracking.ArrayOf("packages[0].events")
eventsSize = eventsArray.Size

j = eventsSize - 1 '---- REVERSES ORDER OF EVENTS FROM NEWEST TO OLDEST

    For e = 0 To eventsSize - 1
        
        Set eventsObj = eventsArray.ObjectAt(j)

        if eventsObj.StringOf("primaryEventDescription") <> "" then
        
            var_message = "<div class=""row py-2 border-top""><div class=""col-5"">" & FormatDateTime(replace(eventsObj.StringOf("date"), "-", "/"), 1) & "<br>" &  REPLACE(FormatDateTime(LEFT(eventsObj.StringOf("time"), 5), 3), ":00", "") & "<br>" &  eventsObj.StringOf("location") & "</div><div class=""col-7"">" &  eventsObj.StringOf("primaryEventDescription") & " " & eventsObj.StringOf("secondaryEventDescription")  & "</div></div>" & var_message 

            ' code 447 = ARRIVED AT CUSTOMS
            ' code 361 = ARRIVAL IN COUNTRY
            if (eventsObj.StringOf("primaryEventId") = 360 or eventsObj.StringOf("primaryEventId") = 447 or eventsObj.StringOf("primaryEventId") = 361) and rsGetTrackType("international_tracking_num") <> "" then
                display_international_track_link = "yes"
            end if


            if j = 0 then '===== ONLY WRITE LAST EVENT =====
                '======== WRITE LAST EVENT TO DB ===================================
                set objCmd = Server.CreateObject("ADODB.command")
                objCmd.ActiveConnection = DataConn
                objCmd.CommandText = "UPDATE sent_items SET last_event_date = ?, last_status_id = ?, last_shipment_status = ? WHERE USPS_tracking = ?"
                objCmd.Parameters.Append(objCmd.CreateParameter("last_event_date",133,1,30, eventsObj.StringOf("date") ))
                objCmd.Parameters.Append(objCmd.CreateParameter("last_status_id",3,1,10, eventsObj.IntOf("primaryEventId") ))
                objCmd.Parameters.Append(objCmd.CreateParameter("last_shipment_status",200,1,500,  eventsObj.StringOf("primaryEventDescription") & " " & eventsObj.StringOf("secondaryEventDescription") ))
                objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,request.querystring("tracking") ))
                objCmd.Execute()
            end if

        end if '--- only show if there is an event description

        if eventsObj.IntOf("primaryEventId") = 598 then
            var_out_for_delivery = "yes"
            var_expected_delivery = ""
        end if

        if eventsObj.StringOf("primaryEventDescription") = "DELIVERED" then
            var_expected_delivery = ""
            var_out_for_delivery = ""
            var_delivered = "yes"
            var_delivery_date = MonthName(Month(eventsObj.StringOf("date")),True) & " " & Day(eventsObj.StringOf("date"))

            '======== WRITE DELIVERED STATUS TO DB ===================================
            set objCmd = Server.CreateObject("ADODB.command")
            objCmd.ActiveConnection = DataConn
            objCmd.CommandText = "UPDATE sent_items SET date_delivered = ?, packaged_delivered = ? WHERE USPS_tracking = ?"
            objCmd.Parameters.Append(objCmd.CreateParameter("date_delivered",135,1,30, eventsObj.StringOf("date") & " " & eventsObj.StringOf("time") ))
            objCmd.Parameters.Append(objCmd.CreateParameter("packaged_delivered",3,1,2, 1))
            objCmd.Parameters.Append(objCmd.CreateParameter("USPS_tracking",200,1,200,request.querystring("tracking") ))
            objCmd.Execute()

        end if

        j = j-1
    Next
    
else ' NO EVENTS FOUND
    var_message = "No tracking information available yet"

end if ' only show if there are events

else ' NO PACKAGE FOUND
    var_message = "No tracking information available yet"
end if



%>
<div class="font-weight-bold pb-2">
    Tracking number <%= request.querystring("tracking") %>
</div> 

<div class="container-fluid font-weight-bold my-4">
    <div class="row">
        <div class="col mr-2 pb-2 border-bottom border-success" style="border-width:3px!important">
            <i class="fa fa-check-circle fa-2x text-success"></i>
        </div>
        <div class="col mr-2 pb-2 text-center border-bottom border-success" style="border-width:3px!important">
            <i class="fa fa-check-circle fa-2x text-success"></i>
        </div>
        <% if var_expected_delivery <> "" then %>
        <div class="col mr-2 pb-2 text-center border-bottom border-secondary" style="border-width:3px!important">
            <i class="fa fa-package fa-2x text-secondary"></i>
        </div>
        <% end if %>
        <% if var_out_for_delivery = "yes" then %>
        <div class="col mr-2 pb-2 text-center border-bottom border-info" style="border-width:3px!important">
            <i class="fa fa-truck fa-2x text-info"></i>
        </div>
        <% end if %>
        <% if var_delivered = "yes" then 
                delivery_status_class = "28a745"
            else
                delivery_status_class = "D8D8D8"
            end if
        %>
        <div class="col text-right border-bottom" style="border-width:3px!important;border-color:#<%= delivery_status_class %>!important">
            <i class="fa fa-check-circle fa-2x" style="color:#<%= delivery_status_class %>"></i>
        </div>
      </div>
    <div class="row">
      <div class="col">
        Ordered
        <% if NOT rsGetTrackType.EOF then %>
        <br/>
        <%= MonthName(Month(rsGetTrackType.Fields.Item("date_order_placed").Value),True) & " " & Day(rsGetTrackType.Fields.Item("date_order_placed").Value) %>
        <% end if %>
      </div>
      <div class="col text-center">
        Shipped
        <% if NOT rsGetTrackType.EOF then %>
        <br/>
        <%= MonthName(Month(rsGetTrackType.Fields.Item("date_sent").Value),True) & " " & Day(rsGetTrackType.Fields.Item("date_sent").Value) %>
        <% end if %>
      </div>
      <% if var_expected_delivery <> "" then %>
      <div class="col text-center">
        Estimated Delivery<br/>
        <%= var_expected_delivery %>
      </div>
      <% end if %>
      <% if var_out_for_delivery = "yes" then %>
      <div class="col text-center">
        TODAY<br>  
        Out for delivery
      </div>
      <% end if %>
      <div class="col text-right">
        Delivered<br/>
        <%= var_delivery_date %>
      </div>
    </div>
  </div>

            <div class="container-fluid">
            <% if display_international_track_link = "yes" and rsGetTrackType("international_tracking_url") <> "" then
            international_link = replace(rsGetTrackType("international_tracking_url"), "{}", rsGetTrackType("international_tracking_num"))
            %>
                <div class="py-2 border-top">
                    
                    International tracking # <%= rsGetTrackType("international_tracking_num") %><br>
                    Now that your package has officially left the USA, you can use the button below to track the package inside your country<br>
                    <a class="btn btn-sm btn-secondary" href="<%=  international_link %>" target="_blank">Track your package</a>
                </div>
            <% end if %> 
        <%= var_message %>
    </div><!-- container-->
<% else %>
No tracking # provided
<% end if 'request("tracking") <> "" %>

