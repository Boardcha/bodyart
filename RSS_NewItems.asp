<?xml version="1.0" encoding="ISO-8859-1"?>
<!--#include virtual="/Connections/sql_connection.asp" -->
<% 
Response.Buffer = true
Response.ContentType = "text/xml"

set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT ProductID, title, picture, date_added, material, description FROM jewelry WHERE (customorder <> N'yes') AND (jewelry <> N'save') AND (date_added <= '" & date()+21 & "') AND (date_added > '" & date()-45 & "') AND active = 1 ORDER BY date_added DESC, ProductID DESC"
Set rsGetRecords = objCmd.Execute()


Function return_RFC822_Date(myDate, offset)
   Dim myDay, myDays, myMonth, myYear
   Dim myHours, myMonths, mySeconds

   myDate = CDate(myDate)
   myDay = WeekdayName(Weekday(myDate),true)
   myDays = Day(myDate)
   myMonth = MonthName(Month(myDate), true)
   myYear = Year(myDate)
   myHours = zeroPad(Hour(myDate), 2)
   myMinutes = zeroPad(Minute(myDate), 2)
   mySeconds = zeroPad(Second(myDate), 2)

   return_RFC822_Date = myDay&", "& _
                                  myDays&" "& _
                                  myMonth&" "& _ 
                                  myYear&" "& _
                                  myHours&":"& _
                                  myMinutes&":"& _
                                  mySeconds&" "& _ 
                                  offset
End Function 
Function zeroPad(m, t)
   zeroPad = String(t-Len(m),"0")&m
End Function
%>
<rss version="2.0" xmlns:atom="http://www.w3.org/2005/Atom">
  <channel>
    <title>BAF new items</title>
    <link>http://www.bodyartforms.com/products.asp?new=Yes</link>
    <description>Newest items listed at BAF</description>
    <language>en-us</language>
    <pubDate><%=return_RFC822_Date((Now()), "GMT")%></pubDate>
    <category>Body jewelry</category>
    <ttl>20</ttl>
    <atom:link href="http://www.bodyartforms.com/RSS_NewItems.asp" rel="self" type="application/rss+xml" />
<%
With rsGetRecords
Do While Not.Eof
%>
    <item>
      <title><%=(rsGetRecords.Fields.Item("title").Value)%></title>
      <link>http://www.bodyartforms.com/productdetails.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%></link>
      <guid>http://www.bodyartforms.com/productdetails.asp?ProductID=<%=(rsGetRecords.Fields.Item("ProductID").Value)%></guid>
      <pubDate><%=return_RFC822_Date((rsGetRecords.Fields.Item("date_added").Value), "GMT")%></pubDate>
	  <description>&lt;br/&gt;&lt;img src='http://bodyartforms-products.bodyartforms.com/<%=(rsGetRecords.Fields.Item("picture").Value)%>' /&gt;</description>
    </item>
<%
.Movenext()
Loop
End With 

rsGetRecords.Close()
Set rsGetRecords = Nothing
%>
  </channel>
</rss>