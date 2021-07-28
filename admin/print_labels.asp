<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/bodyartforms_sql_ADMIN.asp" -->
<% 
Response.ContentType = "application/msword"


startrow  = 1

set rs = Server.CreateObject("ADODB.Recordset")
rs.ActiveConnection = MM_bodyartforms_sql_STRING
rs.Source = "SELECT * FROM sent_items  WHERE ship_code = 'paid' AND (Review_OrderError <> 1 OR  Review_OrderError IS NULL) AND (shipping_type = '4) Global basic ground' OR shipping_type = '3) Global priority mail') AND (shipped = 'Pending...' OR shipped = 'MISSING ITEM' OR shipped = 'SHIPPING BACKORDER' OR shipped = 'RETURN ENVELOPE' OR shipped = 'RESHIP PACKAGE' OR shipped = 'DEFECTIVE ITEM' OR shipped = 'INCORRRECT ITEM' OR shipped = 'ORDER ERROR' OR shipped = 'ORDER PROBLEM') ORDER BY shipping_type ASC, pay_method ASC, ID ASC"
rs.CursorLocation = 3 'adUseClient
rs.LockType = 1 'Read-only records
rs.Open()
rs_numRows = 0
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=FrontPage.Editor.Document>
<meta name=Generator content="Microsoft FrontPage 5.0">
<meta name=Originator content="Microsoft Word 10">
<title>Address labels</title>
<style>
<!--
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
     {mso-style-parent:"";
     margin-bottom:.0001pt;
     mso-pagination:widow-orphan;
     font-size:9.0pt;
     font-family:"Arial";
     mso-fareast-font-family:"Arial"; margin-left:0in; margin-right:0in; margin-top:0in}
span.SpellE
     {mso-style-name:"";
     mso-spl-e:yes}
@page Section1
     {size:8.5in 11.0in;
     margin:.5in 13.6pt 0in 13.6pt;
     mso-header-margin:.5in;
     mso-footer-margin:.5in;
     mso-paper-source:4;}
div.Section1
     {page:Section1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
     {mso-style-name:"Table Normal";
     mso-tstyle-rowband-size:0;
     mso-tstyle-colband-size:0;
     mso-style-noshow:yes;
     mso-style-parent:"";
     mso-padding-alt:0in 5.4pt 0in 5.4pt;
     mso-para-margin:0in;
     mso-para-margin-bottom:.0001pt;
     mso-pagination:widow-orphan;
     font-size:9.0pt;
     font-family:"Arial"}
</style>
<![endif]-->
</head>

<body lang=EN-US style='tab-interval:.5in'>

<div class=Section1>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 style='border-collapse:collapse;mso-padding-top-alt:0in;mso-padding-bottom-alt: 0in'>
<%for x=1 to startrow-1%>
 <tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes;page-break-inside:avoid;  height:1.0in'>
 
  <td width=252 style='width:189.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b  style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  </td>

  <td width=12 style='width:9.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
  <p class=MsoNormal style='margin-top:0in;margin-right:5.3pt;margin-bottom:  0in;margin-left:15pt;margin-bottom:.0001pt'><o:p>&nbsp;</o:p></p>
  </td>
    <td width=252 style='width:189.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b  style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  </td>

  <td width=12 style='width:9.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
  <p class=MsoNormal style='margin-top:0in;margin-right:5.3pt;margin-bottom:  0in;margin-left:15pt;margin-bottom:.0001pt'><o:p>&nbsp;</o:p></p>
  </td>
  
    <td width=252 style='width:189.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  <p class=MsoNormal align=left style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
  <b  style='mso-bidi-font-weight:normal'><span style='font-size:16.0pt'>&nbsp;</span></b></p>
  </td>

  </tr>
  <%next
%>
  
<%i = 0%>
<%do until rs.EOF%>

     <%If i = 0 then%>
          <tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes;page-break-inside:avoid;  height:1.0in'>
     <%elseIf i MOD 3 = 0 then%>
          </tr><tr style='mso-yfti-irow:0;mso-yfti-lastrow:yes;page-break-inside:avoid;  height:1.0in'>
     <%else%>
          <td width=12 style='width:9.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
          <p class=MsoNormal style='margin-top:0in;margin-right:5.3pt;margin-bottom:  0in;margin-left:15pt;margin-bottom:.0001pt'><o:p>&nbsp;</o:p></p>
          </td>
     <%end if%> 

     <td width=252 style='width:189.0pt;padding:0in .75pt 0in .75pt;height:1.0in'>
     <p class=MsoNormal align=center style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
     <t style='mso-bidi-font-weight:normal'><span style='font-size:9.0pt'><%=rs("customer_first")%>&nbsp;<%=rs("customer_last")%>&nbsp;(#<%=rs("ID")%>)</span></p>
     <p class=MsoNormal align=center style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
     <t  style='mso-bidi-font-weight:normal'><span style='font-size:9.0pt'><%=rs("address")%></span></p>
     <p class=MsoNormal align=center style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>	 <t  style='mso-bidi-font-weight:normal'><span style='font-size:9.0pt'><%=rs("address2")%></span></p>
     <p class=MsoNormal align=center style='margin-top:0in;margin-right:5.3pt;  margin-bottom:0in;margin-left:15pt;margin-bottom:.0001pt;text-align:left'>
     <t style='mso-bidi-font-weight:normal'><span style='font-size:9.0pt'><%=rs("city")%>,&nbsp;<%=rs("state")%><%=rs("province")%>&nbsp;&nbsp;<%=rs("zip")%><br><%=rs("country")%></span></p>
     </td>

<%
i = i +1
rs.movenext
loop%>
</tr>  
</table>

<p class=MsoNormal><span style='display:none;mso-hide:all'><o:p>&nbsp;</o:p></span></p>

</div>

</body>

</html>

<%
rs.Close()
Set rs = Nothing
%>
