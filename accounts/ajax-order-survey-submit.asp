<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<!--#include virtual="/functions/encrypt.asp"-->

<%

' Decrypt invoice #

Set objCrypt = Server.CreateObject("Bodyartforms.BAFCrypt")

password = "3uBRUbrat77V"
data = Request.form("invoiceid")

If len(data) > 5 then ' if
	decrypted = objCrypt.Decrypt(password, data)
end if

  if data <> decrypted then
	  InvoiceID = decrypted
  else
	  InvoiceID = 0
  end if

Set objCrypt = Nothing


' Check to make sure they don't hit refresh and get another $1 store credit
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT InvoiceID FROM TBL_Surveys WHERE InvoiceID = ?" 
objCmd.Prepared = true
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,InvoiceID))
Set rsNoDuplicate = objCmd.Execute()

' Only do the following if there already isn't an invoice # in the survey table
If rsNoDuplicate.EOF Then

if Request.Form("customerservice") <> "" then
	cs_rating = Request.Form("customerservice")
else
	cs_rating = 0
end if

'response.write "<br/>id: " & InvoiceID
For Each item In Request.Form
    Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
Next

	set objCmd = Server.CreateObject("ADODB.command")
	objCmd.ActiveConnection = DataConn
	objCmd.CommandText = "INSERT INTO TBL_Surveys(InvoiceID,SurveyCompleted,JewelrySelection,Prices,StockLevels,Website,Packaging,Items,Quality,Timeframe,CustomerService,Overall,NewJewelry,Comments,CustomerID,SelectionElaborate,PricesElaborate,StockElaborate,WebsiteElaborate,PackagingElaborate,ItemsElaborate,QualityElaborate,TimeframeElaborate,CSElaborate,OverallElaborate) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"
	objCmd.Prepared = true
	objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,InvoiceID))
	objCmd.Parameters.Append(objCmd.CreateParameter("SurveyCompleted",3,1,2,"1"))
	objCmd.Parameters.Append(objCmd.CreateParameter("selection",3,1,2,Request.Form("selection")))
	objCmd.Parameters.Append(objCmd.CreateParameter("pricing",3,1,2,Request.Form("pricing")))
	objCmd.Parameters.Append(objCmd.CreateParameter("stocklevels",3,1,2,Request.Form("stocklevels")))
	objCmd.Parameters.Append(objCmd.CreateParameter("experience",3,1,2,Request.Form("experience")))
	objCmd.Parameters.Append(objCmd.CreateParameter("packaging",3,1,2,Request.Form("packaging")))
	objCmd.Parameters.Append(objCmd.CreateParameter("items",3,1,2,Request.Form("items")))
	objCmd.Parameters.Append(objCmd.CreateParameter("quality",3,1,2,Request.Form("quality")))
	objCmd.Parameters.Append(objCmd.CreateParameter("delivery",3,1,2,Request.Form("delivery")))
	objCmd.Parameters.Append(objCmd.CreateParameter("customerservice",3,1,2,cs_rating))
	objCmd.Parameters.Append(objCmd.CreateParameter("overall",3,1,2,Request.Form("overall")))
	objCmd.Parameters.Append(objCmd.CreateParameter("new-jewelry",200,1,500,Request.Form("new-jewelry")))
	objCmd.Parameters.Append(objCmd.CreateParameter("Comments",200,1,500,Request.Form("Comments")))
	objCmd.Parameters.Append(objCmd.CreateParameter("CustomerID",3,1,10,CustID_Cookie))
	objCmd.Parameters.Append(objCmd.CreateParameter("SelectionElaborate",200,1,500,Request.Form("SelectionElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("pricingElaborate",200,1,500,Request.Form("pricingElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("stocklevelsElaborate",200,1,500,Request.Form("stocklevelsElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("experienceElaborate",200,1,500,Request.Form("experienceElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("packagingElaborate",200,1,500,Request.Form("packagingElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("itemsElaborate",200,1,500,Request.Form("itemsElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("qualityElaborate",200,1,500,Request.Form("qualityElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("deliveryElaborate",200,1,500,Request.Form("deliveryElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("customerserviceElaborate",200,1,500,Request.Form("customerserviceElaborate")))
	objCmd.Parameters.Append(objCmd.CreateParameter("overallElaborate",200,1,500,Request.Form("overallElaborate")))
	objCmd.Execute()


	'If they are a registered customer give them a $1 store credit for completing
	If Request.Cookies("ID") <> "" then

		set objCmd = Server.CreateObject("ADODB.command")
		objCmd.ActiveConnection = DataConn
		objCmd.CommandText = "UPDATE customers SET credits = credits + .5 WHERE customer_ID = ?" 
		objCmd.Prepared = true
		objCmd.Parameters.Append(objCmd.CreateParameter("customer_ID",3,1,10,CustID_Cookie))
		objCmd.Execute()

	End if ' if they are a registered customer

' Only send email if there's a low rating on the survey
If (cs_rating <= 3) OR Request.Form("selection") <= 3 OR Request.Form("pricing") <= 3 OR Request.Form("stocklevels") = 0 OR Request.Form("experience") <= 3 OR Request.Form("packaging") <= 3 OR Request.Form("items") = 0 OR Request.Form("quality") = 0 OR Request.Form("delivery") <= 3 OR Request.Form("overall") <= 3 OR Request.Form("new-jewelry") <> "" OR Request.Form("comments") <> "" Then



' Get invoice main invoice information
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT customer_first, customer_last, PackagedBy, email, date_sent FROM sent_items WHERE ID = ?" 
objCmd.Prepared = true
objCmd.Parameters.Append(objCmd.CreateParameter("ID",3,1,10,InvoiceID))
Set rsInvoice = objCmd.Execute()

' get item details if there is a low rating for that
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT * FROM dbo.QRY_OrderDetails WHERE InvoiceID = ? ORDER BY OrderDetailID ASC" 
objCmd.Prepared = true
objCmd.Parameters.Append(objCmd.CreateParameter("InvoiceID",3,1,10,InvoiceID))
Set rsGetOrderItems = objCmd.Execute()

ItemInfo = ""

While NOT rsGetOrderItems.EOF

	ItemInfo = ItemInfo & rsGetOrderItems.Fields.Item("qty").Value & "   " & rsGetOrderItems.Fields.Item("title").Value & " " & rsGetOrderItems.Fields.Item("ProductDetail1").Value & " " & rsGetOrderItems.Fields.Item("Gauge").Value & " " & rsGetOrderItems.Fields.Item("Length").Value & "<br/>"

rsGetOrderItems.MoveNext()
Wend 

If Request.Form("items") = 0 OR Request.Form("quality") = 0 Then
	ItemInfo = "<strong>Order details:</strong><br/>" & ItemInfo
else
	ItemInfo = ""
End if



If cs_rating <= 4 and cs_rating <> 0 Then
		Rate = ""
		For i = 1 To cs_rating
			Rate = Rate & "&#9733;"
		next
	
		CS = "<p><strong>Customer service feedback:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("customerserviceElaborate") & "<br/><br/><br/>"  
End if




If Request.Form("selection") <= 3 Then
		Rate = ""
		For i = 1 To Request.Form("selection")
			Rate = Rate + "&#9733;"
		next
	
		Selection = "<p><strong>Jewelry selection rating:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("selectionElaborate") & "<br/><br/><br/>"   
End if

If Request.Form("pricing") <= 3 Then 
		Rate = ""
		For i = 1 To Request.Form("pricing")
			Rate = Rate & "&#9733;"
		next
	
		Pricing = "<p><strong>Pricing rating:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("pricingElaborate") & "<br/><br/><br/>"  
End if

If Request.Form("experience") <= 3 Then
		Rate = ""
		For i = 1 To Request.Form("experience")
			Rate = Rate & "&#9733;"
		next
	
		Experience = "<p><strong>Website experience rating:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("experienceElaborate") & "<br/><br/><br/>"  
End if


If Request.Form("packaging") <= 3 Then
		Rate = ""
		For i = 1 To Request.Form("packagingElaborate")
			Rate = Rate & "&#9733;"
		next
	
		Packaging = "<p><strong>Packaging rating</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("PackagingElaborate") & "<br/><br/><br/>" 
End if

If Request.Form("delivery") <= 3 Then  
		Rate = ""
		For i = 1 To Request.Form("delivery")
			Rate = Rate & "&#9733;"
		next
	
		Delivery = "<p><strong>Delivery of order rating:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("deliveryElaborate") & "<br/><br/><br/>"
End if

If Request.Form("overall") <= 3 Then	  
		Rate = ""
		For i = 1 To Request.Form("overall")
			Rate = Rate & "&#9733;"
		next
	
		Overall = "<p><strong>Overall rating:</strong><br/><span style=""font-size: 1.8em;"">" & Rate & "</span><br/>" & Request.Form("overallElaborate") & "<br/><br/><br/>"		
End if

If Request.Form("stocklevels") = 0 Then
	  StockLevels = "<p><b>Items not in stock they wanted:</b><br/>" & Request.Form("stocklevelsElaborate") & "<br/><br/><br/>"
End if

If Request.Form("items") = 0 Then
	  Items = "<strong>Items that weren't correct:</strong><br/>" & Request.Form("itemsElaborate") & "<br/><br/><br/>"
End if

If Request.Form("quality") = 0 Then
	  Quality = "<strong>Dissatisfied with items:</strong><br/>" & Request.Form("qualityElaborate") & "<br/><br/><br/>"
End if

If Request.Form("new-jewelry") <> "" Then
		NewJewelry = "<strong>New types of jewelry customer would like to see:</strong><br/>" & Request.Form("new-jewelry") & "<br/><br/><br/>"
End if

If Request.Form("comments") <> "" Then
	Comments = "<strong>Additional comments:</strong><br/>" & Request.Form("comments") & "<br/><br/><br/>"
End If



'For Each item In Request.Form
'    Response.Write "Key: " & item & " - Value: " & Request.Form(item) & "<BR />"
'Next

mailer_type = "order-survey"
%>
<!--#include virtual="/emails/function-send-email.asp"-->
<!--#include virtual="/emails/email_variables.asp"-->
<%

End if ' ony send if anything is a low rating


End if 'If rsNoDuplicate.EOF 




DataConn.Close()
Set DataConn = Nothing
%>