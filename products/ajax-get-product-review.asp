<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/Connections/chilkat.asp" -->
<!--#include virtual="/template/inc_includes_ajax.asp" -->
<%
set objCmd = Server.CreateObject("ADODB.command")
objCmd.ActiveConnection = DataConn
objCmd.CommandText = "SELECT  review, review_rating  FROM TBLReviews WHERE ReviewOrderDetailID = ?"
objCmd.NamedParameters = True
objCmd.Parameters.Append(objCmd.CreateParameter("@ProductID",3,1,10, request.form("id")))
Set rsReview = objCmd.Execute()

If Not rsReview.EOF Then
	review = rsReview("review")
	rating = rsReview("review_rating")
End If

set sbReview = Server.CreateObject("Chilkat_9_5_0.StringBuilder")
success = sbReview.Append(review)
success = sbReview.Encode("json","utf-8")
Response.Write "{""review"":""" & sbReview.GetAsString() & """, ""rating"":""" & rating & """}"

DataConn.Close()
Set DataConn = Nothing
%>