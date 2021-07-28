<%@LANGUAGE="VBSCRIPT"%>
<%
	Session("consumerKey") = "1hlfivo6em3zdq0yxk7re2ag"
	Session("consumerSecret") = "0ftq6t2rnm"
	Session("oauthCallback") = "https://127.0.0.1/admin/etsy/etsy-oauth1-step2.asp"
	Session("chilkatUnlockCode") = "BDYART.CB1042021_UB3WE6ih77lJ"
	
	Session("requestTokenUrl") = "https://openapi.etsy.com/v2/oauth/request_token?scope=transactions_r%20transactions_w%20listings_r%20listings_w%20listings_d"
    Session("authorizeTokenUrl") = "https://www.etsy.com/oauth/signin"
    Session("accessTokenUrl") = "https://openapi.etsy.com/v2/oauth/access_token"
	
	set http = Server.CreateObject("Chilkat_9_5_0.Http")
	success = http.UnlockComponent(Session("chilkatUnlockCode"))
	If (success <> 1) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If
	http.OAuth1 = 1
	http.OAuthConsumerKey = Session("consumerKey")
	http.OAuthConsumerSecret = Session("consumerSecret")
	http.OAuthCallback = Session("oauthCallback")
	set req = Server.CreateObject("Chilkat_9_5_0.HttpRequest")
    set resp = http.PostUrlEncoded(Session("requestTokenUrl"), req)
	If (resp Is Nothing ) Then
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	ElseIf (resp.StatusCode = 200) Then
		' Success
		set hashTab = Server.CreateObject("Chilkat_9_5_0.Hashtable")
		hashTab.AddQueryParams(resp.BodyStr)
		'Response.Write "BodyStr: [" & resp.BodyStr & "]<p>"
        Session("oauth_token") = hashTab.LookupStr("oauth_token")
        Session("oauth_token_secret") = hashTab.LookupStr("oauth_token_secret")
		'Response.Write "oauth_token: [" & Session("oauth_token") & "]<p>"
		'Response.Write "oauth_token_secret: [" & Session("oauth_token_secret") & "]<p>"
	    Response.Redirect Session("authorizeTokenUrl") + "?oauth_token=" + Session("oauth_token")
	Else
	    Response.Write "<pre>" & Server.HTMLEncode( http.LastErrorText) & "</pre>"
	End If
%>