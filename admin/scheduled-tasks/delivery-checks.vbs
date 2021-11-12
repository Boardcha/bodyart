Dim objWinHttp, strURL
strURL = "https://bodyartforms.com/emails/delivery-check.asp"
Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Open "GET", strURL
objWinHttp.Send
Set objWinHttp = Nothing