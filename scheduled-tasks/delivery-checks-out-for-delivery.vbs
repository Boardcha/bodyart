Dim objWinHttp, strURL
strURL = "https://bodyartforms.com/scheduled-tasks/delivery-check-out-for-delivery.asp"
Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
objWinHttp.Option(4) = &H3300 'Added by outserv to account for Nginx
objWinHttp.Open "GET", strURL
objWinHttp.Send
Set objWinHttp = Nothing