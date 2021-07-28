<%
' Functin that decodes into UTF-8
Public Function DecodeUTF8(decode_string)
  Set stmANSI = Server.CreateObject("ADODB.Stream")
  decode_string = decode_string & ""
  On Error Resume Next

  With stmANSI
    .Open
    .Position = 0
    .CharSet = "Windows-1252"
    .WriteText decode_string
    .Position = 0
    .CharSet = "UTF-8"
  End With

  DecodeUTF8 = stmANSI.ReadText
  stmANSI.Close

  If Err.number <> 0 Then
    lib.logger.error "str.DecodeUTF8( " & decode_string & " ): " & Err.Description
    DecodeUTF8 = decode_string
  End If
  On error Goto 0
End Function
%>