<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<link href="/CSS/baf.min.css?v=040220" id="lightmode" rel="stylesheet" type="text/css" />
<html>
    <html>
        <body>
        <% 
        
        Dim str, x
        var_cleaned_zip = ""
        var_cleaned_zip = "12345-1234"
        xx = Instr(var_cleaned_zip,"-")
        If xx Then var_cleaned_zip = Left(var_cleaned_zip,xx -1)
        %> 
        <% =var_cleaned_zip %> 
        </body>
        </html>
</html>