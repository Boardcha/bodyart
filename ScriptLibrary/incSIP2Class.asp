<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'------------------------------------------
' Smart Image Processor 2
' Copyright 2006 (c) DMXzone
' Version: 2.0.2
'------------------------------------------
'Option Explicit

Const IP_Version = "2.0.2"

Const PI = 3.1415926535897932

' Available components in order of detection
'------------------------------------------
' AspNet
' GraphicsMill (http://www.comobjects.net/Products/GraphicsMill)
' ImageX (http://www.fathsoft.com)
' ShotGraph (http://www.shotgraph.com)
' ImgWriter (http://www.softartisans.com/imgwriter.html)
' AspJpeg (http://www.aspjpeg.com)
' AspImage (http://www.serverobjects.com/comp/Aspimage.htm)
' SmartImage (http://www.aspsmart.com/aspSmartImage)
' PictureProcessor
' AspThumb (http://www.brizsoft.com/asp/thumb)
'------------------------------------------

' The main Class
Class ImageProcessor
  
  Private m_Debug
  Private m_Font
  Private m_Size
  Private m_Color
  Private m_Bold
  Private m_Italic
  Private m_Underline
  Private m_ScriptFolder
  Private m_Components
  Private m_Object
  Private m_File
  Private m_Overwrite
  Private m_Mask
  
  Private m_FileCollection
  Private m_Source
  Private m_UploadFields
  Private m_Rs
  Private m_RsField
  Private m_Folder
  Private m_HasFiles
  Private m_Index
  
  Private m_Pattern
  
  ' Turn debugging on/off
  Public Property Let Debug(showDebug)
    m_Debug = CBool(showDebug)
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    If m_Object Is Nothing Then
      Width = 0
    Else
      Width = m_Object.Width
    End If
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    If m_Object Is Nothing Then
      Height = 0
    Else
      Height = m_Object.Height
    End If
  End Property
  
  Public Property Get FileName()
    FileName = m_File
  End Property
  
  Public Property Get HasFiles()
    HasFiles = m_HasFiles
  End Property
  
  ' Returns the font used for text
  Public Property Get FontFamily()
    FontFamily = m_Font
  End Property
  
  ' Set the font to use for text
  Public Property Let FontFamily(font)
    m_Font = CStr(font)
  End Property
  
  ' Returns the fontsize used for text
  Public Property Get FontSize()
    FontSize = m_Size
  End Property
  
  ' Set the fontsize to use for text
  Public Property Let FontSize(size)
    m_Size = CInt(size)
  End Property
  
  ' Returns the text color used for text
  Public Property Get FontColor()
    FontColor = m_Color
  End Property
  
  ' Set the text color to use for text
  Public Property Let FontColor(color)
    Dim regEx
    Set regEx = New RegExp
    regEx.Pattern = "#[A-Fa-f0-9]{6}"
    If regEx.Test(CStr(color)) Then
      m_Color = CStr(color)
    End If
  End Property
  
  Public Property Get VbColor()
    Dim clr
    clr = RgbColor
    VbColor = RGB(clr(0), clr(1), clr(2))
  End Property
  
  Public Property Get RgbColor()
    Dim R, G, B
    R = Eval("&H" & Mid(m_Color, 2, 2))
    G = Eval("&H" & Mid(m_Color, 4, 2))
    B = Eval("&H" & Mid(m_Color, 6, 2))
    RgbColor = Array(R, G, B)
  End Property
  
  ' Returns if text uses bold
  Public Property Get Bold()
    Bold = m_Bold
  End Property
  
  ' Set if text must be bold
  Public Property Let Bold(useBold)
    m_Bold = CBool(useBold)
  End Property
  
  ' Returns if text uses italic
  Public Property Get Italic()
    Italic = m_Italic
  End Property
  
  ' Set if text must be italic
  Public Property Let Italic(useItalic)
    m_Italic = CBool(useItalic)
  End Property
  
  ' Returns if text uses underline
  Public Property Get Underline()
    Underline = m_Underline
  End Property
  
  ' Set if text must use underline
  Public Property Let Underline(useUnderline)
    m_Underline = CBool(useUnderline)
  End Property
  
  ' Return overwrite value
  Public Property Get Overwrite()
    Overwrite = m_Overwrite
  End Property
  
  ' Set the overwrite option
  Public Property Let Overwrite(bOverwrite)
    m_Overwrite = CBool(bOverwrite)
  End Property
  
  ' Return the current mask
  Public Property Get Mask()
    Mask = m_Mask
  End Property
  
  ' Set a new mask
  Public Property Let Mask(str)
    m_Mask = CStr(str)
  End Property
  
  Public Property Get Source()
    Source = m_Source
  End Property
  
  Public Property Let Source(str)
    m_Source = CStr(str)
  End Property
  
  Public Property Get UploadFields()
    UploadFields = m_UploadFields
  End Property
  
  Public Property Let UploadFields(str)
    m_UploadFields = CStr(str)
  End Property
  
  Public Property Let Recordset(rs)
    Set m_Rs = rs
  End Property
  
  Public Property Let RecordsetField(field)
    m_RsField = CStr(field)
  End Property
  
  Public Property Let Folder(f)
    m_Folder = CStr(f)
  End Property
  
  Public Property Get ScriptFolder()
    ScriptFolder = m_ScriptFolder
  End Property
  
  Public Property Let ScriptFolder(str)
    m_ScriptFolder = CStr(str)
    
    If LCase(Request.QueryString("sipinfo")) <> "" Then
      ShowInfo
    End If
    
    If Request.QueryString("sipdebug") <> "" Then
      DebugStart
      m_Debug = True
    End If
  End Property
  
  ' Returns the current component
  Public Property Get Component()
    Set Component = m_Object
  End Property
  
  ' Set the component to use
  Public Property Let Component(str)
    If LCase(str) = "auto" Then
      If Not DetectComponent() Then
        WriteError "No component detected", "There is no supported component detected."
      End If
      WriteDebug m_Object.ToString, "Detected"
    Else
      If Not TestComponent(str) Then
        WriteError "Component " & str & " could not be loaded", "Please make sure that the component " & str & " is installed."
      End If
    End If
  End Property
  
  ' Set some defaults
  Private Sub Class_Initialize()
    Set m_Object = Nothing
    Set m_Rs = Nothing
    m_Pattern = "jpg|jpeg|gif|png|tif"
    m_Debug = False
    m_Font = "Arial"
    m_Size = 14
    m_Color = "#000000"
    m_Bold = False
    m_Italic = False
    m_Underline = False
    m_ScriptFolder = ""
    m_Overwrite = False
    m_Mask = "##path####filename##"
    m_Source = ""
    m_UploadFields = ""
    m_RsField = ""
    m_Folder = ""
    m_HasFiles = False
    m_Index = 0
    m_Components = Array("AspNet", _
                         "GraphicsMill", _
                         "ImageX", _
                         "ShotGraph", _
                         "ImgWriter", _
                         "AspJpeg", _
                         "AspImage", _
                         "SmartImage", _
                         "PictureProcessor", _
                         "AspThumb")
  End Sub
  
  ' Cleanup resources
  Private Sub Class_Terminate()
    DebugEnd
    Set m_Object = Nothing
    Set m_FileCollection = Nothing
  End Sub
  
  Public Sub GetFiles()
    Dim regEx, index, file, ext, i, col
    
    Set m_FileCollection = Server.CreateObject("Scripting.Dictionary")
    
    Set regEx = New RegExp
    regEx.Pattern = m_Pattern
    regEx.IgnoreCase = True
    
    index = 0
    
    WriteDebug "Getting file collection", "Source is " & m_Source
    
    Select Case LCase(m_Source)
      Case "upload":
        'Get files from upload
        Dim regEx2
        If m_UploadFields = "" Then m_UploadFields = ".*"
        Set regEx2 = New RegExp
        regEx2.Pattern = m_UploadFields
        regEx2.IgnoreCase = True
        If (CStr(Request.QueryString("GP_upload")) <> "") And IsObject(UploadRequest) Then
          Dim uploadFields, uploadKeys
          uploadFields = UploadRequest.Items
          uploadKeys = UploadRequest.Keys
          For i = 0 To UBound(uploadFields)
            If uploadFields(i).Exists("FileName") Then
              If uploadFields(i).Item("FileName") <> "" And regEx2.Test(uploadKeys(i)) Then
                If Not uploadFields(i).Exists("Field") Then
                  uploadFields(i).Add "Field", uploadKeys(i)
                End If
                file = uploadFields(i).Item("FileName")
                ext = Mid(file, InStrRev(file, ".") + 1)
                If regEx.Test(ext) Then
                  m_FileCollection.Add index, uploadFields(i)
                  col = col & uploadFields(i).Item("FileName") & "<br>"
                  index = index + 1
                End If
              End If
            End If
          Next
        End If
        Set regEx2 = Nothing
        
      Case "recordset":
        'Get files from recordset
        While Not m_Rs.EOF
          file = m_Rs.Fields.Item(m_RsField).Value
          ext = Mid(file, InStrRev(file, ".") + 1)
          If regEx.Test(ext) Then
            m_FileCollection.Add index, file
            col = col & file & "<br>"
            index = index + 1
          End If
        Wend
        
      Case "folder":
        'Get files from folder
        Dim fso, f, fc, fl
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        If InStr(m_Folder, ":\") = 0 Then m_Folder = Server.MapPath(m_Folder)
        Set f = fso.GetFolder(m_Folder)
        Set fc = f.Files
        For Each fl In fc
          ext = Mid(fl.Name, InStrRev(fl.Name, ".") + 1)
          If regEx.Test(ext) Then
            m_FileCollection.Add index, fl.Path
            col = col & fl.Path & "<br>"
            index = index + 1
          End If
        Next
        Set fc = Nothing
        Set f = Nothing
        Set fso = Nothing
    End Select
    
    If index > 0 Then
      WriteDebug "Files added to collection", col
    Else
      WriteDebug "Files added to collection", "No files where added"
    End If
    
    If index > 0 Then m_HasFiles = True
    m_Index = 0
    Set regEx = Nothing
  End Sub
  
  Public Sub LoadFromSource()
    Select Case LCase(m_Source)
      Case "upload":
        'Load file from upload
        Dim path
        If InStr(m_Path, """") > 0 Then
          path = Eval(pau_thePath)
        Else
          path = pau_thePath
        End If
        Load path & "/" & m_FileCollection.Item(m_Index).Item("FileName")
      Case "recordset":
        'Load file from recordset
        Load m_FileCollection.Item(m_Index)
      Case "folder":
        'Load file from folder
        Load m_FileCollection.Item(m_Index)
    End Select
  End Sub
  
  Public Sub MoveNext()
    If LCase(m_Source) = "upload" And m_Overwrite Then
      Dim path, oldValue, lastPos, file, ext, fso, f
      
      oldValue = m_FileCollection.Item(m_Index).Item("Value")
      path = Left(oldValue, InStrRev(oldValue, "/"))
      file = Mid(m_File, InStrRev(m_File, "\") + 1)
      ext = Mid(file, InStrRev(file, ".") + 1)
      m_FileCollection.Item(m_Index).Item("Value") = path & file
      m_FileCollection.Item(m_Index).Item("FileName") = file
      
      WriteDebug "Updating fields", "old value = " & oldValue & "<br>new value = " & path & file
      
      If IsObject(pau) Then
        On Error Resume Next
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        Set f = fso.GetFile(m_File)
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).Value = path & file
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).FileName = file
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).ContentType = "image/" & ext
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).Width = Width
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).Height = Height
        pau.Fields(m_FileCollection.Item(m_Index).Item("Field")).Size = f.Size
        Set f = Nothing
        Set fso = Nothing
        On Error GoTo 0
      End If
    End If
    
    m_Index = m_Index + 1
    If m_Index = m_FileCollection.Count Then m_HasFiles = False
  End Sub
  
  ' Load image from file
  Public Sub Load(file)
    Dim regEx, ext
    Set regEx = New RegExp
    regEx.Pattern = m_Pattern
    regEx.IgnoreCase = True
    m_File = CStr(file)
    If file <> "" Then
       ext = LCase(Mid(m_File, InStrRev(m_File, ".") + 1))
       If regEx.Test(ext) Then
        If m_Object Is Nothing Then
          Component = "auto"
        End If
        If InStr(m_File, ":\") = 0 Then m_File = Server.MapPath(m_File)
        m_Object.Load m_File
        End If
    End If
    WriteDebug "Load file", "file = " & m_File
    Set regEx = Nothing
  End Sub
  
  ' Save image as JPEG
  Public Sub SaveJPEG(quality)
    Dim curPath, curFilename, curName, newFilename
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    curPath = Left(m_File, InStrRev(m_File, "\"))
    curFilename = Right(m_File, Len(m_File) - InStrRev(m_File, "\"))
    curName = Left(curFilename, InStrRev(curFilename, ".") - 1)
    newFilename = m_Mask
    newFilename = Replace(newFilename, "##path##", curPath)
    newFilename = Replace(newFilename, "##filename##", curFilename)
    newFilename = Replace(newFilename, "##name##", curName)
    newFilename = Left(newFilename, InStrRev(newFilename, ".")) & "jpg"
    If InStr(newFilename, ":\") = 0 Then newFilename = Server.MapPath(newFilename)
    m_Object.SaveJPEG newFilename, quality
    WriteDebug "Write JPEG", "file = " & newFilename
    
    If LCase(m_File) <> LCase(newFilename) And m_Overwrite Then
      Set fso = Server.CreateObject("Scripting.FileSystemObject")
      fso.DeleteFile m_File, True
      WriteDebug "Delete file", "file = " & m_File
      Set fso = Nothing
    End If
    
    m_File = newFilename
  End Sub
  
  ' Save image as PNG
  Public Sub SavePNG()
    Dim curPath, curFilename, curName, newFilename
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    curPath = Left(m_File, InStrRev(m_File, "\"))
    curFilename = Right(m_File, Len(m_File) - InStrRev(m_File, "\"))
    curName = Left(curFilename, InStrRev(curFilename, ".") - 1)
    newFilename = m_Mask
    newFilename = Replace(newFilename, "##path##", curPath)
    newFilename = Replace(newFilename, "##filename##", curFilename)
    newFilename = Replace(newFilename, "##name##", curName)
    newFilename = Left(newFilename, InStrRev(newFilename, ".")) & "png"
    If InStr(newFilename, ":\") = 0 Then newFilename = Server.MapPath(newFilename)
    m_Object.SavePNG newFilename
    WriteDebug "Write PNG", "file = " & newFilename
    
    If LCase(m_File) <> LCase(newFilename) And m_Overwrite Then
      Set fso = Server.CreateObject("Scripting.FileSystemObject")
      fso.DeleteFile m_File, True
      WriteDebug "Delete file", "file = " & m_File
      Set fso = Nothing
    End If
    
    m_File = newFilename
  End Sub
  
  ' Save image as GIF
  Public Sub SaveGIF(palette, dither, colors)
    Dim curPath, curFilename, curName, newFilename
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    curPath = Left(m_File, InStrRev(m_File, "\"))
    curFilename = Right(m_File, Len(m_File) - InStrRev(m_File, "\"))
    curName = Left(curFilename, InStrRev(curFilename, ".") - 1)
    newFilename = m_Mask
    newFilename = Replace(newFilename, "##path##", curPath)
    newFilename = Replace(newFilename, "##filename##", curFilename)
    newFilename = Replace(newFilename, "##name##", curName)
    newFilename = Left(newFilename, InStrRev(newFilename, ".")) & "gif"
    If InStr(newFilename, ":\") = 0 Then newFilename = Server.MapPath(newFilename)
    m_Object.SaveGIF newFilename, palette, dither, colors
    WriteDebug "Write GIF", "file = " & newFilename
    
    If LCase(m_File) <> LCase(newFilename) And m_Overwrite Then
      Set fso = Server.CreateObject("Scripting.FileSystemObject")
      fso.DeleteFile m_File, True
      WriteDebug "Delete file", "file = " & m_File
      Set fso = Nothing
    End If
    
    m_File = newFilename
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight, keepAspect)
    newWidth = CInt(newWidth)
    newHeight = CInt(newHeight)
    keepAspect = CBool(keepAspect)
    If newWidth > 0 And newHeight > 0 Then
      If m_Object Is Nothing Then
        Component = "auto"
      End If
      If keepAspect Then
        Dim arrSize
        arrSize = CalculateSize(m_Object.Width, m_Object.Height, newWidth, newHeight)
        newWidth = arrSize(0)
        newHeight = arrSize(1)
      End If
      m_Object.Resize newWidth, newHeight
    End If
    WriteDebug "Resize", "width = " & newWidth & "<br>height = " & newHeight
  End Sub
  
  ' Crop at specific location
  ' Position options are
  ' "Top-Left"
  ' "Top-Center"
  ' "Top-Right"
  ' "Center-Left"
  ' "Center-Center"
  ' "Center-Right"
  ' "Bottom-Left"
  ' "Bottom-Center"
  ' "Bottom-Right"
  Public Sub CropPos(position, width, height)
    Dim posArr, x, y
    position = CStr(position)
    width = CLng(width)
    If width > m_Object.Width Then width = m_Object.Width
    height = CLng(height)
    If height > m_Object.Height Then height = m_Object.Height
    posArr = Split(position, "-")
    Select Case LCase(posArr(1))
      Case "left"   x = 0
      Case "center" x = Round((m_Object.Width / 2) - (width / 2), 0)
      Case "right"  x = m_Object.Width - width
    End Select
    Select Case LCase(posArr(0))
      Case "top"    y = 0
      Case "center" y = Round((m_Object.Height / 2) - (height / 2), 0)
      Case "bottom" y = m_Object.Height - height
    End Select
    Crop x, y, x + width, y + width
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    x1 = CLng(x1)
    y1 = CLng(y1)
    x2 = CLng(x2)
    y2 = CLng(y2)
    If x2 > x1 And y2 > y1 Then
      If m_Object Is Nothing Then
        Component = "auto"
      End If
      m_Object.Crop x1, y1, x2, y2
    End If
    WriteDebug "Crop", "x1 = " & x1 & "<br>y1 = " & y1 & "<br>x2 = " & x2 & "<br>y2 = " & y2
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.RotateLeft
    WriteDebug "Rotate left", ""
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.RotateRight
    WriteDebug "Rotate right", ""
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.Sharpen
    WriteDebug "Sharpen", ""
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.Blur
    WriteDebug "Blur", ""
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.GrayScale
    WriteDebug "GrayScale", ""
  End Sub
  
  ' Add text to an image
  ' Position options are
  ' "Top-Left"
  ' "Top-Center"
  ' "Top-Right"
  ' "Center-Left"
  ' "Center-Center"
  ' "Center-Right"
  ' "Bottom-Left"
  ' "Bottom-Center"
  ' "Bottom-Right"
  Public Sub AddText(str, position)
    Dim regEx, pos
    Set regEx = New RegExp
    regEx.Pattern = "(Top|Center|Bottom)-(Left|Center|Right)"
    regEx.IgnoreCase = True
    If Not regEx.Test(position) Then position = "Top-Left"
    str = CStr(str)
    position = CStr(position)
    pos = GetPosition(position)
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.AddText CStr(str), CStr(position), pos(0), pos(1)
    WriteDebug "Adding Text", "text = " & str & "<br>position = " & position
  End Sub
  
  ' Position options are
  ' "Top-Left"
  ' "Top-Center"
  ' "Top-Right"
  ' "Center-Left"
  ' "Center-Center"
  ' "Center-Right"
  ' "Bottom-Left"
  ' "Bottom-Center"
  ' "Bottom-Right"
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim regEx, pos
    Set regEx = New RegExp
    regEx.Pattern = "(Top|Center|Bottom)-(Left|Center|Right)"
    regEx.IgnoreCase = True
    If Not regEx.Test(position) Then position = "Top-Left"
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    If InStr(file, ":\") = 0 Then file = Server.MapPath(file)
    m_Object.AddWatermark file, position, shrinkToFit, transColor, opacity
    WriteDebug "Adding watermark", "file = " & file & "<br>position = " & position & "<br>shrinkToFit = " & shrinkToFit & "<br>transColor = " & transColor & "<br>opacity = " & opacity
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    If InStr(file, ":\") = 0 Then file = Server.MapPath(file)
    m_Object.AddTiledWatermark file, hspace, vspace, transColor, opacity
    WriteDebug "Adding tiled watermark", "file = " & file & "<br>hspace = " & hspace & "<br>vspace = " & vspace & "<br>transColor = " & transColor & "<br>opacity = " & opacity
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    If InStr(file, ":\") = 0 Then file = Server.MapPath(file)
    m_Object.AddStretchedWatermark file, transColor, opacity
    WriteDebug "Adding stretched watermark", "file = " & file & "<br>transColor = " & transColor & "<br>opacity = " & opacity
  End Sub
  
  ' Flip an image horizontal
  Public Sub FlipHorizontal()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.FlipHorizontal
    WriteDebug "Flip horizontal", ""
  End Sub
  
  ' Flip an image vertical
  Public Sub FlipVertical()
    If m_Object Is Nothing Then
      Component = "auto"
    End If
    m_Object.FlipVertical
    WriteDebug "Flip vertical", ""
  End Sub
  
  ' Detect the component
  Private Function DetectComponent()
    Dim component
    
    DetectComponent = True
    
    For Each component In m_Components
      If TestComponent(component) Then Exit Function
    Next
    
    DetectComponent = False
  End Function
  
  ' Test if the component is installed
  Private Function TestComponent(str)
    On Error Resume Next
    Execute("Set m_Object = New " & str)
    On Error GoTo 0
    
    If m_Object Is Nothing Then
      TestComponent = False
      Set m_Object = Nothing
      Exit Function
    End If
    
    If Not m_Object.CreateComponent(Me) Then
      TestComponent = False
      Set m_Object = Nothing
      Exit Function
    End If
    
    TestComponent = True
  End Function
  
  ' Calculate the new size while keeping aspect ratio
  Private Function CalculateSize(orgWidth, orgHeight, maxWidth, maxHeight)
    Dim newWidth, newHeight
    
    If maxWidth < orgWidth Or maxHeight < orgHeight Then
      If maxWidth >= maxHeight Then
        newWidth = maxHeight * (orgWidth / orgHeight)
        newHeight = maxHeight
      Else
        newWidth = maxWidth
        newHeight = maxWidth * (orgHeight / orgWidth)
      End If
      If newWidth > maxWidth Then
        newWidth = maxWidth
        newHeight = maxWidth * (orgHeight / orgWidth)
      End If
      If newHeight > maxHeight Then
        newWidth = maxHeight * (orgWidth / orgHeight)
        newHeight = maxHeight
      End If
    Else
      newWidth = orgWidth
      newHeight = orgHeight
    End If
    
    CalculateSize = Array(Round(newWidth, 0), Round(newHeight, 0))
  End Function
  
  Public Function GetPosition(position)
    Dim posArr, x, y
    posArr = Split(position, "-")
    Select Case LCase(posArr(1))
      Case "left"   x = 5
      Case "center" x = Round(m_Object.Width / 2, 0)
      Case "right"  x = m_Object.Width - 5
    End Select
    Select Case LCase(posArr(0))
      Case "top"    y = 5
      Case "center" y = Round(m_Object.Height / 2, 0)
      Case "bottom" y = m_Object.Height - 5
    End Select
    GetPosition = Array(x, y)
  End Function
  
  Public Sub DebugStart()
    Response.Write "<style type=""text/css"">" & _
                   ".debugTable {" & _
                   "  background-color: #FFFFFF;" & _
                   "  width: 100%;" & _
                   "  border: 1px solid #000000;" & _
                   "  margin-bottom: 10px;" & _
                   "}" & _
                   ".debugTitle {" & _
                   "  background-color: #009900;" & _
                   "  color: #FFFFFF;" & _
                   "  border-bottom: 1px solid #000000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 12px;" & _
                   "  padding: 3px;" & _
                   "}" & _
                   ".debugCell {" & _
                   "  border-bottom: 1px solid #CCCCCC;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 11px;" & _
                   "  padding: 2px;" & _
                   "}" & _
                   "</style>"
    Response.Write "<center>"
    Response.Write "<table class=""debugTable"" cellpadding=""0"" cellspacing=""0"">"
    Response.Write "<tr><th class=""debugTitle"" colspan=""2"">Image Processor " & IP_Version & " debug</th></tr>"
  End Sub
  
  ' Write debug information to the screen
  Public Sub WriteDebug(cat, str)
    If (m_Debug) Then
      If str = "" Then str = "&nbsp;"
      Response.Write "<tr><td class=""debugCell""><b>" & cat & "</b></td><td class=""debugCell"">" & str & "</td></tr>"
    End If
  End Sub
  
  Public Sub DebugEnd()
    Response.Write "</table>"
    Response.Write "</center>"
  End Sub
  
  Public Sub WriteError(title, str)
    Response.Write "<style type=""text/css"">" & _
                   ".errorTable {" & _
                   "  background-color: #FFFFFF;" & _
                   "  width: 600px;" & _
                   "  border: 1px solid #000000;" & _
                   "  margin-bottom: 10px;" & _
                   "}" & _
                   ".errorTitle {" & _
                   "  background-color: #CC0000;" & _
                   "  color: #FFFFFF;" & _
                   "  border-bottom: 1px solid #000000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 12px;" & _
                   "  padding: 3px;" & _
                   "}" & _
                   ".errorCell {" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 12px;" & _
                   "  padding: 5px;" & _
                   "}" & _
                   "</style>"
    Response.Write "<center>"
    Response.Write "<table class=""errorTable"" cellpadding=""0"" cellspacing=""0"">"
    Response.Write "<tr><th class=""errorTitle"">Image Processor: " & title & "</th></tr>"
    Response.Write "<tr><td class=""errorCell"">" & str & "</td></tr>"
    Response.Write "</table>"
    Response.Write "</center>"
    Response.End
  End Sub
  
  Public Sub ShowInfo()
    Dim component, ver

    Response.Write "<style type=""text/css"">" & _
                   ".infoTable {" & _
                   "  background-color: #FFFFFF;" & _
                   "  width: 500px;" & _
                   "  border-bottom: 1px solid #000000;" & _
                   "  border-right: 1px solid #000000;" & _
                   "  margin-bottom: 10px;" & _
                   "}" & _
                   ".infoTitle {" & _
                   "  background-color: #0066CC;" & _
                   "  color: #FFFFFF;" & _
                   "  border-top: 1px solid #000000;" & _
                   "  border-left: 1px solid #000000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 12px;" & _
                   "  padding: 3px;" & _
                   "}" & _
                   ".infoHeader {" & _
                   "  background-color: #DDDDDD;" & _
                   "  border-top: 1px solid #000000;" & _
                   "  border-left: 1px solid #000000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 11px;" & _
                   "  padding: 2px;" & _
                   "}" & _
                   ".infoCell {" & _
                   "  border-top: 1px solid #000000;" & _
                   "  border-left: 1px solid #000000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 11px;" & _
                   "  padding: 2px;" & _
                   "}" & _
                   ".infoTrue {" & _
                   "  border-top: 1px solid #000000;" & _
                   "  border-left: 1px solid #000000;" & _
                   "  background-color: #00CC00;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 11px;" & _
                   "}" & _
                   ".infoFalse {" & _
                   "  border-top: 1px solid #000000;" & _
                   "  border-left: 1px solid #000000;" & _
                   "  background-color: #CC0000;" & _
                   "  font-family: Verdana;" & _
                   "  font-size: 11px;" & _
                   "}" & _
                   "</style>"

    ShowServerInfo
    ShowResponseObject
    ShowServerObject
    ShowSessionObject
    ShowComponentsInfo
  End Sub
  
  Private Sub TableStart(title, cols)
    Response.Write "<center>"
    Response.Write "<table class=""infoTable"" cellpadding=""0"" cellspacing=""0"">"
    Response.Write "<tr><th class=""infoTitle"" colspan=""" & cols & """>" & title & "</th></tr>"
  End Sub
  
  Private Sub TableEnd()
    Response.Write "</table>"
    Response.Write "</center>"
  End Sub
  
  Private Sub TableHeader(arr)
    Dim str
    Response.Write "<tr>"
    For Each str In arr
      If IsNull(str) Or str = "" Then str = "&nbsp;"
      Response.Write "<th class=""infoHeader"">" & str & "</th>"
    Next
    Response.Write "</tr>"
  End Sub
  
  Private Sub TableRowStart()
    Response.Write "<tr>"
  End Sub
  
  Private Sub TableRowEnd()
    Response.Write "<tr>"
  End Sub
  
  Private Sub TableRow(arr)
    Dim str
    Response.Write "<tr>"
    For Each str In arr
      If IsNull(str) Or str = "" Then str = "&nbsp;"
      Response.Write "<td class=""infoCell"">" & str & "</td>"
    Next
    Response.Write "</tr>"
  End Sub
  
  Private Sub TableCell(arr)
    Dim str
    For Each str In arr
      If IsNull(str) Or str = "" Then str = "&nbsp;"
      Response.Write "<td class=""infoCell"">" & str & "</td>"
    Next
  End Sub
  
  Private Sub TableCellTrue()
    Response.Write "<td class=""infoTrue"">&nbsp;</td>"
  End Sub
  
  Private Sub TableCellFalse()
    Response.Write "<td class=""infoFalse"">&nbsp;</td>"
  End Sub
  
  Private Sub ShowServerInfo
    TableStart "Server Information", 2
    TableHeader Array("Info", "Value")
    TableRow Array("Server&nbsp;Name", Request.ServerVariables("SERVER_NAME"))
    TableRow Array("Server&nbsp;IP&nbsp;Address", Request.ServerVariables("LOCAL_ADDR"))
    TableRow Array("Server&nbsp;Port", Request.ServerVariables("SERVER_PORT"))
    TableRow Array("Server&nbsp;Software", Request.ServerVariables("SERVER_SOFTWARE"))
    TableRow Array("Operating&nbsp;System", Request.ServerVariables("OS"))
    TableRow Array("Script&nbsp;Engine", ScriptEngine & " (version: " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion & ")")
    TableEnd
  End Sub
  
  Private Sub ShowResponseObject
    TableStart "Response Object", 2
    TableHeader Array("Attribute", "Value")
    TableRow Array("Response.Buffer", Response.Buffer)
    TableRow Array("Response.CacheControl", Response.CacheControl)
    TableRow Array("Response.Charset", Response.Charset)
    TableRow Array("Response.ContentType", Response.ContentType)
    TableRow Array("Response.Expires", Response.Expires)
    TableRow Array("Response.ExpiresAbsolute", Response.ExpiresAbsolute)
    TableRow Array("Response.isClientConnected", Response.isClientConnected)
    TableRow Array("Response.Status", Response.Status)
    TableEnd
  End Sub
  
  Private Sub ShowServerObject
    TableStart "Server Object", 2
    TableHeader Array("Attribute", "Value")
    TableRow Array("Server.ScriptTimeout", Server.ScriptTimeout)
    TableEnd
  End Sub
  
  Private Sub ShowSessionObject
    TableStart "Session Object", 2
    TableHeader Array("Attribute", "Value")
    TableRow Array("Session.CodePage", Session.CodePage)
    TableRow Array("Session.LCID", Session.LCID)
    TableRow Array("Session.SessionID", Session.SessionID)
    TableRow Array("Session.TimeOut", Session.TimeOut)
    TableEnd
  End Sub
  
  Private Sub ShowComponentsInfo
    Dim component, ver, funcList, func

    On Error Resume Next
    
    TableStart "Installed Components", 15
    TableHeader Array("Component","Version","Resize","Crop","RotateLeft","RotateRight","Sharpen","Blur","GrayScale","AddText","AddWatermark","AddTiledWatermark","AddStretchedWatermark","FlipHorizontal","FlipVertical")
    
    funcList = Array("Resize","Crop","RotateLeft","RotateRight","Sharpen","Blur","GrayScale","AddText","AddWatermark","AddTiledWatermark","AddStretchedWatermark","FlipHorizontal","FlipVertical")
    
    For Each component In m_Components
      Execute("Set m_Object = New " & component)
      ver = "&nbsp;"
      
      If Not m_Object Is Nothing Then
        If m_Object.CreateComponent(Me) Then
          ver = m_Object.Version
          TableRowStart
          TableCell Array(component, ver)
          For Each func In funcList
            If m_Object.QuerySupport(func) Then
              TableCellTrue
            Else
              TableCellFalse
            End If
          Next
          TableRowEnd
        End If
      End If
      
      Set m_Object = Nothing
    Next
    
    TableEnd
    
    On Error GoTo 0
  End Sub
  
End Class


'------------------------------------------
' Make the class accessable by JScript
'------------------------------------------
Function CreateImageProcessor()
  Set CreateImageProcessor = New ImageProcessor
End Function


Class PictureProcessor
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = False
      Case "blur": QuerySupport = False
      Case "grayscale": QuerySupport = False
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = False
      Case "addtiledwatermark": QuerySupport = False
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = False
      Case "flipvertical": QuerySupport = False
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    If m_Object Is Nothing Then
      Width = 0
    Else
      Width = m_Object.Width
    End If
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    If m_Object Is Nothing Then
      Height = 0
    Else
      Height = m_Object.Height
    End If
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "PictureProcessor"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("COMobjects.NET.PictureProcessor")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
     On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.LoadFromFile file
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.Quality = quality
    m_Object.SaveToFileAsJpeg file
  End Sub
  
  Public Sub SavePNG(file)
    ' Not supported
    Main.WriteError "Error", "SavePNG is not supported by " & ToString
  End Sub
  
  Public Sub SaveGIF(file, palette, dither, colors)
    ' Not supported
    Main.WriteError "Error", "SaveGIF is not supported by " & ToString
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.Resize newWidth, newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    m_Object.Crop x1, y1, x2, y2
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.Rotate (PI / 2)
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.Rotate -(PI / 2)
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    ' Not supported
    Main.WriteError "Error", "Sharpen is not supported by " & ToString
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    ' Not supported
    Main.WriteError "Error", "Blur is not supported by " & ToString
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    ' Not supported
    Main.WriteError "Error", "GrayScale is not supported by " & ToString
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim posArr
    m_Object.FontName = Main.FontFamily
    m_Object.FontHeight = Main.FontSize
    m_Object.FontColor = Main.VbColor
    If Main.Bold Then
      m_Object.FontWeight = 700
    Else
      m_Object.FontWeight = 400
    End If
    m_Object.FontItalicOn = Main.Italic
    m_Object.FontUnderlineOn = Main.Underline
    posArr = Split(position, "-")
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize
    End Select
    m_Object.TextOut str, x, y, 255
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddWatermark is not supported by " & ToString
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddTiledWatermark is not supported by " & ToString
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddStretchedWatermark is not supported by " & ToString
  End Sub
  
  Public Sub FlipHorizontal()
    ' Not supported
    Main.WriteError "Error", "FlipHorizontal is not supported by " & ToString
  End Sub
  
  Public Sub FlipVertical()
    ' Not supported
    Main.WriteError "Error", "FlipVertical is not supported by " & ToString
  End Sub
  
End Class


Class ShotGraph
  
  Public Main
  
  Private m_Object
  Private m_Width
  Private m_Height
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = True
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Width
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
    m_Width = 0
    m_Height = 0
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "ShotGraph"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("shotgraph.image")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
     On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    Dim pal
    m_Object.GetFileDimensions file, m_Width, m_Height
    m_Object.CreateImage m_Width, m_Height, 8
    m_Object.ReadImage file, pal, 0, 0
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.JpegImage quality, 0, file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    m_Object.PngImage -1, 0, file
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    m_Object.GifImage -1, 0, file
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.InitClipboard newWidth, newHeight
    m_Object.Resize 0, 0, newWidth, newHeight, 0, 0, m_Width, m_Height, 3
    m_Object.CreateImage newWidth, newHeight, 8
    m_Object.SelectClipboard True
    m_Object.Copy 0, 0, newWidth, newHeight, 0, 0, "SRCCOPY"
    m_Object.SelectClipboard False
  m_Width = newWidth
  m_Height = newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    Dim newWidth, newHeight
    newWidth = x2 - x1
    newHeight = y2 - y1
    m_Object.InitClipboard newWidth, newHeight
    m_Object.Copy 0, 0, newWidth, newHeight, x1, y1, "SRCCOPY"
    m_Object.CreateImage newWidth, newHeight, 8
    m_Object.SelectClipboard True
    m_Object.Copy 0, 0, newWidth, newHeight, 0, 0, "SRCCOPY"
    m_Object.SelectClipboard False
  m_Width = newWidth
  m_Height = newHeight
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    Rotate 90
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    Rotate -90
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    m_Object.Sharpen
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    m_Object.Blur
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    m_Object.GrayScale
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim clr, pos
    m_Object.CreateFont Main.FontFamily, 1, Main.FontSize, 0, Main.Bold, Main.Italic, Main.Underline, False
    clr = Main.RgbColor
    m_Object.SetColor 0, clr(0), clr(1), clr(2)
    m_Object.SetTextColor 0
    m_Object.FontSmoothing = 2
    Select Case LCase(position)
      Case "top-left"      m_Object.SetTextAlign "TA_LEFT",   "TA_TOP"
      Case "top-center"    m_Object.SetTextAlign "TA_CENTER", "TA_TOP"
      Case "top-right"     m_Object.SetTextAlign "TA_RIGHT",  "TA_TOP"
      Case "center-left"   m_Object.SetTextAlign "TA_LEFT",   "TA_TOP"
      Case "center-center" m_Object.SetTextAlign "TA_CENTER", "TA_TOP"
      Case "center-right"  m_Object.SetTextAlign "TA_RIGHT",  "TA_TOP"
      Case "bottom-left"   m_Object.SetTextAlign "TA_LEFT",   "TA_BOTTOM"
      Case "bottom-center" m_Object.SetTextAlign "TA_CENTER", "TA_BOTTOM"
      Case "bottom-right"  m_Object.SetTextAlign "TA_RIGHT",  "TA_BOTTOM"
    End Select
    pos = Split(position, "-")
    If LCase(pos(0)) = "center" Then y = Round(y - (Main.FontSize / 2), 0)
    m_Object.TextOut x, y, str
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim wmWidth, wmHeight, pal, x, y
    m_Object.GetFileDimensions file, wmWidth, wmHeight
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((m_Width - wmWidth) / 2, 0)
      Case "right"  x = m_Width - wmWidth
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((m_Height - wmHeight) / 2, 0)
      Case "bottom" y = m_Height - wmHeight
    End Select
    m_Object.InitClipboard wmWidth, wmHeight
    m_Object.SelectClipboard True
    m_Object.ReadImage file, pal, 0, 0
    m_Object.Copy x, y, wmWidth, wmHeight, 0, 0, "SRCCOPY"
    m_Object.SelectClipboard False
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim wmWidth, wmHeight, pal, x, y
    m_Object.GetFileDimensions file, wmWidth, wmHeight
    For y = vspace To m_Height Step wmWidth + vspace
      For x = hspace To m_Width Step wmHeight + hspace
        m_Object.ReadImage file, pal, x, y
      Next
    Next
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    Dim wmWidth, wmHeight, pal
    m_Object.GetFileDimensions file, wmWidth, wmHeight
    m_Object.InitClipboard wmWidth, wmHeight
    m_Object.SelectClipboard True
    m_Object.ReadImage file, pal, 0, 0
    m_Object.Resize 0, 0, m_Width, m_Height, 0, 0, wmWidth, wmHeight, 3
    m_Object.SelectClipboard False
  End Sub
  
  Public Sub FlipHorizontal()
    m_Object.Flip True
  End Sub
  
  Public Sub FlipVertical()
    m_Object.Flip False
  End Sub
  
  ' Rotate an image
  Private Sub Rotate(degrees)
    Dim newWidth, newHeight
    newWidth = m_Width
    newHeight = m_Height
    m_Object.Rotate degrees, True, newWidth, newHeight
    m_Object.InitClipboard newWidth, newHeight
    m_Object.Rotate degrees, False, m_Width, m_Height
    m_Object.CreateImage m_Width, m_Height, 8
    m_Object.SelectClipboard True
    m_Object.Copy 0, 0, m_Width, m_Height, 0, 0, "SRCCOPY"
    m_Object.SelectClipboard False
  End Sub
  
End Class


Class AspJpeg
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = False
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = True
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.Width
  End Property
  
  ' Returns the height o fthe loaded image
  Public Property Get Height()
    Height = m_Object.Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = m_Object.Version
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("Persits.Jpeg")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If

    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.Open file
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.Quality = quality
    m_Object.Save file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    ' Not Supported
    Main.WriteError "Error", "SavePNG is not supported by " & ToString
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    ' Not Supported
    Main.WriteError "Error", "SaveGIF is not supported by " & ToString
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.Width = newWidth
    m_Object.Height = newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    m_Object.Crop x1, y1, x2, y2
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.RotateL
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.RotateR
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    m_Object.Sharpen 1, 120
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    ' Not supported
    Main.WriteError "Error", "Blur is not supported by " & ToString
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    m_Object.Grayscale 1
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim posArr
    m_Object.Canvas.Font.Family = Main.FontFamily
    m_Object.Canvas.Font.Size = Main.FontSize
    m_Object.Canvas.Font.Color = Eval("&H" & Mid(Main.FontColor, 2))
    m_Object.Canvas.Font.Bold = Main.Bold
    m_Object.Canvas.Font.Italic = Main.Italic
    m_Object.Canvas.Font.Underlined = Main.Underline
    posArr = Split(position, "-")
    Select Case LCase(posArr(1))
      Case "center" x = Round(x - (m_Object.Canvas.GetTextExtent(str) / 2), 0)
      Case "right"  x = x - m_Object.Canvas.GetTextExtent(str)
    End Select
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize
    End Select
    m_Object.Canvas.Print x, y, str
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim img, pos
    Set img = Server.CreateObject("Persits.Jpeg")
    img.Open file
    If shrinkToFit Then
      img.PreserveAspectRatio = True
      If img.OriginalWidth > Width Then
        img.Width = Width
      End If
      If img.OriginalHeight > Height Then
        img.Height = Height
      End If
    End If
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((Width - img.Width) / 2, 0)
      Case "right"  x = Width - img.Width
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((Height - img.Height) / 2, 0)
      Case "bottom" y = Height - img.Height
    End Select
    transColor = Eval("&H" & Mid(transColor, 2))
    If img.TransparencyColorExists Then transColor = img.TransparencyColor
    m_Object.DrawImage x, y, img, opacity / 100, transColor
    Set img = Nothing
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("Persits.Jpeg")
    img.Open file
    transColor = Eval("&H" & Mid(transColor, 2))
    If img.TransparencyColorExists Then transColor = img.TransparencyColor
    For y = vspace To Height Step img.Height + vspace
      For x = hspace To Width Step img.Width + hspace
        m_Object.DrawImage x, y, img, opacity / 100, transColor
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("Persits.Jpeg")
    img.Open file
    img.PreserveAspectRatio = False
    img.Width = Width
    img.Height = Height
    transColor = Eval("&H" & Mid(transColor, 2))
    If img.TransparencyColorExists Then transColor = img.TransparencyColor
    m_Object.DrawImage x, y, img, opacity / 100, transColor
    Set img = Nothing
  End Sub
  
  Public Sub FlipHorizontal()
    m_Object.FlipH
  End Sub
  
  Public Sub FlipVertical()
    m_Object.FlipV
  End Sub
  
End Class


Class AspImage
  
  Public Main
  
  Private m_Object
  Private m_Version
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.MaxX
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.MaxY
  End Property
  
  Public Property Get Version()
    Version = m_Object.Version
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "AspImage"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("AspImage.Image")

    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      m_Version = Split(m_Object.Version, ".")
      CreateComponent = True
    End If
    
    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.LoadImage(file)
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.ImageFormat = 1
    m_Object.JPEGQuality = quality
    m_Object.FileName = file
    m_Object.SaveImage
  End Sub
  
  Public Sub SavePNG(file)
    m_Object.ImageFormat = 3
    m_Object.FileName = file
    m_Object.SaveImage
  End Sub
  
  Public Sub SaveGIF(file, palette, dither, colors)
    m_Object.ImageFormat = 5
    If colors > 16 Then
      m_Object.PixelFormat = 3
    Else
      m_Object.PixelFormat = 2
    End If
    m_Object.FileName = file
    m_Object.SaveImage
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    If CInt(m_Version(0)) >= 2 Then
      m_Object.ResizeR newWidth, newHeight
    Else
      m_Object.Resize newWidth, newHeight
    End If
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    Dim newWidth, newHeight
    newWidth = x2 - x1
    newHeight = y2 - y1
    m_Object.CropImage x1, y1, newWidth, newHeight
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.RotateImage 90
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.RotateImage -90
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    If CInt(m_Version(0)) >= 2 Then
      m_Object.Sharpen 1
    End If
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    If CInt(m_Version(0)) >= 2 Then
      m_Object.Blur 1
    End If
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    If CInt(m_Version(0)) >= 2 Then
      m_Object.CreateGrayScale
    End If
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim posArr
    m_Object.AntiAliasText = True
    m_Object.FontName = Main.FontFamily
    m_Object.FontSize = Main.FontSize
    m_Object.FontColor = Main.VbColor
    m_Object.Bold = Main.Bold
    m_Object.Italic = Main.Italic
    m_Object.Underline = Main.Underline
    posArr = Split(position, "-")
    Select Case LCase(posArr(1))
      Case "center" x = Round(x - (m_Object.TextWidth(str) / 2), 0)
      Case "right"  x = x - m_Object.TextWidth(str)
    End Select
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (m_Object.TextHeight(str) / 2), 0)
      Case "bottom"  y = y - m_Object.TextHeight(str)
    End Select
    m_Object.TextOut str, x, y, False
  End Sub
  
  Private Function GetVbColor(color)
    Dim R, G, B
    R = Eval("&H" & Mid(color, 2, 2))
    G = Eval("&H" & Mid(color, 4, 2))
    B = Eval("&H" & Mid(color, 6, 2))
    GetVbColor = RGB(R, G, B)
  End Function
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim img, wmWidth, wmHeight, pos
    Set img = Server.CreateObject("AspImage.Image")
    img.LoadImage(file)
    wmWidth = img.MaxX
    wmHeight = img.MaxY
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((Width - wmWidth) / 2, 0)
      Case "right"  x = Width - wmWidth
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((Height - wmHeight) / 2, 0)
      Case "bottom" y = Height - wmHeight
    End Select
    m_Object.AddImageTransparent file, x, y, GetVbColor(transColor)
    Set img = Nothing
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img, wmWidth, wmHeight, pos
    Set img = Server.CreateObject("AspImage.Image")
    img.LoadImage(file)
    wmWidth = img.MaxX
    wmHeight = img.MaxY
    For y = vspace To Height Step vspace + wmHeight
      For x = hspace To Width Step hspace + wmWidth
        m_Object.AddImageTransparent file, x, y, GetVbColor(transColor)
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    ' Not supported
  End Sub
  
  Public Sub FlipHorizontal()
    m_Object.FlipImage 1
  End Sub
  
  Public Sub FlipVertical()
    m_Object.FlipImage 2
  End Sub
  
End Class


Class SmartImage
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.Width
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "SmartImage"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("aspSmartImage.SmartImage")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.OpenFile CStr(file)
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.Quality = quality
    m_Object.SaveFile file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    ' Not Supported
    Main.WriteError "Error", "SavePNG is not supported by " & ToString
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    ' Not Supported
    Main.WriteError "Error", "SaveGIF is not supported by " & ToString
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.Resample CLng(newWidth), CLng(newHeight)
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    m_Object.Crop CLng(x1), CLng(y1), CLng(x2), CLng(y2)
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.RotateLeft
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.RotateRight
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    m_Object.Sharpen
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    m_Object.Blur
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    m_Object.GreyScale
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim posArr
    m_Object.FontFace = Main.FontFamily
    m_Object.FontSize = Main.FontSize
    m_Object.FontColor = Mid(Main.FontColor, 2)
    m_Object.FontBold = Main.Bold
    m_Object.FontItalic = Main.Italic
    m_Object.FontUnderline = Main.Underline
    posArr = Split(position, "-")
    m_Object.AddText CStr(str), UCase(posArr(1)), UCase(posArr(0))
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim pos
    pos = Split(position, "-")
    m_Object.AddImage CStr(file), UCase(pos(1)), UCase(pos(0)), Right(transColor, 6)
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("aspSmartImage.SmartImage")
    img.OpenFile CStr(file)
    For y = vspace To Height Step vspace + img.Height
      For x = hspace To Width Step hspace + img.Width
        m_Object.AddImage CStr(file), x, y, Right(transColor, 6)
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddStretchedWatermark is not supported by " & ToString
  End Sub
  
  Public Sub FlipHorizontal()
    m_Object.Mirror
  End Sub
  
  Public Sub FlipVertical()
    m_Object.Flip
  End Sub
  
End Class


Class ImgWriter
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = False
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.Width
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "ImgWriter"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("softartisans.ImageGen")
    
    On Error GoTo 0

    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.LoadImage file
  End Sub
  
  ' Save image to file
  Public Sub Save(file, quality)
    m_Object.ImageQuality = quality
    m_Object.SaveImage 0, 3, file
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.ResizeFilter = 13
    m_Object.ResizeImage newWidth, newHeight, 3
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    Dim newWidth, newHeight
    newWidth = x2 - x1
    newHeight = y2 - y1
    m_Object.CropImage x1, y1, newWidth, newHeight
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.RotateImage -90, RGB(0, 0, 0)
  End Sub

  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.RotateImage 90, RGB(0, 0, 0)
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    m_Object.SharpenImage 35
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    m_Object.BlurImage 35
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    ' Not supported
    Main.WriteError "Error", "GrayScale is not supported by " & ToString
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim width, height, posArr
    m_Object.Font.name = Main.FontFamily
    m_Object.Font.Height = Main.FontSize
    m_Object.Font.Color = Main.VbColor
    If Main.Bold Then
      m_Object.Font.Weight = 700
    Else
      m_Object.Font.Weight = 0
    End If
    m_Object.Font.Italic = Main.Italic
    m_Object.Font.Underline = Main.Underline
    width = m_Object.TextWidth(str)
    height = m_Object.TextHeight(str)
    posArr = Split(position)
    Select Case LCase(posArr(1))
      Case "center" x = Round(x - (width / 2), 0)
      Case "right"  x = x - width
    End Select
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (height / 2), 0)
      Case "bottom" y = y - height
    End Select
    m_Object.DrawTextOnImage x, y, width, height, str
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Select Case LCase(position)
      Case "top-center"    pos = 0
      Case "center-center" pos = 1
      Case "bottom-center" pos = 2
      Case "top-left"      pos = 3
      Case "center-left"   pos = 4
      Case "bottom-left"   pos = 5
      Case "top-right"     pos = 6
      Case "center-right"  pos = 7
      Case "bottom-right"  pos = 8
    End Select
    m_Object.AddWatermark file, pos, opacity / 100, Eval("&H" & Right(transColor, 6)), shrinkToFit
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("softartisans.ImageGen")
    img.LoadImage file
    For y = vspace To Height Step vspace + img.Height
      For x = hspace To Width Step hspace + img.Width
        m_Object.MergeWithImage x, y, file, opacity / 100, Eval("&H" & Right(transColor, 6))
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddStretchedWatermark is not supported by " & ToString
  End Sub
  
  Public Sub FlipHorizontal()
    m_Object.FlipImage 0
  End Sub
  
  Public Sub FlipVertical()
    m_Object.FlipImage 1
  End Sub
  
End Class


Class AspThumb
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = False
      Case "rotateleft": QuerySupport = False
      Case "rotateright": QuerySupport = False
      Case "sharpen": QuerySupport = False
      Case "blur": QuerySupport = False
      Case "grayscale": QuerySupport = False
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = False
      Case "addtiledwatermark": QuerySupport = False
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = False
      Case "flipvertical": QuerySupport = False
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.Width
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "AspThumb"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("briz.AspThumb")

    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.Load file
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.EncodingQuality = quality
    m_Object.Save file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    ' Not Supported
    Main.WriteError "Error", "SavePNG is not supported by " & ToString
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    ' Not Supported
    Main.WriteError "Error", "SaveGIF is not supported by " & ToString
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    ' Always keeps aspect ratio
    m_Object.ResizeQuality = 2
    m_Object.Resize newWidth, newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    ' Not Supported
    Main.WriteError "Error", "Crop is not supported by " & ToString
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    ' Not supported
    Main.WriteError "Error", "RotateLeft is not supported by " & ToString
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    ' Not supported
    Main.WriteError "Error", "RotateRight is not supported by " & ToString
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    ' Not supported
    Main.WriteError "Error", "Sharpen is not supported by " & ToString
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    ' Not supported
    Main.WriteError "Error", "Blur is not supported by " & ToString
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    ' Not supported
    Main.WriteError "Error", "GrayScale is not supported by " & ToString
  End Sub
  
  Public Sub AddText(str, position, x, y)
    On Error Resume Next
    Dim posArr
    m_Object.SetFont Main.FontFamily, Main.FontSize, Main.Bold, Main.Italic, Main.Underline, Main.VbColor
    posArr = Split(position, "-")
    ' Todo: bepaal de x waarde voor center en Right
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize
    End Select
    m_Object.DrawText x, y, str
    On Error GoTo 0
  End Sub
  
  Public Sub AddWatermark(file, x, y, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddWatermark is not supported by " & ToString
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddTiledWatermark is not supported by " & ToString
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    ' Not supported
    Main.WriteError "Error", "AddStretchedWatermark is not supported by " & ToString
  End Sub
  
  Public Sub FlipHorizontal()
    ' Not supported
    Main.WriteError "Error", "FlipHorizontal is not supported by " & ToString
  End Sub
  
  Public Sub FlipVertical()
    ' Not supported
    Main.WriteError "Error", "FlipVertical is not supported by " & ToString
  End Sub
  
End Class


Class AspNet
  
  Public Main
  
  Private m_Object
  Private m_Url
  
  Private m_Width
  Private m_Height
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = True
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  Public Property Get Width()
    Width = m_Width
  End Property
  
  Public Property Get Height()
    Height = m_Height
  End Property
  
  Public Property Get Version()
    ExecuteAction "Version", Array()
    Version = m_Object.ResponseXML.Text
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  Public Function ToString()
    ToString = "AspNet"
  End Function
  
  Public Function CreateComponent(parent)
    Set Main = parent
    
    m_Url = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
    m_Url = Left(m_Url, InStrRev(m_Url, "/")) & Main.ScriptFolder & "/sip_service.asmx"

    Main.WriteDebug "Url to ASP.NET service", m_Url
    
    CreateComponent = True
    
    If CheckComponent("Msxml2.ServerXMLHTTP.6.0") Then Exit Function
    If CheckComponent("Msxml2.ServerXMLHTTP.5.0") Then Exit Function
    If CheckComponent("Msxml2.ServerXMLHTTP.4.0") Then Exit Function
    If CheckComponent("Msxml2.ServerXMLHTTP.3.0") Then Exit Function
    If CheckComponent("Msxml2.ServerXMLHTTP") Then Exit Function
    If CheckComponent("Msxml2.XMLHTTP.3.0") Then Exit Function
    If CheckComponent("Microsoft.XMLHTTP") Then Exit Function
    
    Main.WriteDebug "Error creating object " & ToString, Err.Description
    
    CreateComponent = False
  End Function
  
  Private Function CheckComponent(component)
    On Error Resume Next
    Err.Clear
    
    Dim arrResponse
    
    Set m_Object = Server.CreateObject(component)
    
    If Err.Number = 0 Then
      m_Object.Open "POST", m_Url & "/Version", False
      
      If Err.Number = 0 Then
        m_Object.Send ""
        
        If Left(m_object.ResponseXML.Text, 14) = "ImageProcessor" Then
          CheckComponent = True
          On Error GoTo 0
          Exit Function
        Else
          Main.WriteDebug "Error", m_object.ResponseText
        End If
      Else
        Main.WriteDebug "Error", m_object.ResponseText
      End If
    Else
      Main.WriteDebug "Error creating object " & component, Err.Description
    End If
    
    CheckComponent = False
    
    On Error GoTo 0
  End Function
  
  Private Sub ExecuteAction(action, parms)
    Dim p, i, debugParms
    
    p = ""
    For i = 0 To UBound(parms) Step 2
      p = p & "&" & parms(i) & "=" & Server.URLEncode(parms(i+1))
    Next
    
    debugParms = ""
    For i = 0 To UBound(parms) Step 2
      debugParms = debugParms & parms(i) & " = " & parms(i+1) & "<br>"
    Next
    If debugParms = "" Then debugParms = "no parameters"
    Main.WriteDebug action, debugParms
                                      
    m_Object.Open "POST", m_Url & "/" & action, False
    m_Object.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    m_Object.Send p
    
    If m_Object.Status <> 200 Then
      Main.WriteError "Error executing the action " & action, m_Object.ResponseText
    End If
  
    Main.WriteDebug action & " Response", m_Object.ResponseXML.Text
  
    If LCase(action) <> "version" And LCase(action) <> "getimagesize" And LCase(action) <> "getsize" And LCase(action) <> "measuretextwidth" And LCase(m_Object.ResponseXML.Text) <> "ok" Then
      Main.WriteError "Error executing the action " & action, m_Object.ResponseXML.Text
    End If
  End Sub
  
  ' Load image from file
  Public Sub Load(file)
    Dim theSize
    
    ExecuteAction "Load", Array("filename", file)
    
    ExecuteAction "GetSize", Array()
    
    theSize = Split(m_Object.ResponseXML.Text, " ")
    m_Width = theSize(0)
    m_Height = theSize(1)
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
  ExecuteAction "SaveJPEG", Array("filename", file, "quality", quality)
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
  ExecuteAction "SavePNG", Array("filename", file)
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
  ExecuteAction "SaveGIF", Array("filename", file, "palette", palette, "dither", dither, "maxColors", colors)
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
  ExecuteAction "Resize", Array("width", newWidth, "height", newHeight, "keepAspect", "false")
  
  m_Width = newWidth
  m_Height = newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    ExecuteAction "Crop", Array("left", x1, "top", y1, "width", x2 - x1, "height", y2 - y1)
    
    m_Width = x2 - x1
    m_Height = y2 - y1
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    Dim tmp
    
    ExecuteAction "Rotate270", Array()
    
    tmp = m_Width
    m_Width = m_Height
    m_Height = tmp
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    Dim tmp
    
    ExecuteAction "Rotate90", Array()
    
    tmp = m_Width
    m_Width = m_Height
    m_Height = tmp
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    ExecuteAction "Sharpen", Array("weight", 15)
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    ExecuteAction "GaussianBlur", Array("weight", 4)
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    ExecuteAction "GrayScale", Array()
  End Sub
  
  Public Sub AddText(str, position, x, y)
    Dim textWidth, pos, clr
    ExecuteAction "MeasureTextWidth", Array("text", str, "fontFamily", Main.FontFamily, "fontSize", Main.FontSize)
    textWidth = M_Object.ResponseXML.Text
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "center" x = Round(x - (textWidth / 2), 0)
      Case "right"  x = x - textWidth
    End Select
    Select Case LCase(pos(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize + 2
    End Select
    x = Round(x, 0)
    y = Round(y, 0)
    clr = Main.RgbColor
    ExecuteAction "AddText", Array("text", str, "x", x, "y", y, "fontFamily", Main.FontFamily, "fontSize", Main.FontSize, "alpha", 255, "red", clr(0), "green", clr(1), "blue", clr(2), "bold", Main.Bold, "italic", Main.Italic, "underline", Main.Underline)
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim wmSize, pos, R, G, B
    ExecuteAction "GetImageSize", Array("filename", file)
    wmSize = Split(m_Object.ResponseXML.Text, " ")
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((m_Width - CInt(wmSize(0))) / 2, 0)
      Case "right"  x = m_Width - CInt(wmSize(0))
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((m_Height - CInt(wmSize(1))) / 2, 0)
      Case "bottom" y = m_Height - CInt(wmSize(1))
    End Select
    R = Eval("&H" & Mid(transColor, 2, 2))
    G = Eval("&H" & Mid(transColor, 4, 2))
    B = Eval("&H" & Mid(transColor, 6, 2))
    ExecuteAction "AddWatermark", Array("filename", file, "x", x, "y", y, "alpha", 255, "red", R, "green", G, "blue", B, "opacity", opacity)
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim R, G, B
    R = Eval("&H" & Mid(transColor, 2, 2))
    G = Eval("&H" & Mid(transColor, 4, 2))
    B = Eval("&H" & Mid(transColor, 6, 2))
    ExecuteAction "AddTiledWatermark", Array("filename", file, "hspace", hspace, "vspace", vspace, "alpha", 255, "red", R, "green", G, "blue", B, "opacity", opacity)
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    Dim R, G, B
    R = Eval("&H" & Mid(transColor, 2, 2))
    G = Eval("&H" & Mid(transColor, 4, 2))
    B = Eval("&H" & Mid(transColor, 6, 2))
    ExecuteAction "AddStretchedWatermark", Array("filename", file, "alpha", 255, "red", R, "green", G, "blue", B, "opacity", opacity)
  End Sub
  
  Public Sub FlipHorizontal()
    ExecuteAction "FlipHorizontal", Array()
  End Sub
  
  Public Sub FlipVertical()
    ExecuteAction "FlipVertical", Array()
  End Sub
  
End Class


Class ImageX
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = True
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = True
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.ImgW
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.ImgH
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "ImageX"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("ImageX.CImageX.1")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object" & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.Load file
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.JpegQuality = quality
    m_Object.Save file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    m_Object.Save file
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    If colors > 16 Then
      m_Object.BitDepth = 8
    ElseIf colors > 4 Then
      m_Object.BitDepth = 4
    ElseIf colors > 2 Then
      m_Object.BitDepth = 2
    Else
      m_Object.BitDepth = 1
    End If
    
    m_Object.Save file
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.Resize newWidth, newHeight
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    Dim width, height
    width = x2 - x1
    height = y2 - y1
    m_Object.Crop x1, y1, width, height
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.RotateLeft
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.RotateRight
  End Sub
  
  ' Sharpen an image
  Public Sub Sharpen()
    m_Object.Sharpen
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    m_Object.Blur
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    m_Object.GrayScale
  End Sub
  
  ' Add text to an image
  Public Sub AddText(str, position, x, y)
    Dim pos
    pos = Split(position, "-")
    Select Case LCase(pos(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize
    End Select
    m_Object.DrawText x, y, str, Main.FontFamily, Main.FontSize, Main.VbColor, -1
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("ImageX.CImageX.1")
    img.Load file
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((Width - img.ImgW) / 2, 0)
      Case "right"  x = Width - img.ImgW
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((Height - img.ImgH) / 2, 0)
      Case "bottom" y = Height - img.ImgH
    End Select
    m_Object.Mix img.GetImageRef, 6, x, y, opacity
    Set img = Nothing
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("ImageX.CImageX.1")
    img.Load file
    For y = vspace To Height Step vspace + img.ImgW
      For x = hspace To Width Step hspace + img.ImgH
        m_Object.Mix img.GetImageRef, 6, x, y, opacity
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("ImageX.CImageX.1")
    img.Load file
    img.Resize Width, Height
     m_Object.Mix img.GetImageRef, 6, 0, 0, opacity
    Set img = Nothing
  End Sub
  
  ' Flip an image horizontal
  Public Sub FlipHorizontal()
    m_Object.Flip
  End Sub
  
  ' Flip an image vertical
  Public Sub FlipVertical()
    m_Object.Mirror
  End Sub
  
End Class


Class GraphicsMill
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = False
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = True
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = m_Object.Data.Width
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = m_Object.Data.Height
  End Property
  
  Public Property Get Version()
    Version = "Unknown"
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "GraphicsMill"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    On Error Resume Next
    Err.Clear
    
    Set Main = parent
    
    Set m_Object = Server.CreateObject("GraphicsMill.Bitmap")
    
    If Err.Number <> 0 Or Not IsObject(m_Object) Then
      Main.WriteDebug "Error creating object " & ToString, Err.Description
      CreateComponent = False
    Else
      CreateComponent = True
    End If
    
    On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
    m_Object.LoadFromFile file
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
    m_Object.Formats.SelectCurrent "JPEG"
    m_Object.Formats.Current.EncoderOptions("JpegQuality").Value = quality
    m_Object.SaveToFile file
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
    m_Object.Formats.SelectCurrent "PNG"
    m_Object.SaveToFile file
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
    Dim d, p
    d = 3
    p = 0
    
    Select Case LCase(dither)
      Case "none"              d = 0
      Case "floydsteinberg"    d = 3
      Case "jarvisjudiceninke" d = 5
      Case "stucki"            d = 6
      Case "sierra3"           d = 7
      Case "sierra2"           d = 7
      Case "burkes"            d = 8
    End Select
    
    Select Case LCase(palette)
      Case "adaptive" p = 0
      Case "websafe"  p = 2
    End Select
    
    m_Object.Data.ConvertTo8bppIndexed d, 100, 0, 2, colors, p
    m_Object.Formats.SelectCurrent "GIF"
    m_Object.IsLzwEnabled = True
    m_Object.SaveToFile file
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
    m_Object.Transforms.Resize newWidth, newHeight, 2
    'm_Object.Transforms.Resize 150, 150, 2
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
    Dim width, height
    width = x2 - x1
    height = y2 - y1
    m_Object.Transforms.Crop x1, y1, width, height
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
    m_Object.Transforms.RotateAndFlip 3
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
    m_Object.Transforms.RotateAndFlip 1
  End Sub
  
  ' Sharpen the image
  Public Sub Sharpen()
    m_Object.Effects.Sharpen 50
  End Sub
  
  ' Blur an image
  Public Sub Blur()
    m_Object.Effects.Blur 2, 0
  End Sub
  
  ' Convert to grayscale
  Public Sub GrayScale()
    ' Not supported
    Main.WriteError "Error", "GrayScale is not supported by " & ToString
  End Sub
  
  ' Add text to an image
  Public Sub AddText(str, position, x, y)
    Dim posArr
    m_Object.Graphics.TextFormat.FontName = Main.FontFamily
    m_Object.Graphics.TextFormat.FontSize = Main.FontSize
    m_Object.Graphics.TextFormat.FontColor = Eval("&HFF" & Mid(Main.FontColor, 2))
    m_Object.Graphics.TextFormat.IsBold = Main.Bold
    m_Object.Graphics.TextFormat.IsItalic = Main.Italic
    m_Object.Graphics.TextFormat.IsUnderlined = Main.Underline
    posArr = Split(position, "-")
    Select Case LCase(posArr(1))
      Case "center" x = Round(x - (m_Object.Graphics.TextFormat.MeasureTextWidth(str) / 2), 0)
      Case "right"  x = x - m_Object.Graphics.TextFormat.MeasureTextWidth(str)
    End Select
    Select Case LCase(posArr(0))
      Case "center" y = Round(y - (Main.FontSize / 2), 0)
      Case "bottom" y = y - Main.FontSize
    End Select
    m_Object.Graphics.DrawText str, x, y
  End Sub
  
  Public Sub AddWatermark(file, position, shrinkToFit, transColor, opacity)
    Dim img, pos
    Set img = Server.CreateObject("GraphicsMill.Bitmap")
    img.LoadFromFile file
    pos = Split(position, "-")
    Select Case LCase(pos(1))
      Case "left"   x = 0
      Case "center" x = Round((Width - img.Data.Width) / 2, 0)
      Case "right"  x = Width - img.Data.Width
    End Select
    Select Case LCase(pos(0))
      Case "top"    y = 0
      Case "center" y = Round((Height - img.Data.Height) / 2, 0)
      Case "bottom" y = Height - img.Data.Height
    End Select
    img.DrawOnBitmap m_Object, x, y, img.Data.Width, img.Data.Height, 0, 0, -1, -1, 0, (opacity / 100) * 255
    Set img = Nothing
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("GraphicsMill.Bitmap")
    img.LoadFromFile file
    For y = vspace To Height Step vspace + img.Data.Height
      For x = hspace To Width Step hspace + img.Data.Width
        img.DrawOnBitmap m_Object, x, y, img.Data.Width, img.Data.Height, 0, 0, -1, -1, 0, (opacity / 100) * 255
      Next
    Next
    Set img = Nothing
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
    Dim img
    Set img = Server.CreateObject("GraphicsMill.Bitmap")
    img.LoadFromFile file
     img.DrawOnBitmap m_Object, 0, 0, Width, Height, 0, 0, -1, -1, 0, (opacity / 100) * 255
    Set img = Nothing
  End Sub
  
  ' Flip an image horizontal
  Public Sub FlipHorizontal()
    m_Object.Transforms.RotateAndFlip 4
  End Sub
  
  ' Flip an image vertical
  Public Sub FlipVertical()
    m_Object.Transforms.RotateAndFlip 6
  End Sub
  
End Class


Class ComponentTemplateClass
  
  Public Main
  
  Private m_Object
  
  Public Function QuerySupport(funcName)
    Select Case LCase(funcName)
      Case "resize": QuerySupport = True
      Case "crop": QuerySupport = True
      Case "rotateleft": QuerySupport = True
      Case "rotateright": QuerySupport = True
      Case "sharpen": QuerySupport = True
      Case "blur": QuerySupport = True
      Case "grayscale": QuerySupport = False
      Case "addtext": QuerySupport = True
      Case "addwatermark": QuerySupport = True
      Case "addtiledwatermark": QuerySupport = True
      Case "addstretchedwatermark": QuerySupport = False
      Case "fliphorizontal": QuerySupport = True
      Case "flipvertical": QuerySupport = True
    End Select
  End Function
  
  Public Property Get Object()
    Set Object = m_Object
  End Property
  
  ' Returns the width of the loaded image
  Public Property Get Width()
    Width = 0
  End Property
  
  ' Returns the height of the loaded image
  Public Property Get Height()
    Height = 0
  End Property
  
  Private Sub Class_Initialize()
  End Sub
  
  Private Sub Class_Terminate()
    Set m_Object = Nothing
  End Sub
  
  ' Returns a string representation of the component used
  Public Function ToString()
    ToString = "NewComponent"
  End Function
  
  ' Try to create the component
  Public Function CreateComponent(parent)
    ' Try creating the object
    'On Error Resume Next
    'Err.Clear
    
    'Set Main = parent
    
    'Set m_Object = Server.CreateObject("NewComponent")
    
    'If Err.Number <> 0 Or Not IsObject(m_Object) Then
      'Main.WriteDebug "Error creating object " & ToString, Err.Description
      'CreateComponent = False
    'Else
      'CreateComponent = True
    'End If
    
    'On Error GoTo 0
  End Function
  
  ' Load image from file
  Public Sub Load(file)
  End Sub
  
  ' Save image to file
  Public Sub SaveJPEG(file, quality)
  End Sub
  
  ' Save image to file
  Public Sub SavePNG(file)
  End Sub
  
  ' Save image to file
  Public Sub SaveGIF(file, palette, dither, colors)
  End Sub
  
  ' Resize an image
  Public Sub Resize(newWidth, newHeight)
  End Sub
  
  ' Crop an image
  Public Sub Crop(x1, y1, x2, y2)
  End Sub
  
  ' Rotate an image counter-clockwise
  Public Sub RotateLeft()
  End Sub
  
  ' Rotate an image clockwise
  Public Sub RotateRight()
  End Sub
  
  ' Sharpen the image
  Public Sub Sharpen()
  End Sub
  
  ' Add text to an image
  Public Sub AddText(str, position, x, y)
  End Sub
  
  Public Sub AddWatermark(file, x, y, transColor, opacity)
  End Sub
  
  Public Sub AddTiledWatermark(file, hspace, vspace, transColor, opacity)
  End Sub
  
  Public Sub AddStretchedWatermark(file, transColor, opacity)
  End Sub
  
  ' Flip an image horizontal
  Public Sub FlipHorizontal()
  End Sub
  
  ' Flip an image vertical
  Public Sub FlipVertical()
  End Sub
  
End Class

</SCRIPT>