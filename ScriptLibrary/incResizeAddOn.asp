<SCRIPT LANGUAGE="VBSCRIPT" RUNAT="SERVER">
'*** Resize Files After Upload -----------------------------------------------
' Copyright 2001-2006 (c) George Petrov, www.DMXzone.com
'
' Version: 1.1.5
'------------------------------------------------------------------------------

sub FitImage_Comp(compType,DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  select case compType
  case "AUTO"
    FitImage_Comp DetectImageComponent(DotNetResize),DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "PICPROC"
    FitImage_PicProc imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "SHOTGRAPH"
    FitImage_ShotGraph imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "ASPJPEG"
    FitImage_AspJpeg imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "ASPIMAGE"
    FitImage_AspImage imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "ASPSMART"
    FitImage_AspSmart imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "IMGWRITER"
    FitImage_ImgWriter imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "ASPTHUMB"
    FitImage_AspThumb imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
  case "ASP.NET"
    select case DetectDotNetComponent(DotNetResize)
    case "DOTNET1"
      FitImage_DotNet "Msxml2.ServerXMLHTTP.3.0",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
    case "DOTNET2"
      FitImage_DotNet "Msxml2.ServerXMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
    case "DOTNET3"
      FitImage_DotNet "Microsoft.XMLHTTP",DotNetResize,imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect
    end select
  end select
end sub

function DetectImageComponent(DotNetResize)
  Dim objPictureProcessor, objASPjpeg, AspImage, AspSmart, objImgWriter, objAspThumb, ImageComponent
  ImageComponent = ""
  if Application("ResizeAutoComponent112") = "" then
    on error resume next
   'Check for our own Picture Processor
    err.clear
    Set objPictureProcessor = Server.CreateObject("COMobjects.NET.PictureProcessor")
    if err.number = 0 then
      Set objPictureProcessor = nothing
      ImageComponent = "PICPROC"
    else
     'Check for ShotGraph
      err.clear
      Set objShotGraph = Server.CreateObject("shotgraph.image")
      'Response.Write err & " - " & err.number & ":" & err.description & "<br/>"
      if err.number = 0 then
        Set objShotGraph = nothing
        ImageComponent = "SHOTGRAPH"
      else
     'Check for AspJpeg
      err.clear
      Set objASPjpeg = Server.CreateObject("Persits.Jpeg")
      'Response.Write err & " - " & err.number & ":" & err.description & "<br/>"
      if err.number = 0 then
        Set objASPjpeg = nothing
        ImageComponent = "ASPJPEG"
      else
        'Check for AspImage
        err.clear
        Set AspImage = Server.CreateObject("AspImage.Image")
        if err.number = 0 then
          Set AspImage = nothing
          ImageComponent = "ASPIMAGE"
        else
          'Check for AspSmart
          err.clear
          Set AspSmart = Server.CreateObject("aspSmartImage.SmartImage")
          if err.number = 0 then
            Set AspSmartImage = nothing
            ImageComponent = "ASPSMART"
          else
            'Check for ImgWriter
            err.clear
            Set objImgWriter = Server.CreateObject("softartisans.ImageGen")
            if err.number = 0 then
              Set objImgWriter = nothing
              ImageComponent = "IMGWRITER"
            else
              'Check for AspThumb
              err.clear
              Set objAspThumb = Server.CreateObject("briz.AspThumb")
              if err.number = 0 then
                Set objAspThumb = nothing
                ImageComponent = "ASPTHUMB"
              else
              	if DetectDotNetComponent(DotNetResize) <> "" then
                	ImageComponent = "ASP.NET"
                end if
              end if
            end if
          end if
        end if
      end if
			end if
    end if
    on error goto 0
    Application("ResizeAutoComponent112") = ImageComponent
  else
	'use application var
    ImageComponent = Application("ResizeAutoComponent112")
  end if
  if ImageComponent = "" then
  	Response.Write "SMART IMAGE PROCESSOR ERROR: Can not detect any Resize Server Components!<br/>Please install at least the supplied server component. Read the online docs for more info."
  	Response.End
  end if 
  
  DetectImageComponent = ImageComponent
end function

function DetectDotNetComponent(DotNetResize)
  Dim DotNetImageComponent, ResizeComUrl, LastPath
  if Application("ResizeDotNetComponent112") = "" then
    DotNetImageComponent = ""
    ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
    LastPath = InStrRev(ResizeComUrl,"/")
    if LastPath > 0 then
      ResizeComUrl = left(ResizeComUrl,Lastpath)
    end if
    ResizeComUrl = ResizeComUrl & DotNetResize
    'Response.Write ResizeComUrl & "<br/>"
    
    'Check for ASP.NET 1
    if DotNetCheckComponent("Msxml2.ServerXMLHTTP.3.0", ResizeComUrl) = true then 
		  DotNetImageComponent = "DOTNET1"
    else
		  if DotNetCheckComponent("Msxml2.ServerXMLHTTP", ResizeComUrl) = true then
        DotNetImageComponent = "DOTNET2"
			else
        if DotNetCheckComponent("Microsoft.XMLHTTP", ResizeComUrl) = true then
          DotNetImageComponent = "DOTNET3"
				else
				end if
			end if
    end if
    on error goto 0
    Application("ResizeDotNetComponent112") = DotNetImageComponent
  else 'use application var
    DotNetImageComponent = Application("ResizeDotNetComponent112")
  end if
  DetectDotNetComponent = DotNetImageComponent
end function

function DotNetCheckComponent(DotNetObj, ResizeComUrl)
  dim objHttp, Detection
	Detection = false
  on error resume next
  err.clear
	'response.write("Checking "&DotNetObj&"<br/>")
  Set objHttp = Server.CreateObject(DotNetObj)
  if err.number = 0 then
  	'response.write("Object "&DotNetObj&" created<br/>")
    objHttp.open "GET", ResizeComUrl, false
		if err.number = 0 then
      objHttp.Send ""
      if trim(objHttp.responseText) <> "" and instr(objHttp.responseText,"@ Page Language=""C#""") = 0 then
        Detection = true
      end if
		end if
    Set objHttp = nothing
  End if
  on error goto 0
 	'response.write("Detection is "&Detection&"<br/>")
  DotNetCheckComponent = Detection
end function


sub FitImage_PicProc(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objPictureProcessor, intNewWidth, intNewHeight
  on error resume next
  Set objPictureProcessor = Server.CreateObject("COMobjects.NET.PictureProcessor")
  if err.number <> 0 then
    Response.Write "ERROR: Picture Processor Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objPictureProcessor.LoadFromFile imgFile
  objPictureProcessor.Quality = Quality
	if aspect = true Then
  	calculateNewImageSize objPictureProcessor.Width, objPictureProcessor.Height, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
	else
		intNewWidth = maxWidth
		intNewHeight = maxHeight
	end if
  objPictureProcessor.Resize intNewWidth, intNewHeight
  objPictureProcessor.SaveToFileAsJpeg newImgFile
  Set objPictureProcessor = nothing
end sub

sub FitImage_ShotGraph(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objShotGraph, intNewWidth, intNewHeight, xsize, ysize
  on error resume next
  Set objShotGraph = Server.CreateObject("shotgraph.image")
  if err.number <> 0 then
    Response.Write "ERROR: ShotGraph Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objShotGraph.GetFileDimensions imgFile, xsize, ysize
	if aspect = true then
  	calculateNewImageSize xsize, ysize, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
	else
		intNewWidth = maxWidth
		intNewHeight = maxHeight
	end if
	objShotGraph.CreateImage intNewWidth, intNewHeight, 8
	objShotGraph.InitClipboard xsize, ysize
	objShotGraph.SelectClipboard True
	objShotGraph.ReadImage imgFile, pal, 0, 0
	objShotGraph.Resize 0, 0, intNewWidth, intNewHeight, 0, 0, xsize, ysize, 3
	objShotGraph.SelectClipboard False
	objShotGraph.Sharpen
	objShotGraph.JpegImage quality, 0, newImgFile
  Set objShotGraph = nothing
end sub

sub FitImage_AspJpeg(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objAspJpeg, intNewWidth, intNewHeight
  on error resume next
  Set objAspJpeg = Server.CreateObject("Persits.Jpeg")
  if err.number <> 0 then
    Response.Write "ERROR: AspJpeg Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objAspJpeg.Open imgFile
  objAspJpeg.Quality = Quality
	if aspect = true then
  	calculateNewImageSize objAspJpeg.OriginalWidth, objAspJpeg.OriginalHeight, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
	else
		intNewWidth = maxWidth
		intNewHeight = maxHeight
	end if
  objAspJpeg.Width = intNewWidth
  objAspJpeg.Height = intNewHeight
  objAspJpeg.Sharpen 1, 150
  objAspJpeg.Save newImgFile
  Set objAspJpeg = nothing
end sub

sub FitImage_AspImage(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objAspImage, intNewWidth, intNewHeight, ver
  on error resume next
  Set objAspImage = Server.CreateObject("AspImage.Image")
  if err.number <> 0 then
    Response.Write "ERROR: AspImage Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objAspImage.LoadImage imgFile
  objAspImage.JPEGQuality = Quality
  if aspect = true then
  calculateNewImageSize objAspImage.MaxX, objAspImage.MaxY, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
  else
    intNewWidth = maxWidth
    intNewHeight = maxHeight
  end if
  ver = Split(objAspImage.Version, ".")
  If CInt(ver(0)) >= 2 Then
    objAspImage.ResizeR intNewWidth, intNewHeight
    objAspImage.Sharpen 1
  Else
    objAspImage.Resize intNewWidth, intNewHeight
  End If
  objAspImage.FileName = newImgFile
  objAspImage.SaveImage
  Set objAspImage = nothing
end sub

sub FitImage_AspSmart(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objAspSmart, intNewWidth, intNewHeight
  on error resume next
  Set objAspSmart = Server.CreateObject("aspSmartImage.SmartImage")
  if err.number <> 0 then
    Response.Write "ERROR: AspSmart Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objAspSmart.OpenFile CStr(imgFile)
  objAspSmart.Quality = Quality
	if aspect = true then
  	calculateNewImageSize objAspSmart.OriginalWidth, objAspSmart.OriginalHeight, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
	else
		intNewWidth = maxWidth
		intNewHeight = maxHeight
	end if
  objAspSmart.Resample CInt(intNewWidth), Cint(intNewHeight)
  objAspSmart.Sharpen
  objAspSmart.SaveFile newImgFile
  Set objAspSmart = nothing
end sub

sub FitImage_ImgWriter(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objImgWriter, intNewWidth, intNewHeight
  on error resume next
  Set objImgWriter = Server.CreateObject("softartisans.ImageGen")
  if err.number <> 0 then
    Response.Write "ERROR: ImgWriter Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objImgWriter.LoadImage imgFile
  objImgWriter.ImageQuality = Quality
	if aspect = true then
  	calculateNewImageSize objImgWriter.Width, objImgWriter.Height, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
	else
		intNewWidth = maxWidth
		intNewHeight = maxHeight
	end if
  objImgWriter.ResizeImage intNewWidth, intNewHeight
  objImgWriter.SharpenImage 80
  objImgWriter.SaveImage 0,3,newImgFile
  Set objImgWriter = nothing
end sub

sub FitImage_AspThumb(imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objAspThumb, intNewWidth, intNewHeight
  on error resume next
  Set objAspThumb = Server.CreateObject("briz.AspThumb")
  if err.number <> 0 then
    Response.Write "ERROR: ImgWriter Server Component is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  objAspThumb.Load imgFile
  objAspThumb.EncodingQuality = Quality
  if aspect = true then
  	calculateNewImageSize objAspThumb.Width, objAspThumb.Height, maxWidth, maxHeight, intNewWidth, intNewHeight, saveWidth, saveHeight, isNoThumb
  Else
    intNewWidth = maxWidth
    intNewHeight = maxHeight
  end If
  objAspThumb.Resize intNewWidth, intNewHeight
  objAspThumb.Save newImgFile
  Set objAspThumb = nothing
end sub


sub FitImage_DotNet(DotNetComp, DotNetResize, imgFile,newImgFile,maxWidth,maxHeight,Quality,saveWidth,saveHeight,isNoThumb,aspect)
  Dim objHttp, objText, ResizeComUrl, ResizeParams, LastPath, newSize, netaspect
	if aspect = true then
		netaspect = "true"
	else
		netaspect = "false"
	end if
  ResizeParams = "?f=" & Server.UrlEncode(imgFile) & "&nf=" & Server.UrlEncode(newImgFile) & "&w=" & maxWidth & "&h=" & maxHeight & "&q=" & Quality & "&a=" & netaspect
  ResizeComUrl = "http://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("PATH_INFO")
  LastPath = InStrRev(ResizeComUrl,"/")
  if LastPath > 0 then
    ResizeComUrl = left(ResizeComUrl,Lastpath)
  end if
  ResizeComUrl = ResizeComUrl & DotNetResize & ResizeParams

  on error resume next
  set objHttp = Server.CreateObject(DotNetComp)
  if err.number <> 0 then
    Response.Write "ERROR: ASP.NET (" & DotNetComp & ") is not installed!<br/>Please select a different Server Component and try again"
    Response.End
  end if
  on error goto 0
  
  objHttp.open "GET", ResizeComUrl, false
  objHttp.Send ""
	objText = objHttp.responseText
  ' Check notification validation
  if (objHttp.status <> 200 ) then
    ' HTTP error handling
    Response.Write "HTTP ERROR: " & objHttp.status & "<br/>"
    Response.Write "Returned:<br/>" & objHttp.responseText 
    Response.End
  elseif (left(objText, 6) = "Error:") then
  	Response.Write objHttp.responseText
  	Response.End
  elseif (right(objText, 4) = "DONE") then
    if (len(objText) > 4 and instr(objText, ";") > 0) then
			newSize = split(objText, ";")
 		  if saveWidth <> "" and isNoThumb then
			  saveWidth = LCase(saveWidth)
 	      if UploadRequest.Exists(saveWidth) then
 	        if UploadRequest.Item(saveWidth).Exists("Value") then
 	          UploadRequest.Item(saveWidth).Item("Value") = newSize(0)
 	        end if  
 	      end if  
		  end if
		  if saveHeight <> "" and isNoThumb then
			  saveHeight = LCase(saveHeight)
 	      if UploadRequest.Exists(saveHeight) then
 	        if UploadRequest.Item(saveHeight).Exists("Value") then
 	          UploadRequest.Item(saveHeight).Item("Value") = newSize(1)
 	        end if  
 	      end if  
 	 	  end if
		end if
  else
    if trim(objHttp.responseText) = "" or instr(objHttp.responseText,"@ Page Language=""C#""") > 0 then
      Response.Write "DOT NET Unsupported"
      Response.End
    end if
  end if
  Set objHttp = Nothing
end sub

sub calculateNewImageSize(curWidth, curHeight, maxWidth, maxHeight, newWidth, newHeight, saveWidth, saveHeight, isNoThumb)
  if maxWidth < curWidth or maxHeight < curHeight then
    if maxWidth >= maxHeight then
      newWidth = CInt(maxHeight*(curWidth/curHeight))
      newHeight = maxHeight
    else
      newWidth = maxWidth
      newHeight = CInt(maxWidth*(curHeight/curWidth))
    end if
    if newWidth > maxWidth then
      newWidth = maxWidth
      newHeight = CInt(maxWidth*(curHeight/curWidth))
    end if
    if newHeight > maxHeight then
      newWidth = CInt(maxHeight*(curWidth/curHeight))
      newHeight = maxHeight
    end if
  else
    newWidth = curWidth
    newHeight = curHeight
  end if
	if saveWidth <> "" and isNoThumb then
		saveWidth = LCase(saveWidth)
    if UploadRequest.Exists(saveWidth) then
      if UploadRequest.Item(saveWidth).Exists("Value") then
        UploadRequest.Item(saveWidth).Item("Value") = newWidth
      end if  
    end if  
	end if
	if saveHeight <> "" and isNoThumb then
		saveHeight = LCase(saveHeight)
    if UploadRequest.Exists(saveHeight) then
      if UploadRequest.Item(saveHeight).Exists("Value") then
        UploadRequest.Item(saveHeight).Item("Value") = newHeight
      end if  
    end if  
	end if
end sub

Sub ResizeUploadedFiles(RUF_Component, RUF_DotNetResize, RUF_path, RUF_Suffix, RUF_maxWidth, RUF_maxHeight, RUF_Quality, RUF_RemoveOrig, RUF_saveWidth, RUF_saveHeight, RUF_aspect, RUF_nameConflict, RUF_ResizeFields)
  Dim RUF_keys, RUF_KeysCount, RUF_i, RUF_curKey, RUF_fileName, RUF_fso, RUF_newFileName, RUF_curPath, RUF_curName, RUF_curExt, RUF_lastPos, RUF_orgCurPath, RUF_FileExist
  if RUF_path <> "" and right(RUF_path,1) <> "/" then RUF_path = RUF_path & "/"
  Set RUF_fso = Server.CreateObject("Scripting.FileSystemObject")  
  if RUF_maxWidth <> "" then
		RUF_maxWidth = Cint(RUF_maxWidth)
	else
		RUF_maxWidth = 100000
	end if
  if RUF_maxHeight <> "" then
  	RUF_maxHeight  = Cint(RUF_maxHeight)
	else
		RUF_maxHeight = 100000
	end if
	
	if RUF_ResizeFields <> "" then
  	RUF_keys = Split(RUF_ResizeFields, ",")
  	RUF_KeysCount = UBOUND(RUF_Keys)
  else
  	RUF_keys = UploadRequest.Keys
  	RUF_KeysCount = UploadRequest.Count - 1
  end if
  
  for RUF_i = 0 to RUF_KeysCount
    RUF_curKey = Trim(LCase(RUF_keys(RUF_i)))
    if UploadRequest.Exists(RUF_curKey) then
      if UploadRequest.Item(RUF_curKey).Exists("FileName") then    
  	    if UploadRequest.Item(RUF_curKey).Item("FileName") <> "" then    
          RUF_fileName = UploadRequest.Item(RUF_curKey).Item("Value")
          if RUF_fileName <> "" then
            RUF_curPath = "" : RUF_curName = "" : RUF_curExt = ""
            RUF_lastPos = InStrRev(RUF_fileName,"/")
            if RUF_lastPos > 0 then
              RUF_curPath = mid(RUF_fileName,1,RUF_lastPos)	
              RUF_curName = mid(RUF_fileName,RUF_lastPos+1,Len(RUF_fileName)-RUF_lastPos)	
              RUF_fileName = RUF_curName
            Else
              RUF_curName = RUF_fileName	
            end if
            RUF_lastPos = InStrRev(RUF_curName,".")
            if RUF_lastPos > 0 then
              RUF_curExt = mid(RUF_curName,RUF_lastPos+1,Len(RUF_curName)-RUF_lastPos)	
              RUF_curName = mid(RUF_curName,1,RUF_lastPos-1)
            end if
            RUF_curExt = LCase(RUF_curExt)
      		RUF_orgCurPath = RUF_curPath
            if RUF_curPath = "" then RUF_curPath = RUF_path
            if RUF_fso.FileExists(GetPhysicalPath(RUF_curPath & RUF_fileName)) then
              if RUF_curExt = "jpg" or RUF_curExt = "jpeg" or RUF_curExt = "gif" or RUF_curExt = "bmp" or RUF_curExt = "png" or RUF_curExt = "pgm" or RUF_curExt = "tga" or RUF_curExt = "tif" or RUF_curExt = "tiff" or RUF_curExt = "jfif" then
                RUF_newFileName = RUF_curName & RUF_Suffix & ".jpg"
				RUF_FileExist = False
				If RUF_fso.FileExists(GetPhysicalPath(RUF_curPath & RUF_newFileName)) Then
					RUF_FileExist = true
				End If    
				if RUF_nameConflict = "uniq" and RUF_FileExist and LCase(RUF_fileName) <> LCase(RUF_newFileName) then
					Begin_Name_Num = 0
					while RUF_FileExist
						Begin_Name_Num = Begin_Name_Num + 1
						RUF_newFileName = RUF_curName & "_" & Begin_Name_Num & RUF_Suffix & ".jpg"
						RUF_FileExist = RUF_fso.FileExists(GetPhysicalPath(RUF_curPath & RUF_newFileName))
					Wend
					'UploadRequest.Item(RUF_curKey).Item("FileName") = RUF_newFileName
					'UploadRequest.Item(RUF_curKey).Item("Value") = RUF_curPath & RUF_newFileName
				end If
				FitImage_Comp RUF_Component, RUF_DotNetResize, GetPhysicalPath(RUF_CurPath & RUF_fileName), GetPhysicalPath(RUF_curPath & RUF_newFileName), RUF_maxWidth, RUF_maxHeight, RUF_Quality, RUF_saveWidth, RUF_saveHeight, RUF_RemoveOrig, RUF_aspect
				if RUF_RemoveOrig then
					if LCase(RUF_fileName) <> LCase(RUF_newFileName) then
						RUF_fso.DeleteFile GetPhysicalPath(RUF_curPath & RUF_fileName)
					end if  
					if RUF_orgCurPath <> "" then
						UploadRequest.Item(RUF_curKey).Item("Value") = RUF_orgCurPath & RUF_newFileName
					else
						UploadRequest.Item(RUF_curKey).Item("Value") = RUF_newFileName
					end if
					UploadRequest.Item(RUF_curKey).Item("FileName") = RUF_newFileName
				end If
              end if  
            end if
          end if
        end if
      end if
    end if
  next
End Sub

Function GetPhysicalPath(strPath)
  Dim curPath, arrPath, newPath, backCount, i
  newPath = strPath
  If newPath = "" Then newPath = "."
  If InStr(newPath, "\") = 0 Then
    If InStr(newPath, "../") = 0 Then
      newPath = Server.MapPath(newPath)
    Else
      curPath = Server.MapPath(".")
      If newPath = "" Then
        GetPhysicalPath = curPath
        Exit Function
      End If
      arrPath = Split(curPath & "\" & Replace(newPath, "/", "\"), "\")
      newPath = ""
      backCount = 0
      For i = UBound(arrPath) To 1 Step -1
        If arrPath(i) = ".." Then
          backCount = backCount + 1
        Else
          If backCount = 0 Then
            newPath = arrPath(i) & "\" & newPath
          Else
            backCount = backCount - 1
          End If
        End If
      Next
      If backCount > 0 Then
        ' Incorrect Path
        Response.Write "Incorrect Path: " & strPath
        Response.End
      End If
      newPath = arrPath(0) & "\" & Left(newPath, Len(newPath) - 1)
    End If
  End If
  If Right(newPath, 1) = "\" Then newPath = Left(newPath, Len(newPath) - 1)
  GetPhysicalPath = newPath
End Function

</SCRIPT>