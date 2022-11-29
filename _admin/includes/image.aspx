<%@Import Namespace="System.Drawing" %>
<%@Import Namespace="System.Drawing.Imaging" %>
<%@Import Namespace="System.Drawing.Drawing2D" %>

<script runat="server">
	
	Function ThumbnailCallback() as Boolean
	  Return False
	End Function

	' FUNCTION Make Thumb
	' Parameters
	' - strUploadFolder - Folder for Original Image
	' - strDestinationFolder - Folder for Destination Image
	' - strUploadFile - Original Image Name
	' - strDestinationFile - Destination Image Name
	' - intWidth - New Image Width in Pixels
	' - intHeight - New Image Height in Pixels
	' - intQuality - Quality of Optimized Image 0-100 - 100 BEST
	' - intKeepAspectRatio - 1 or 0 - Keep Ratio of Image
	function MakeThumb(strUploadFolder, strDestinationFolder, strUploadFile, strDestinationFile, intWidth, intHeight, intQuality, intKeepAspectRatio) As String
		
		Dim originalimg, thumb As System.Drawing.Image
		Dim strFileName As String
		Dim strExtension As String
		Dim strResponseType As String
		Dim width, height As Integer
		Dim intChange As Integer
		
		' Figure Out Extension/Response Type
		strExtension = ""
		If InStr(strUploadFile, ".") > 0 Then
			strExtension = Right(strUploadFile, Len(strUploadFile) - InStr(strUploadFile, "."))
		End If
		
		Select Case UCase(strExtension)
			Case "JPEG", "JPG"
				strResponseType = "image/jpeg"
			Case "GIF"
				strResponseType = "image/gif"
			Case "BMP"
				strResponseType = "image/bmp"
			Case "PNG"
				strResponseType = "image/png"
			Case Else
				MakeThumb = "-1"
				Exit Function
		End Select
		'Response.ContentType = strResponseType
		
		' Get the Full Size Image
		strFileName = strUploadFolder & "/" & strUploadFile
		
		
		originalimg = originalimg.FromFile(Request.ServerVariables("APPL_PHYSICAL_PATH") & strFileName) ' Fetch User Filename
		
		
		
		If intHeight = 0 And intWidth = 0 Then
			intChange = 0
			height = originalimg.height
			width = originalimg.width
			
			' DoubleCheck boundary
			If originalimg.width > 471 Then
				' Force 470
				width = 471
				intChange = 2
			End If
		Else
			IF intHeight <= 0 THEN 
				height = originalimg.height
				intChange = 2
			ELSE
				If intHeight > originalimg.height Then
					height = originalimg.height
					intChange = 2
				Else
					height = intHeight
				End If
			END IF
			
			IF intWidth <= 0 THEN 
				width = originalimg.width
				intChange = 1
			ELSE
				If intWidth > originalimg.width Then
					width = originalimg.width
					intChange = 2
				Else
					width = intWidth
				End If
			End If
	    End If
		
		'Response.Write(height & ":" & width)
		
		' Response.End()
		
	    'Response.ContentType = strResponseType
	    thumb = Produce_Thumbnail(originalimg, height, width, intKeepAspectRatio, intChange)
	    Dim codecEncoder As ImageCodecInfo = GetEncoder(strResponseType)
	    Dim quality As Integer = intQuality
	    Dim encodeParams As New EncoderParameters(1)
	    Dim qualityParam As New EncoderParameter(Imaging.Encoder.Quality, quality)
	    encodeParams.Param(0) = qualityParam
		
		strFileName = strDestinationFolder & "/" & strDestinationFile
	          thumb.Save(Request.ServerVariables("APPL_PHYSICAL_PATH") & strFileName, codecEncoder, encodeParams)
	
		originalimg.Dispose()
	    thumb.Dispose()

	end function
	
    Function GetEncoder(ByVal mimeType As String) As ImageCodecInfo
        Dim codecs() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()
        For Each codec As ImageCodecInfo In codecs
            If codec.MimeType = mimeType Then
                Return codec
            End If
        Next
    End Function
    
  
    Function Produce_Thumbnail(ByVal Source_Img As System.Drawing.Image, ByVal height As Integer, ByVal width As Integer, ByVal intKeepAspectRatio As Boolean, ByVal intChange As Integer) As System.Drawing.Image
        Dim Dest_Img As System.Drawing.Image
        
        Dim inp As New IntPtr()
        Dim Orginal_Width As Integer
        Dim Orginal_Height As Integer
        Dim Aspect_ratio As Integer
        Orginal_Width = Source_Img.Width
        Orginal_Height = Source_Img.Height

		Select Case intChange
			Case 1
		        Aspect_ratio = Orginal_Height / height
		
				' Turned Off Aspect Ratio
				If intKeepAspectRatio Then
		       		width = Orginal_Width / Aspect_ratio
		        End If
			Case 2
		        Aspect_ratio = Orginal_Width / width
		
				' Turned Off Aspect Ratio
				If intKeepAspectRatio Then
		       		height = Orginal_Height / Aspect_ratio
		        End If
			Case 0
				Aspect_ratio = Orginal_Width / width
		End Select
		
        Dim thumbnail As System.Drawing.Image = New Bitmap(width, height)
        Dim objGraphics As System.Drawing.Graphics
        objGraphics = System.Drawing.Graphics.FromImage(thumbnail)
        objGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
        objGraphics.SmoothingMode = SmoothingMode.HighQuality
        objGraphics.PixelOffsetMode = PixelOffsetMode.HighQuality
        objGraphics.CompositingQuality = CompositingQuality.HighQuality
        If Request.QueryString("thumb") = "false" Then
            Return Source_Img
        Else
            objGraphics.DrawImage(Source_Img, 0, 0, width, height)
            Return thumbnail
            
        End If
    End Function

</script>
