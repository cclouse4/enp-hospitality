
<script language="vb" runat="server">
 ' Resize and Crop Image. 
' Resizes image to fill the specified dimensions and crops off the remainder. Will always result in an image of the exact specified dimensions.
Public Shared Function ResizeAndCropImage(ByVal originalImage As System.Drawing.Image, ByVal newWidth As Integer, ByVal newHeight As Integer) As System.Drawing.Image
    Dim newImage As System.Drawing.Image
    Dim tempBmp As System.Drawing.Bitmap
    Dim tempWidth As Integer, tempHeight As Integer
    Dim originalAspect As Single, newAspect As Single
    
    tempWidth = 1
    tempHeight = 1
    
    ' Aspect ratios
    originalAspect = (CSng(originalImage.Width) / CSng(originalImage.Height))
    newAspect = (CSng(newWidth) / CSng(newHeight))
    
    
    ' Is image proportionally taller or wider than the new one? We want to fill the new image completely.
    
    
    If originalAspect <= newAspect Then
        tempWidth = newWidth
        tempHeight = CInt((CSng(newWidth) / originalAspect))
    Else
        tempWidth = CInt((CSng(newHeight) * originalAspect))
        tempHeight = newHeight
    End If
    
    
    tempBmp = New Bitmap(originalImage)
    
    System.Console.WriteLine(tempHeight)
    System.Console.WriteLine(tempWidth)
    
    Dim scaledImage As New Bitmap(tempWidth, tempHeight)
    Dim croppedImage As New Bitmap(newWidth, newHeight)
    
    scaledImage.SetResolution(originalImage.HorizontalResolution, originalImage.VerticalResolution)
    croppedImage.SetResolution(originalImage.HorizontalResolution, originalImage.VerticalResolution)
    
    ' Resize the image to our new temporary size
    Dim ConvertGraphics As Graphics
    ConvertGraphics = Graphics.FromImage(scaledImage)
    ConvertGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
    ConvertGraphics.SmoothingMode = SmoothingMode.HighQuality
    ConvertGraphics.DrawImage(originalImage, 0, 0, tempWidth, tempHeight)
    
    ' Crop the image to our final size
    ConvertGraphics = Graphics.FromImage(croppedImage)
    ConvertGraphics.InterpolationMode = InterpolationMode.HighQualityBicubic
    ConvertGraphics.DrawImage(scaledImage, New Rectangle(0, 0, newWidth, newHeight), 0, 0, newWidth, newHeight, _
    	GraphicsUnit.Pixel)
    
    Return croppedImage
End Function

' Need encoder to save image using compression
Public Shared Function GetEncoder(ByVal mimeType As String) As ImageCodecInfo
    Dim codecs As ImageCodecInfo() = ImageCodecInfo.GetImageEncoders()
    For Each codec As ImageCodecInfo In codecs
        If (codec.MimeType = mimeType) Then
            Return codec
        End If
    Next
    Return Nothing
End Function

' Encoder params specifiy the quality of the output image
Public Shared Function GetEncoderParams(ByVal quality As Integer) As EncoderParameters
    Dim encodeParams As New EncoderParameters()
    Dim qualityArray As Long() = New Long(0) {}
    qualityArray(0) = quality
    Dim encoderParam As New EncoderParameter( System.Drawing.Imaging.Encoder.Quality, qualityArray)
    encodeParams.Param(0) = encoderParam
    Return encodeParams
    
End Function

	</script>
