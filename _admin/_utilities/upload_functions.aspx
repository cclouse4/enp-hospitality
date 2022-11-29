<script runat="server">
' upload function
public function UploadFiles(ByVal sPath As String, ByVal oFiles As HttpFileCollection) As String()
	Dim sUploadName As String
	Dim sUploadDirectory As String
	Dim iCounter As Integer
	Dim oFileCollection As HttpFileCollection
	Dim oFile As HttpPostedFile
	Dim bUploaded As Boolean = False
	Dim sFileNames(oFiles.Count) As String
	Dim sUploadNameArray() As String
	
	oFileCollection = oFiles
	sUploadDirectory = Request.ServerVariables("APPL_PHYSICAL_PATH") &"bfcp\"& sPath
	
	' test if the folder exists, if not add the unitedway prefix to it
	Dim oFolder As New System.IO.DirectoryInfo(sUploadDirectory)
	if not oFolder.Exists then
		sUploadDirectory = Request.ServerVariables("APPL_PHYSICAL_PATH") &"bfcp\"& sPath
	end if
	
	' run through the files and process them
	for iCounter = 0 to oFileCollection.Count - 1
		oFile = oFileCollection.Item(iCounter)
		
		' test the validity of the file
		if not (oFile is nothing) and len(oFile.Filename) > 0 then
			' parse the filename
			sUploadName = oFile.Filename
			sUploadNameArray = sUploadName.split("\")
			if len(sUploadNameArray(UBound(sUploadNameArray))) > 0 then
				sUploadName = sUploadNameArray(UBound(sUploadNameArray))
			end if
			sUploadName = Replace(sUploadName, " ", "") ' remove spaces
			sUploadName = Replace(sUploadName, "&", "") ' remove &amp;
			sUploadName = Replace(sUploadName, """", "")' remove quotes
			
			' upload the file
			try
				oFile.SaveAs(sUploadDirectory &"\"& sUploadName)			
				bUploaded = True
				sFileNames(iCounter) = sUploadName
			catch e As Exception
				Response.Write(sUploadName &"<br/>")
				Response.Write(e.Message)
				bUploaded = False
				response.end()
			end try		
		end if
	next iCounter
	
	oFileCollection = nothing
	oFile = nothing
	
	if UBound(sFileNames) = 0 then
		Dim sReturn(100) As String
		return sReturn
	end if
	
	return sFileNames
end function

</script>