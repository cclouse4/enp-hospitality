<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="includes/image.aspx" -->

<!--#include file="../_src/classes.aspx"-->
<%
	' page variables - do not change
	Dim sAction As String = Request.Form("Action")
	Dim iSortOrder As Integer = CInt(Request.QueryString("SortID"))
	Dim iSortDirection As Integer = Request.QueryString("SortDirection")
	Dim iSortID As Integer
	Dim iOldSortOrder As Integer
	Dim i As Integer
	Dim sPostFieldName As String = Request.QueryString("Field")
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	
	db.OpenConnection()
	

	' upload files
	try
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\upload")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	

	
	Dim g As New Gallery
	
	g.image = Utilities.GetFilenameByKey("image")
	g.thumbnail = Utilities.GetFilenameByKey("thumbnail")
	

	if g.thumbnail <> "" then
		' Move to Thumbnail folder
		if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & g.thumbnail) ) then
			System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & g.thumbnail)
		end if
		
		System.IO.File.Move(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\upload\" & g.thumbnail, Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & g.thumbnail)
	end if
	
	if g.image <> "" then
		' Move to Image folder
		
		if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\" & g.image) ) then
			System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\" & g.image)
		end if


		System.IO.File.Move(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\upload\" & g.image, Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\" & g.image)
		
		' Make Thumbnail if None
		if g.thumbnail = "" then
		
			if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & g.image) ) then
				System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs\" & g.image)
			end if
			
			if (System.IO.File.Exists(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile\" & g.image) ) then
				System.IO.File.Delete(Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile\" & g.image)
			end if
		
			' strUploadFolder, strDestinationFolder, strUploadFile, strDestinationFile, intWidth, intHeight, intQuality, intKeepAspectRatio
			Call MakeThumb(ConfigurationSettings.AppSettings("uploadPath") &"\gallery\", ConfigurationSettings.AppSettings("uploadPath") &"\gallery\thumbs", g.image, g.image, 150, 150, 81, 1)
			
			Call MakeThumb(ConfigurationSettings.AppSettings("uploadPath") &"\gallery\", ConfigurationSettings.AppSettings("uploadPath") &"\gallery\mobile", g.image, g.image, 320, 0, 61, 1)
        
			
			g.thumbnail = g.image
			
		end if
		
	end if
	
	g.image = db.Prepare(""& g.image)
	g.thumbnail = db.Prepare(""& g.thumbnail)
	
	
	g.caption = db.Prepare(""& Request.Form("caption"))
	g.type = CInt(Request.Form("type"))
	
	
	
	
	if CInt(Request.QueryString("t")) > 0 then
		g.type = CInt(Request.QueryString("t"))
	end if
	
	if g.type = 0 then
		g.type = 1
	end if
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Dim SORT_WHERE As String = " where type = "& g.type
	Const HAS_SORT_ORDER As Boolean = true			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = true				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "image", "thumbnail" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblGallery"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "GalleryList.aspx?t="& g.type		' url to redirect to once action is complete
	Dim sDetailURL As String = "GalleryDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "type","image","caption","SortOrder", "thumbnail" }
	Dim sValues() As String = { g.type,g.image,g.caption,1, g.thumbnail}
	'###############################################################################################################################

	' format the variables if there is a query string
	Dim iID As Integer = CInt(Request.Form(sPrimaryKey))
	if len(Request.QueryString(sPrimaryKey)) > 0 then
		iID = CInt(Request.QueryString(sPrimaryKey))
	end if
	
	if len(Request.QueryString("Action")) > 0 then
		sAction = Request.QueryString("Action")
	end if
	
	Dim _s As String
	Dim _SortOrder As String
	
	' perform the actions

%>
<!--#include file="templates/Action.aspx" -->