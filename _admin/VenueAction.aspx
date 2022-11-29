<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/sitewide.aspx" -->
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
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\venues")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	Dim v As New Venue
	v.name = db.Prepare(""& Request.Form("name"))
	v.text = db.Prepare(""& Request.Form("content"))
	v.image = db.Prepare(""& sFiles(0))
	v.quickFacts = db.Prepare(""& Request.Form("quickfacts"))
	v.url = db.Prepare(""& Request.Form("url"))
	Dim exc As Integer = 0
	Dim pref AS Integer = 0
	if Request.Form("exclusive") = "1" then
		exc = 1
	end if
	if Request.Form("preferred") = "1" then
		pref = 1
	end if
	
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Const SORT_WHERE As String = ""
	Const HAS_SORT_ORDER As Boolean = true			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = true				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "img" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblVenues"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "VenueList.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "VenueDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "name","text","img","quickfacts","exclusive","preferred","url","SortOrder" }
	Dim sValues() As String = { v.name,v.text,v.image,v.quickFacts,exc,pref,v.url, 1 }
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
	
	sBridgeQueries = "delete from brgVenueLandmarks where venueId = "& iID
	Dim _ls() As String = split(Request.Form("landmarks[]"),",")
	for each _l As String in _ls
		sBridgeQueries += "insert into brgVenueLandmarks (landmarkId, venueId) values ("& _l &","& iID &"); |"
	next _l
	' perform the actions

%>
<!--#include file="templates/Action.aspx" -->