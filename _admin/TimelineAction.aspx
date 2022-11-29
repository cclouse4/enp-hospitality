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
		Dim bFile As Boolean = Utilities.Upload(Request.Files,Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\page")
	catch e As Exception
		'response.write(e.message())
		'response.end()
	end try
	
	Dim sFiles() As String = Utilities.GetUploadedFileNames()
	' data variables	
	Dim p As New TimelineEntry
	p.id = CInt(Request.Form("id"))
	p.year = CInt(Request.Form("year"))
	p.name = db.Prepare(""& Request.Form("name"),false)
	p.mainText = db.Prepare(""& Request.Form("mainText"),false)

	
	'if len(Request.Form("landingPage")) > 0 then
	'	p.isLandingPage = cBool(Request.Form("landingPage"))
	'end if
	
	'if CInt(Request.QueryString("parentId")) > 0 then
	'	p.parentId = CInt(Request.QueryString("parentId"))
	'end if
	
	' used for template purposes. these control what goes into the database
	'###############################################################################################################################
	Dim sBridgeQueries As String = ""
	Dim SORT_WHERE As String = "where year = "& p.year
	Const HAS_SORT_ORDER As Boolean = true			' set this to true to enable sort actions
	Const HAS_FILES As Boolean = false				' set this to true to enable file deletion actions
	Const DEBUG As Boolean = false					' set this to true to output the query that bombed the script. false for live
	Dim sFileFieldName() As String = { "backgroundImage","image","tout1img","tout2img","tout3img" } ' name of the file field name, used for file deletion
	Dim sTableName As String = "tblTimeline"				' name of the database table to perform actions on
	Dim sPrimaryKey As String = "id"				' name of the primary key field in the database
	Dim sRedirectURL As String = "TimelineList.aspx"		' url to redirect to once action is complete
	Dim sDetailURL As String = "TimelineDetail.aspx"		' url to redirect to detail page
	Dim sColumnNames() As String = { "year","name","mainText","SortOrder" }
	Dim sValues() As String = { p.year, p.name, p.mainText, 1 }
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
	select case sAction
	case "Sort"
		if HAS_SORT_ORDER then
			if iSortDirection = 0 then
				db.RunQuery("select "& sPrimaryKey &", SortOrder from "& sTableName &" where year = "& CInt(Request.QueryString("year")) &" order by SortOrder asc")
				'response.write("select "& sPrimaryKey &", SortOrder from "& sTableName &" where parentId = "& CInt(Request.QueryString("parentId")) &" order by SortOrder asc")
				'response.end()
				while db.InResults()
					if iID = db.GetItem(sPrimaryKey) then
						iOldSortOrder = db.GetItem("SortOrder")
						exit while
					else
						iSortID = db.GetItem(sPrimaryKey)
						iSortOrder = db.GetItem("SortOrder")
					end if
				end while
				db.RunOther("update "& sTableName &" set SortOrder = "& iSortOrder &" where "& sPrimaryKey &" = "& iID)
				db.RunOther("update "& sTableName &" set SortOrder = "& iOldSortOrder &" where "& sPrimaryKey &" = "& iSortID)
			else if iSortDirection = 1 then
				db.RunQuery("select "& sPrimaryKey &", SortOrder from "& sTableName &" where year = "& CInt(Request.QueryString("year")) &" order by SortOrder desc")
				while db.InResults()
					if iID = db.GetItem(sPrimaryKey) then
						iOldSortOrder = db.GetItem("SortOrder")
						exit while
					else
						iSortID = db.GetItem(sPrimaryKey)
						iSortOrder = db.GetItem("SortOrder")
					end if
				end while
				db.RunOther("update "& sTableName &" set SortOrder = "& iSortOrder &" where "& sPrimaryKey &" = "& iID)
				db.RunOther("update "& sTableName &" set SortOrder = "& iOldSortOrder &" where "& sPrimaryKey &" = "& iSortID)
			end if
			Response.Redirect(sRedirectURL)
		end if
	end select
	
	'sBridgeQueries = "delete from brgContentTout where contentId = "& iID &" |"
	'for each _z As String in split(Request.Form("touts[]"),",")
	'	sBridgeQueries += "insert into brgContentTout (contentId, toutId) values ("& iID &","& _z &"); |"
	'next _z

%>
<!--#include file="templates/Action.aspx" -->