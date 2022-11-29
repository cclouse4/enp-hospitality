<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" %>
<%@ Import Namespace="Chemistry" %>
<!--#include file="includes/menu_functions.aspx" -->
<!--#include file="_utilities/formbuilder_class.aspx" -->
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<%
	' page variables
	Dim sAction As String = Request.QueryString("Action")
	Dim oForm As New FormBuilder()
	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	' template variables -- These control the page fucntionality and display
	Dim sPageType As String = "Detail"					' Use Detail, List, Delete
	Dim sPageObject As String = "Venue"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "VenueAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblVenues"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	
	Dim v As New Venue
	Dim _is As String
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		v.name = db.GetItem("name")
		v.text = db.GetItem("text")
		v.image = db.GetItem("img")
		v.imageCaption = db.GetItem("imgcaption")
		v.capacity = db.GetItem("capacity")
		v.address = db.GetItem("address")
		v.url = db.GetItem("url")
	end while	
	
	Dim _vl As String
	db.RunQuery("select * from brgVenueLandmarks where venueId = "& iID)
	while db.InResults()
		_vl += db.GetItem("landmarkId") &"|"
	end while
	
	Dim _select As String
	db.RunQuery("select * from tblLandmarks order by landmark asc")
	while db.InResults()
		_select += "<input type='checkbox' name='landmarks[]' value='"& db.GetItem("id") &"'"
		for each __vl As String in split(_vl,"|")
			try				
				if CInt(__vl) = CInt(db.GetItem("id")) and CInt(__vl) > 0 then
					_select += " CHECKED"
					exit for
				end if
			catch
			end try
		next __vl
		_select += "> "& db.GetItem("landmark") &"<br/>"
	end while	
	db.CloseConnection()
		
	if len(v.image) > 0 then
		_is = "<a href='../upload/venues/"& v.image &"' target='_blank'>"& v.image &"</a> <a href='VenueAction.aspx?Action=RemoveFile&Field=img&id="& iID &"'>[x]</a>"
	end if
		
	' construct the form modules
	oForm.NewFormElement("Name","text","name",v.name,50,false,Nothing)
	oForm.NewFormElement("Image","file","img",v.image,50,false,_is)
	oForm.NewFormElement("Image Caption","text","imgCaption",v.imageCaption,50,false,Nothing)
	oForm.NewFormElement("Capacity","text","capacity",v.capacity,50,false,Nothing)
	oForm.NewFormElement("Address","text","address",v.address,50,false,Nothing)
	oForm.NewFormElement("URL","text","url",v.url,50,false,Nothing)
	oForm.NewFormElement("Content","textarea","content",v.text,50,false,Nothing)
	oForm.NewFormElement("Landmakrs","multiCheckbox","",_select,0,false,Nothing)
	
	
%>
<!--#include file="templates/Detail.aspx" -->