<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="true" %>
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
	Dim sPageObject As String = "Gallery"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "GalleryAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblGallery"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	Dim sType() As String = { "Cake","Events","Food" }
	Dim t As Integer = CInt(Request.Form("t"))
	
	Dim g As New Gallery
	Dim _is As String = ""
	Dim _is_thumb As String = ""
	
	g.type = t
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		g.id = CInt(db.GetItem("id"))
		g.type = CInt(db.GetItem("type"))
		g.image = db.GetItem("image")
		g.caption = db.GetItem("caption")
		g.thumbnail = db.GetItem("thumbnail")
	end while	
	db.CloseConnection()
	
	if len(g.image) > 0 then
		_is = "<a href='../upload/gallery/"& g.image &"'>"& g.image &"</a> <a href='GalleryAction.aspx?Action=RemoveFile&Field=image&id="& g.id &"'>[x]</a>"
	end if
		
	if len(g.thumbnail) > 0 then
		_is_thumb = "<a href='../upload/gallery/thumbs/"& g.thumbnail &"'>"& g.thumbnail &"</a> <a href='GalleryAction.aspx?Action=RemoveFile&Field=thumbnail&id="& g.id &"'>[x]</a>"
	
	end if
	
	' construct the form modules
	oForm.NewFormElement("Assign to Gallery","select","type",sType,g.type,false,Nothing)
	oForm.NewFormElement("Upload Image","file","image","",50,false,_is)
	
	oForm.NewFormElement("Upload Thumbnail (Optional)","file","thumbnail","",50,false,_is_thumb)
	
	oForm.NewFormElement("Caption","text","caption",g.caption,100,false,Nothing)
	
%>
<!--#include file="templates/Detail.aspx" -->