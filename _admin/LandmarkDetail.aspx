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
	Dim sPageObject As String = "Landmark"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "LandmarkAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "id"					' primary key of the table associated with this module
	Dim sTableName As String = "tblLandmarks"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim bActive As Boolean = true
	
	Dim sName As String
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		sName = db.GetItem("landmark")
	end while	
	db.CloseConnection()
		
	' construct the form modules
	oForm.NewFormElement("Landmark Name","text","landmark",sName,50,false,Nothing)
	
%>
<!--#include file="templates/Detail.aspx" -->