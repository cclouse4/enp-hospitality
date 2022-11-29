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
	Dim sPageObject As String = "User"					' The type of page, used for text display to identify what is going on
	Dim sFormActionURL As String = "UserAction.aspx"		' filename of the action page used for this module
	Dim sPrimaryKey As String = "UserID"					' primary key of the table associated with this module
	Dim sTableName As String = "tblUser"					' database table name
	' data variables for data retrieval
	Dim iID As Integer = CInt(Request.QueryString(sPrimaryKey))
	Dim i As Integer
	Dim sTitle As String
	Dim sPassword As String
	Dim sUserType() As String = { "Administrator", "User" }
	Dim iUserType As Integer
	Dim bActive As Boolean			
	
	db.RunQuery("select * from "& sTableName &" where "& sPrimaryKey &" = "& iID)
	while db.InResults()
		iID = db.GetItem(sPrimaryKey)
		sTitle = db.GetItem("Username")
		sPassword = db.GetItem("Password")
		iUserType = db.GetItem("UserType")
		bActive = db.GetItem("Active")
	end while

    db.CloseConnection()
	
	' construct the form modules
	oForm.NewFormElement("Username","text","title",sTitle,50,false,Nothing)
	oForm.NewFormElement("Password","text","password",sPassword,50,false,Nothing)
	oForm.NewFormElement("Type","select","type",sUserType,iUserType,false,Nothing)
	oForm.NewFormElement("Active","checkbox","active",1,0,bActive,Nothing)
%>
<!--#include file="templates/Detail.aspx" -->