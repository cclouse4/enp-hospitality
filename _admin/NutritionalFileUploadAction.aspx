<%@ Page Language="VB" ContentType="text/html" ResponseEncoding="iso-8859-1" Debug="True" aspcompat=true  %>
<%@ Import Namespace="Chemistry" %>
<%@ Import Namespace="System.Data.OleDb" %>
<%@ Import Namespace="System.Data" %>
<%@ Import Namespace="System.Data.SqlClient" %>
<!--#include file="includes/sitewide.aspx" -->
<!--#include file="../_src/classes.aspx"-->
<!--#include file="_utilities/format_sql.aspx" -->
<%
	' page variables - do not change

	Dim db As New DBObject(ConfigurationSettings.AppSettings("connString"))
	db.OpenConnection()
	
	Server.scripttimeout = 3600000 
	
	
	' Database Variables
	dim rstQuery
	dim strSQL
	dim uploadI
	dim uploadIFile
	dim uploadIName
	dim uploadISize
	dim uploadIDirectory
	dim CurPath
	dim FileObjectI
	
    
	' upload files
	try
        Dim bFile As Boolean = Upload.FileHandler.Process(Request.Files, Server.MapPath("../") & ConfigurationSettings.AppSettings("uploadPath") &"\NutritionalFile\")
	catch e As Exception
	end try
	
    Dim sFiles() As String = Upload.FileHandler.getFilenames()
	uploadIName = Upload.FileHandler.GetFilenameByKey("UploadFile")

    
	Dim objFSO, objFile, strReadLine
	Dim strItemID, strItemType, strCategory, strDisplayName, strParentMenuItem, strIngredient, strPortionSize, strParent, strCalories, strFat, strCarbs, strProtein, strCholesterol, strDietaryFiber, strSodium, strPercentCaloriesFromFat, strPercentCaloriesFromCarbs
	Dim strMenuCategory, strMenuDisplayName, strActiveMenu, strShowNutritionals, strShowPrice, strDescription, strPrice, strOHPrice, strItemStartDate, strItemEndDate, strAllergens
    Dim strSmileyChoiceIcon, strGreatForTakeoutIcon, strEatnSmartIcon
	
	Dim arrItems
	
	Dim intCounter, strExtension, strFileType

	strExtension = Right(uploadIName, Len(uploadIName) - InStr(uploadIName, "."))
	
	Select Case LCase(strExtension)
		Case "xls", "xlsx"
		
			strFileType = "xls"
		Case Else
		
	End Select
	
	If Not strFileType = "xls" Then
		Response.Write("File needs to be type Excel (.xls)")
		Response.End()
	End If
	
	Dim s_FDir , s_MDBString
	Dim s_MDBConn, q_ExcelFile
	Dim s_Range, intColumnCounter
	Dim rowdt
    Dim strTableName As String = ""
    Dim intTableCount As Integer = 0
    Dim rowExcel
    
	'  & uploadIName
	
    Dim sPath = "\upload\NutritionalFile"
    Dim sUploadDirectory = Request.ServerVariables("APPL_PHYSICAL_PATH") & sPath
    
	s_FDir = sUploadDirectory & "\" & uploadIName
	
	s_FDir = Replace(s_FDir, "\\", "\")

    s_MDBString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & s_FDir & "; Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"""
				
	
    Dim oExcel, dt, cmdSelect, daCSV, dtCSV
    Dim i
    Dim strItem 
    Dim bFound As Boolean
    
 	oExcel = New System.Data.OleDb.OleDbConnection(s_MDBString)
	oExcel.Open()
    
    dt = oExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, New Object() {nothing, nothing, nothing, "Table"})


    
	Response.Write("<BR>")
    
    For Each rowdt In dt.Rows		
        strTableName = rowdt("TABLE_NAME")
        strTableName = Replace(strTableName, "'", "")
        
        If InStr(strTableName, "WELLLIFE") > 0 Or InStr(strTableName, "Sheet1") > 0 Then
            Exit For
        End If
        ' Response.Write("strTableName:" & strTableName & ":" & intTableCount & "<BR>")
        
        intTableCount += 1
    Next
    dt = Nothing

    
    If strTableName <> "" Then
 		
	
        strSQL = "DELETE FROM tblNutritionalFile"
        db.RunOther(strSQL)
    
        strSQL = "SELECT * FROM [" & strTableName & "]"
        cmdSelect = new OleDbCommand (strSQL, oExcel)
        daCSV = new OleDbDataAdapter()
        daCSV.SelectCommand  = cmdSelect
        dtCSV = new DataTable()
        daCSV.Fill(dtCSV)
    
        Dim dc As DataColumn
        Dim intStartColumn As Integer = 0
        Dim intEndColumn As Integer = 0
        
        Dim intColumn As Integer = 0
			
        intStartColumn = 0
        For Each dc in dtCSV.Columns
            ' Response.Write(intColumn & ":" & dc.ToString() & "<BR>")
            
            If dc.ToString() = "Allergens" Then
                intEndColumn = intColumn
            End If

            intColumn += 1
        Next
        
        Dim intRow As Integer = 0
        
        For Each rowExcel In dtCSV.Rows
            
	
              strItemID = ""
              strItemType = ""
              strCategory = ""
              strDisplayName = ""
              strParentMenuItem = ""
              strIngredient = ""
              strPortionSize = ""
              strParent = ""
              strCalories = ""
              strFat = ""
              strCarbs = ""
              strProtein = ""
              strCholesterol = ""
              strDietaryFiber = ""
              strSodium = ""
              strPercentCaloriesFromFat = ""
              strPercentCaloriesFromCarbs = ""
              strPrice = ""
              strOHPrice = ""
          	  strSmileyChoiceIcon = ""
			  strGreatForTakeoutIcon = ""
			  strEatnSmartIcon = ""
			  
              If Not (IsDBNull(rowExcel(0)) And IsDBNull(rowExcel(15))) Then

                If Not intRow = 0 Then   
                    intColumn = 0
               
                    
                    For i = 0 To intEndColumn
                    
                    
                    	strItem = ""
                        
                        If Not IsDBNull(rowExcel(i)) Then
                        	strItem = rowExcel(i)
                        End If
                        
                        Select Case i
                            Case 0 ' ID
                                strItemID = strItem
                                If Not IsNumeric(strItemID) Then
                                    strItemID = ""
                                End If
                            Case 1
                                strItemType = strItem
                            Case 2
                                strCategory = strItem
                            Case 3
                                strDisplayName = strItem
                            Case 4
                                strParentMenuItem = strItem
                            Case 5
                                strIngredient = strItem
                            Case 6
                                strCalories = strItem
                            Case 7
                                strFat = strItem 
                            Case 8
                                strCarbs = strItem 
                            Case 9
                                strProtein = strItem 
                            Case 10
                                strCholesterol = strItem 
                            Case 11
                                strDietaryFiber = strItem 
                            Case 12
                                strSodium = strItem 
                            Case 13
                                strPercentCaloriesFromFat = strItem 
                            Case 14
                                strPercentCaloriesFromCarbs = strItem
                                If IsNumeric(strPercentCaloriesFromCarbs) Then
                                    strPercentCaloriesFromCarbs = FormatNumber(strPercentCaloriesFromCarbs, 10)
                                End If
                                
                            Case 15
                                strMenuCategory = strItem
                            Case 16
                                strMenuDisplayName = strItem
                            Case 17
                                strActiveMenu = strItem
                            Case 18
                                strDescription = strItem
                            Case 19
                                strPrice = strItem
                            Case 20
                                strItemStartDate = strItem
                            Case 21
                                strItemEndDate = strItem
							Case 22
							    strSmileyChoiceIcon  = strItem
							Case 23
								strGreatForTakeoutIcon = strItem
							Case 24
								strEatnSmartIcon = strItem
                            Case 25
                                strAllergens = strItem
                        End Select
                    Next
                    
                    
                    'Response.Write("strItemID:" & strItemID & ":" & strMenuDisplayName)
                    'Response.End
                    
                    
                    If CStr(strItemID) <> "" Or CStr(strMenuDisplayName) <> "" Then
                        strSQL = "INSERT INTO tblNutritionalFile (ItemType, Category, DisplayName, ParentMenuItem, IngredientWholeItem, PortionSize, ParentIngredient, Calories, Fat, Carbs, Protein, Cholesterol, DietaryFiber, Sodium, CaloriesFat, CaloriesCarbs, MenuCategory, MenuDisplayName, ActiveMenu, ShowPrice, ShowNutritionals, Description, Price, OHPrice, ItemStartDate, ItemEndDate, Allergens, SmileyChoiceIcon, GreatForTakeoutIcon, EatnSmartIcon) VALUES (" & strProcessSQL(strItemType) & "," & strProcessSQL(strCategory) & "," & strProcessSQL(strDisplayName) & "," & strProcessSQL(strParentMenuItem) & "," & strProcessSQL(strIngredient) & "," & strProcessSQL(strPortionSize) & "," & strProcessSQL(strParent) & "," & strProcessSQL(strCalories) & "," & strProcessSQL(strFat) & "," & strProcessSQL(strCarbs) & "," & strProcessSQL(strProtein) & "," & strProcessSQL(strCholesterol) & "," & strProcessSQL(strDietaryFiber) & "," & strProcessSQL(strSodium) & "," & strProcessSQL(strPercentCaloriesFromFat) & "," & strProcessSQL(strPercentCaloriesFromCarbs) & ", " & strProcessSQL(strMenuCategory) & ", " & strProcessSQL(strMenuDisplayName) & ", " & strProcessSQL(strActiveMenu) & ", " &  strProcessSQL(strShowPrice)  & "," & strProcessSQL(strShowNutritionals) & "," & strProcessSQL(strDescription) & ", " & strProcessSQL(strPrice) & ", "  & strProcessSQL(strOHPrice) & "," & strProcessSQL(strItemStartDate) & ", " & strProcessSQL(strItemEndDate) & ", " & strProcessSQL(strAllergens) & ", " & strProcessSQL(strSmileyChoiceIcon) & "," & strProcessSQL(strGreatForTakeoutIcon) & "," & strProcessSQL(strEatnSmartIcon) & ")"
                    

                        db.RunOther(strSQL)
                    End If
                
                    ' Format SQL
                    intCounter = intCounter + 1
                   
                End If   
                bFound  = False
            End If
     
            intRow += 1                    
        Next 
	End If
	oExcel.Close()

    db.CloseConnection()

	
	Dim strMenuName, strMenuDescription, strStartDate, strEndDate
	Dim arrMenuCategory, intMenuCategoryCount
	Dim arrMenuSubCategory, intMenuSubCategoryCount
	Dim arrMenuSplit
	Dim arrMenu, intMenuCount
	
	Const MENU_CATEGORYID = 0
	Const MENU_CATEGORYNAME = 1
	Const MENU_SUBCATEGORYID = 2
	
	Dim lngMenuCategoryID
	Dim lngMenuSubCategoryID
	Dim strConnString As String
    Dim objConn 
	strConnString = ConfigurationSettings.AppSettings("connString")
	strConnString += "Provider=SQLOLEDB;OLE DB Services=-2;"
           
    
	objConn = Server.CreateObject("ADODB.Connection")
 	objConn.Open(strConnString)
        
	' Get Categories
	strSQL = "SELECT MenuCategoryID, MenuCategoryName FROM  tblMenuCategory"
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		arrMenuCategory = rstQuery.GetRows()
		intMenuCategoryCount = (UBound(arrMenuCategory, 2) + 1)	
	End If
	rstQuery.Close
	
	strSQL = "SELECT MenuCategoryID, MenuSubCategoryName, MenuSubCategoryID  FROM  tblMenuSubCategory"
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		arrMenuSubCategory = rstQuery.GetRows()
		intMenuSubCategoryCount = (UBound(arrMenuSubCategory, 2) + 1)	
	End If
	rstQuery.Close
	
	' Save All of the Menu Images
	strSQL = "SELECT MenuName, MenuImage, ParentMenuItem FROM tblMenu WHERE Not MenuImage IS NULL"
	intMenuCount = 0
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		arrMenu = rstQuery.GetRows()
		intMenuCount = (UBound(arrMenu, 2) + 1)	
	End If
	rstQuery.Close
	
	objConn.Close
	objConn = Nothing
	
    
	objConn = Server.CreateObject("ADODB.Connection")
 	objConn.Open(strConnString)
        
	Dim strMenuQuery
	Dim strSubMenuQuery, strMenuImage
	Dim intRecIndex
	Dim strAction
	Dim lngMenuID
    
    Dim rstQuery2
    
	' Backup Menu
	
	Dim intFieldColumnCount As Integer = 16
	Dim arrFieldList As String() = {"MenuID", "MenuName", "MenuDescription", "MenuImage", "MenuTitleImage", "StartDate", "EndDate", "Price", "OHPrice", "ShowNutritionals", "ShowPrice", "Allergens", "ParentMenuItem", "SmileyChoiceIcon", "GreatForTakeoutIcon", "EatnSmartIcon"}
	
	Dim arrSaveList(16) As String
	
	Dim strFieldList As String = ""
	Dim strValueList As String = ""
	Dim strDateTimeStamp As String = Now()
	Dim bSearchFound As Boolean 
	
	strSQL = "DELETE tblBackup_Menu WHERE DateTimeStamp < '" & DateAdd("m", -6, Now()) & "'"
	objConn.Execute(strSQL)
	
	strSQL = "SELECT * FROM tblMenu ORDER BY MenuID "
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		Do Until rstQuery.EOF
	
			strFieldList = ""
			strValueList = ""
			
	        ' Gather Values
			For i = 0 To intFieldColumnCount - 1
			
				arrSaveList(i) = ""
				
				If Not IsDBNull(rstQuery.Fields(arrFieldList(i)).value) Then
					arrSaveList(i) = rstQuery.Fields(arrFieldList(i)).value
				End If
				
				strFieldList += arrFieldList(i) & ","
				strValueList += strProcessSQL(arrSaveList(i)) & ","
			Next
			
			strFieldList = Left(strFieldList, Len(strFieldList) - 1)
			strValueList = Left(strValueList, Len(strValueList) - 1)
			
			strFieldList += ", DateTimeStamp"
			strValueList += ", '" & strDateTimeStamp & "'"
			
			strSQL = "INSERT INTO tblBackup_Menu (" & strFieldList & ") VALUES (" & strValueList & ")"
			objConn.Execute(strSQL)
			
			'Response.Write(strSQL & "<HR>")
			
			rstQuery.MoveNext
	    Loop
	End If
	rstQuery.Close
	
	
	'Response.Write("STOP")
	'Response.End()
	
	strSQL = "DELETE tblMenu"
	objConn.Execute(strSQL)
			
					
	strSQL = "SELECT * FROM  tblNutritionalFile WHERE (ActiveMenu = 'Y') "
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		Do Until rstQuery.EOF
	
			lngMenuID = 1
			strMenuName = rstQuery.Fields("MenuDisplayName").value
			strMenuDescription = rstQuery.Fields("Description").value
			strStartDate = rstQuery.Fields("ItemStartDate").value
			strEndDate = rstQuery.Fields("ItemEndDate").value
			strPrice = rstQuery.Fields("Price").value
			strOHPrice = rstQuery.Fields("OHPrice").value
			strShowNutritionals = "Y"
			strMenuImage = ""
			strSmileyChoiceIcon = 0
			strGreatForTakeoutIcon = 0
			strEatnSmartIcon = 0
			
            Try
                If rstQuery.Fields("SmileyChoiceIcon").value = "Y" Then
                    strSmileyChoiceIcon = 1
                End If
            Catch e As Exception
            End Try
            Try
                If rstQuery.Fields("GreatForTakeoutIcon").value = "Y" Then
                    strGreatForTakeoutIcon = 1
                End If
            Catch e As Exception
            End Try
            Try
                If rstQuery.Fields("EatnSmartIcon").value = "Y" Then
                    strEatnSmartIcon = 1
                End If
            Catch e As Exception
            End Try
            If Not IsDate(strStartDate) Then
                strStartDate = ""
            End If
            If Not IsDate(strEndDate) Then
                strEndDate = ""
            End If
			
            ' Find Name
            intRecIndex = 0
			'Response.Write("intMenuCount:" & intMenuCount)
			
			strParentMenuItem = rstQuery.Fields("ParentMenuItem").value
			
			bSearchFound = false
            If Not intMenuCount = 0 Then
                Do Until intRecIndex > UBound(arrMenu, 2)
					
                    'Response.Write(LCase(arrMenu(0, intRecIndex)) & ":" & Trim(LCase(strMenuName)) & "<BR>")
	
                    If Trim(LCase(arrMenu(0, intRecIndex))) = Trim(LCase(strMenuName)) Then
                        strMenuImage = arrMenu(1, intRecIndex)
						
						 bSearchFound = True
						

                    End If
					
                    intRecIndex = intRecIndex + 1
                Loop
				
				If Not  bFound Then
					Do Until intRecIndex > UBound(arrMenu, 2)
						If Trim(LCase(arrMenu(2, intRecIndex))) = Trim(LCase(strParentMenuItem)) Then
							strMenuImage = arrMenu(1, intRecIndex)
							bSearchFound = True
						End If
						intRecIndex = intRecIndex + 1
					Loop
				End If		
            End If		
			
            ' Response.Write("strShowNutritionals:" & strShowNutritionals & "<BR>")
			
            If strShowNutritionals <> "" And Not strShowNutritionals = "N" Then
                strShowNutritionals = "1"
            Else
                strShowNutritionals = "0"
            End If
            
            
            strShowPrice = "Y"
			
            If strShowPrice <> "" And Not strShowPrice = "N" Then
                strShowPrice = "1"
            Else
                strShowPrice = "0"
            End If
			
            strAllergens = rstQuery.Fields("Allergens").value
            strMenuCategory = rstQuery.Fields("MenuCategory").value
            
			
            strAction = "Add"
				

            If strAction = "Add" Then
                strSQL = "INSERT INTO tblMenu (MenuName, MenuDescription, StartDate, EndDate, Price, OHPrice, ShowNutritionals, ShowPrice, Allergens, ParentMenuItem, MenuImage, SmileyChoiceIcon, GreatForTakeoutIcon, EatnSmartIcon) "
                strSQL = strSQL & " VALUES (" & strProcessSQL(strMenuName) & "," & strProcessSQL(strMenuDescription) & "," & strProcessSQL(strStartDate) & "," & strProcessSQL(strEndDate) & "," & strProcessSQL(strPrice) & "," & strProcessSQL(strOHPrice) & "," & strShowNutritionals & "," & strShowPrice & "," & strProcessSQL(strAllergens) & "," & strProcessSQL(strParentMenuItem) & "," & strProcessSQL(strMenuImage) & "," & strProcessSQL(strSmileyChoiceIcon) & "," & strProcessSQL(strGreatForTakeoutIcon) & "," & strProcessSQL(strEatnSmartIcon) & ")"

            Else
                strSQL = "UPDATE tblMenu SET MenuName = " & strProcessSQL(strMenuName) & ", MenuDescription = " & strProcessSQL(strMenuDescription) & ", StartDate = " & strProcessSQL(strStartDate) & ", EndDate = " & strProcessSQL(strEndDate) & ", Price = " & strProcessSQL(strPrice) & ", OHPrice = " & strProcessSQL(strOHPrice) & ", ShowNutritionals = " & strShowNutritionals & ", ShowPrice = " & strShowPrice & ", Allergens = " & strProcessSQL(strAllergens) & ", ParentMenuItem = " & strProcessSQL(strParentMenuItem) & ", MenuImage = " & strProcessSQL(strMenuImage) & ", SmileyChoiceIcon = " & strProcessSQL(strSmileyChoiceIcon) & ", GreatForTakeoutIcon = " & strProcessSQL(strGreatForTakeoutIcon) & ", EatnSmartIcon = " & strProcessSQL(strEatnSmartIcon) & " WHERE MenuID = " & lngMenuID
			
            End If
			
            ' Determine if this is an Add or an Edit
            Response.Write(strSQL & "<BR>")
            objConn.Execute(strSQL)
			
            If strAction = "Add" Then
                strSQL = "SELECT * FROM tblMenu WHERE MenuName = " & strProcessSQL(strMenuName) & "  ORDER BY MenuID DESC"
                rstQuery2 = objConn.Execute(strSQL)
                If Not (rstQuery2.BOF And rstQuery2.EOF) Then
                    Do Until rstQuery2.EOF
                        lngMenuID = rstQuery2.Fields("MenuID").value
                        Exit Do
                        rstQuery2.MoveNext()
                    Loop
                End If
                rstQuery2.Close()
            End If
			
			
            strSQL = "DELETE tblMenuItem WHERE MenuID = " & lngMenuID
            objConn.Execute(strSQL)
					
            'Response.Write("lngMenuID:" & lngMenuID)
            'Response.Write("strMenuCategory:" & strMenuCategory & "<BR>")
			
			
            'If InStr(strMenuCategory, ",") = 0 And InStr(strMenuCategory, ".") > 0  Then
			
            strMenuCategory = Replace(strMenuCategory, ".", ",")
            'End If
			
            If strMenuCategory = "Celiac-Sides. Vegetarian-Sides" Then
                strMenuCategory = "Celiac-Sides, Vegetarian-Sides"
            End If
            If strMenuCategory = "Sides. Vegetarian-Sides" Then
                strMenuCategory = "Sides, Vegetarian-Sides"
            End If



            arrMenuSplit = Split(strMenuCategory, ",")
            For Each strItem In arrMenuSplit
                If strItem <> "" Then
                    ' Response.Write("X:" & strItem & "<BR>")
					
                    strMenuQuery = ""
                    strSubMenuQuery = ""
                    ' Look for Menu, SubMenu
                    If InStr(strItem, "-") > 0 Then
                        strMenuQuery = Left(strItem, InStr(strItem, "-") - 1)
                        strSubMenuQuery = Right(strItem, Len(strItem) - (InStr(strItem, "-")))
                    Else
                        strMenuQuery = strItem
                    End If
					
					
                    strMenuQuery = Trim(strMenuQuery)
                    strSubMenuQuery = Trim(strSubMenuQuery)
					
                    'Response.Write("Y:" & strMenuQuery & "<BR>")
                    'Response.Write("Y:" & strSubMenuQuery & "<BR>")
					
                    ' Find the Menu Items
                    intRecIndex = 0
                    lngMenuCategoryID = ""
                    Do Until intRecIndex > UBound(arrMenuCategory, 2)
                        If Trim(LCase(arrMenuCategory(MENU_CATEGORYNAME, intRecIndex))) = Trim(LCase(strMenuQuery)) Then
                            lngMenuCategoryID = CStr(arrMenuCategory(MENU_CATEGORYID, intRecIndex))
                            Exit Do
                        End If
                        intRecIndex = intRecIndex + 1
                    Loop
					
                    If Not IsNumeric(lngMenuCategoryID) Then ' = "" 
                        Response.Write("1 - Cannot find:" & strMenuQuery & "<BR>")
                        Response.End()
                    End If
					
                    intRecIndex = 0
                    lngMenuSubCategoryID = ""
					
                    If strSubMenuQuery = "Appetizer" And strMenuQuery = "Vegetarian" Then
                        strSubMenuQuery = "Appetizers"
                    End If
						
                    'If strSubMenuQuery = "Beverages" And strMenuQuery = "Fountain" Then
                    '	strSubMenuQuery = "Fountain (Free Refills)"
                    'End If				
						
                    If strSubMenuQuery <> "" Then
                        Do Until intRecIndex > UBound(arrMenuSubCategory, 2)
                            If Trim(LCase(arrMenuSubCategory(MENU_CATEGORYNAME, intRecIndex))) = Trim(LCase(strSubMenuQuery)) And CStr(arrMenuSubCategory(MENU_CATEGORYID, intRecIndex)) = CStr(lngMenuCategoryID) Then
                                lngMenuSubCategoryID = arrMenuSubCategory(MENU_SUBCATEGORYID, intRecIndex)
                                Exit Do
                            End If
                            intRecIndex = intRecIndex + 1
                        Loop
						
                        If Not IsNumeric(lngMenuSubCategoryID) Then ' lngMenuSubCategoryID = "" Then
                            Response.Write("2 - Cannot find:" & strSubMenuQuery & ":" & lngMenuSubCategoryID & "<BR>")
                            Response.End()
                        End If
                    End If
					
                    If Not IsNumeric(lngMenuSubCategoryID) Then '   = "" Then
                        lngMenuSubCategoryID = "0"
                    End If
					
                    strSQL = "INSERT INTO tblMenuItem (MenuID, MenuCategoryID, MenuSubCategoryID) VALUES (" & lngMenuID & "," & lngMenuCategoryID & "," & lngMenuSubCategoryID & ")"
                    objConn.Execute(strSQL)
					
                    ' Response.Write(strSQL & "<BR>")
					
					
                End If
            Next
			
		
            rstQuery.MoveNext()
        Loop
	End If
	rstQuery.Close
	
	
	strSQL = "DELETE FROM tblNutritionalIngredient"
	objConn.Execute(strSQL)
	
	strSQL = "DELETE FROM tblNutritionalCategory"
	objConn.Execute(strSQL)
	
	strSQL = "DELETE FROM tblNutritionalMenu"
	objConn.Execute(strSQL)

	
	
	' Find the Duplicate Menu Items
	Dim strDupMenuItems
	Dim intDupMenuCount
	Dim arrDupMenuItems
	Dim lngLastItem
	Dim intDupMenuIndex
	Dim strLastItemType
	
	strDupMenuItems = ""
	intDupMenuCount = 0
	
	
	
	
	strSQL = "SELECT COUNT(ID) AS CountMenuID, ParentMenuItem, IngredientWholeItem FROM tblNutritionalFile GROUP BY ParentMenuItem, IngredientWholeItem HAVING (IngredientWholeItem = 'Parent')"
	
	' Response.Write(strSQL)
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		Do Until rstQuery.EOF
			If rstQuery.Fields("CountMenuID").value >= 2 Then
				If intDupMenuCount > 0 Then
					strDupMenuItems = strDupMenuItems & ","
				End If
				strDupMenuItems = strDupMenuItems & rstQuery.Fields("ParentMenuItem").value
				intDupMenuCount = intDupMenuCount + 1
			End If
			rstQuery.MoveNext
		Loop
	
	End If
	rstQuery.Close
	

	arrDupMenuItems = Split(strDupMenuItems, ",")
	For Each strItem In arrDupMenuItems
		
		intDupMenuIndex = 1
		lngLastItem = -1
		strLastItemType = ""
		
		strSQL = "SELECT * FROM tblNutritionalFile WHERE ParentMenuItem = '" & strItem & "'"
		rstQuery = objConn.Execute(strSQL)
		If Not (rstQuery.BOF and rstQuery.EOF) Then
			Do Until rstQuery.EOF
				
				If Not lngLastItem = -1 AND NOT lngLastItem = (rstQuery.Fields("ID").value - 1) OR (NOT strLastItemType = "" And strLastItemType = "Ingredient" AND rstQuery.Fields("IngredientWholeItem").value = "Parent") Then
					intDupMenuIndex = intDupMenuIndex + 1
				End If
				
				strSQL  = "UPDATE tblNutritionalFile SET ParentMenuItem = '" & strItem & intDupMenuIndex & "' WHERE ID = " & rstQuery.Fields("ID").value
				objConn.Execute(strSQL)
				
				'Response.Write("strSQL:" & strSQL & "<BR>")
				
				lngLastItem = rstQuery.Fields("ID").value
				strLastItemType = rstQuery.Fields("IngredientWholeItem").value
				
				rstQuery.MoveNext
			Loop
		End If
		rstQuery.Close
		'Response.Write(strItem & "<BR>")
	Next
	
	
	
	' Insert Categories
	strSQL = "SELECT Category, MAX(ID) AS MaxID FROM tblNutritionalFile GROUP BY Category ORDER BY MaxID"
	rstQuery = objConn.Execute(strSQL)
	
	intCounter = 0
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		
		Do Until rstQuery.EOF
			
			strCategory = rstQuery.Fields("Category").value
			strSQL = "INSERT INTO tblNutritionalCategory (CategoryTitle, SortOrder) VALUES (" & strProcessSQL(strCategory) & "," & intCounter & ")"
			objConn.Execute(strSQL)
			intCounter = intCounter + 1
			
			rstQuery.MoveNext
		Loop
	
	End If
	rstQuery.Close
	

	' Get the Categories
	Const NUTRITIONAL_CATEGORYID = 0
	Const CATEGORY_TITLE = 1
	
	Dim arrCategory, intCategoryCount
	
	strSQL = "SELECT NutritionalCategoryID, CategoryTitle FROM tblNutritionalCategory ORDER BY SortOrder"
	rstQuery = objConn.Execute(strSQL)
	intCategoryCount = 0
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		arrCategory = rstQuery.GetRows()
		intCategoryCount = UBound(arrCategory, 2) + 1
	End If
	rstQuery.Close
	
	' Response.Write("intCategoryCount:" & intCategoryCount)
	
	' Insert the Menu Items
	Dim lngNutritionalCategoryID, intIndex
	' OLD strSQL = "SELECT * FROM tblNutritionalFile WHERE ParentIngredient = 'Parent' ORDER BY ID"
	strSQL = "SELECT * FROM tblNutritionalFile WHERE ItemType = 'I' ORDER BY ID"
	rstQuery = objConn.Execute(strSQL)
	


	If Not (rstQuery.BOF and rstQuery.EOF) Then
		
		Do Until rstQuery.EOF
			
			strItemType = ""
			strCategory = ""
			strDisplayName = ""
			strParentMenuItem = ""
			strIngredient = ""
			strPortionSize = ""
			strParent = ""
			strCalories = ""
			strFat = ""
			strCarbs = ""
			strProtein = ""
			strCholesterol = ""
			strDietaryFiber = ""
			strSodium = ""
			strPercentCaloriesFromFat = ""
			strPercentCaloriesFromCarbs = ""
			
			If Not IsDBNull(rstQuery.Fields("ItemType")) Then
				strItemType = rstQuery.Fields("ItemType").value 
			End If
			If Not IsDBNull(rstQuery.Fields("Category")) Then
				strCategory = rstQuery.Fields("Category").value
			End If
			
			If Not IsDBNull(rstQuery.Fields("DisplayName")) Then
				strDisplayName = rstQuery.Fields("DisplayName").value
			End If
			strParentMenuItem = rstQuery.Fields("ParentMenuItem").value
			
			strIngredient = rstQuery.Fields("IngredientWholeItem").value
			If Not IsDBNull(rstQuery.Fields("PortionSize")) Then
				strPortionSize = rstQuery.Fields("PortionSize").value
			End If
			If Not IsDBNull(rstQuery.Fields("ParentIngredient")) Then
				strParent = rstQuery.Fields("ParentIngredient").value
			End If
			If Not IsDBNull(rstQuery.Fields("Calories")) Then
				strCalories = rstQuery.Fields("Calories").value
			End If
			If Not IsDBNull(rstQuery.Fields("Fat")) Then
				strFat = rstQuery.Fields("Fat").value
			End If
			If Not IsDBNull(rstQuery.Fields("Carbs")) Then
				strCarbs = rstQuery.Fields("Carbs").value
			End If
			If Not IsDBNull(rstQuery.Fields("Protein")) Then
				strProtein = rstQuery.Fields("Protein").value
			End If
			If Not IsDBNull(rstQuery.Fields("Cholesterol")) Then
				strCholesterol = rstQuery.Fields("Cholesterol").value
			End If
			If Not IsDBNull(rstQuery.Fields("DietaryFiber")) Then
				strDietaryFiber = rstQuery.Fields("DietaryFiber").value
			End If
			If Not IsDBNull(rstQuery.Fields("Sodium")) Then
				strSodium = rstQuery.Fields("Sodium").value
			End If
			If Not IsDBNull(rstQuery.Fields("CaloriesFat")) Then
				strPercentCaloriesFromFat = rstQuery.Fields("CaloriesFat").value
			End If
			If Not IsDBNull(rstQuery.Fields("CaloriesCarbs")) Then
				strPercentCaloriesFromCarbs = rstQuery.Fields("CaloriesCarbs").value
			End If
			
			lngNutritionalCategoryID = ""
			If intCategoryCount > 0 Then
				For intIndex = 0 To UBound(arrCategory, 2) 
					If Not IsDBNull(arrCategory(CATEGORY_TITLE, intIndex)) Then
						If arrCategory(CATEGORY_TITLE, intIndex) = strCategory Then
							lngNutritionalCategoryID = arrCategory(NUTRITIONAL_CATEGORYID, intIndex)
						End If
					End If
				Next
			End If


		    strSQL = "INSERT INTO tblNutritionalMenu (NutritionalCategoryID, DisplayName, MenuName, ParentIngredient, Calories, Fat, Carbohydrates, Protein, Cholesterol, DietaryFiber, Sodium, PercentCaloriesFromFat, PercentCaloriesFromCarbs) VALUES (" & strProcessSQL(lngNutritionalCategoryID) & "," & strProcessSQL(strDisplayName) & "," & strProcessSQL(strParentMenuItem) & "," & strProcessSQL(strIngredient) & "," & strProcessSQL(strCalories) & "," & strProcessSQL(strFat) & "," & strProcessSQL(strCarbs) & "," & strProcessSQL(strProtein) & "," & strProcessSQL(strCholesterol) & "," & strProcessSQL(strDietaryFiber) & "," & strProcessSQL(strSodium) & "," & strProcessSQL(strPercentCaloriesFromFat) & "," & strProcessSQL(strPercentCaloriesFromCarbs) & ")"
			
		

			
			objConn.Execute(strSQL)
			
			rstQuery.MoveNext
			
			' Exit Do
		Loop
	End If
	rstQuery.Close
	
	

	' Get the Menu List
	Const MENU_ID = 0
	Const MENU_NAME = 1
	
	Dim lngNutritionalMenuID
	
	strSQL = "SELECT MenuID, MenuName FROM tblNutritionalMenu ORDER BY MenuID"
	rstQuery = objConn.Execute(strSQL)
	intMenuCount = 0
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		arrMenu = rstQuery.GetRows()
		intMenuCount = UBound(arrMenu, 2) + 1
	End If
	rstQuery.Close
	
	
	' Get the Ingredients
	' strSQL = "SELECT * FROM tblNutritionalFile WHERE ParentIngredient = 'Ingredient' ORDER BY ID"
	strSQL = "SELECT * FROM tblNutritionalFile WHERE ItemType <> 'I' ORDER BY ID"
	rstQuery = objConn.Execute(strSQL)
	If Not (rstQuery.BOF and rstQuery.EOF) Then
		
		Do Until rstQuery.EOF
		
			strItemType = ""
			strCategory = ""
			strDisplayName = ""
			strParentMenuItem = ""
			strIngredient = ""
			strPortionSize = ""
			strParent = ""
			strCalories = ""
			strFat = ""
			strCarbs = ""
			strProtein = ""
			strCholesterol = ""
			strDietaryFiber = ""
			strSodium = ""
			strPercentCaloriesFromFat = ""
			strPercentCaloriesFromCarbs = ""
			
			If Not IsDBNull(rstQuery.Fields("ItemType")) Then
				strItemType = rstQuery.Fields("ItemType").value
			End If
			strCategory = rstQuery.Fields("Category").value
			If Not IsDBNull(rstQuery.Fields("DisplayName")) Then
				strDisplayName = rstQuery.Fields("DisplayName").value
				strIngredient = rstQuery.Fields("DisplayName").value
			Else	
				strIngredient = rstQuery.Fields("IngredientWholeItem").value
			End If
			strParentMenuItem = rstQuery.Fields("ParentMenuItem").value
			
			If Not IsDBNull(rstQuery.Fields("PortionSize")) Then
				strPortionSize = rstQuery.Fields("PortionSize").value
			End If
			If Not IsDBNull(rstQuery.Fields("ParentIngredient")) Then
				strParent = rstQuery.Fields("ParentIngredient").value
			End If
			If Not IsDBNull(rstQuery.Fields("Calories")) Then
				strCalories = rstQuery.Fields("Calories").value
			End If
			If Not IsDBNull(rstQuery.Fields("Fat")) Then
				strFat = rstQuery.Fields("Fat").value
			End If
			If Not IsDBNull(rstQuery.Fields("Carbs")) Then
				strCarbs = rstQuery.Fields("Carbs").value
			End If
			If Not IsDBNull(rstQuery.Fields("Protein")) Then
				strProtein = rstQuery.Fields("Protein").value
			End If
			If Not IsDBNull(rstQuery.Fields("Cholesterol")) Then
				strCholesterol = rstQuery.Fields("Cholesterol").value
			End If
			If Not IsDBNull(rstQuery.Fields("DietaryFiber")) Then
				strDietaryFiber = rstQuery.Fields("DietaryFiber").value
			End If
			If Not IsDBNull(rstQuery.Fields("Sodium")) Then
				strSodium = rstQuery.Fields("Sodium").value
			End If
			If Not IsDBNull(rstQuery.Fields("CaloriesFat")) Then
				strPercentCaloriesFromFat = rstQuery.Fields("CaloriesFat").value
			End If
			If Not IsDBNull(rstQuery.Fields("CaloriesCarbs")) Then
				strPercentCaloriesFromCarbs = rstQuery.Fields("CaloriesCarbs").value
			End If
			
			lngNutritionalMenuID = ""
			If intMenuCount > 0 Then
				For intIndex = 0 To UBound(arrMenu, 2) 
					If arrMenu(MENU_NAME, intIndex) = strParentMenuItem Then
						lngNutritionalMenuID = arrMenu(MENU_ID, intIndex)
					End If
				Next
			End If
			
			If IsNumeric(lngNutritionalMenuID) Then

				strSQL = "INSERT INTO tblNutritionalIngredient (NutritionalMenuID, IngredientName, IngredientType, PortionSize , Calories, Fat, Carbohydrates, Protein, Cholesterol, DietaryFiber, Sodium, PercentCaloriesFromFat, PercentCaloriesFromCarbs) VALUES (" & strProcessSQL(lngNutritionalMenuID) & "," & strProcessSQL(strIngredient) & "," & strProcessSQL(strItemType) & "," & strProcessSQL(strPortionSize) & "," & strProcessSQL(strCalories) & "," & strProcessSQL(strFat) & "," & strProcessSQL(strCarbs) & "," & strProcessSQL(strProtein) & "," & strProcessSQL(strCholesterol) & "," & strProcessSQL(strDietaryFiber) & "," & strProcessSQL(strSodium) & "," & strProcessSQL(strPercentCaloriesFromFat) & "," & strProcessSQL(strPercentCaloriesFromCarbs) & ")"	
				Response.Write(strSQL)
				objConn.Execute(strSQL)
				
			End If
			
			
			rstQuery.MoveNext
		Loop
	End If
	rstQuery.Close
	
	' Get the Totals
	objConn.Close
    objConn = Nothing
	

	
	
	Response.Redirect("NutritionalFileUploadAction2.aspx")
    
%>TEST
