<%
' use exception handling outside of the case in order to close any open database connections
Dim s As String
	try
		select case sAction	
			case "Add"
				
				Dim column As String
				Dim columnValue As String
				if HAS_SORT_ORDER then
					
					db.RunQuery("select max(SortOrder) as Sort from "& sTableName &" "& SORT_WHERE)

					while db.InResults()
						
						if db.GetItem("Sort") <> "" then
							
							for i = 0 to UBound(sColumnNames)
								if sColumnNames(i) = "SortOrder"then
									sValues(i) = db.GetItem("Sort") + 1
								end if
							next i
						else
							
							for i = 0 to UBound(sColumnNames)
								
								
								if sColumnNames(i) = "SortOrder"then
									
									sValues(i) = 1

								end if
								
							next i
							
						end if
						
														
						
					end while
				end if
				
				
				
				s = "insert into "& sTableName &" ("
				for each column In sColumnNames
					s += column &","
				next column

				s = left(s, len(s) - 1)
				s += ") values ("
				for each columnValue In sValues
					s += columnValue &","
				next columnValue
				s = left(s, len(s) - 1)
				s += ");"
				
				
				db.RunOther(s)
	
				s = "select * from "& sTableName &" where "
				for i = 0 to UBound(sColumnNames)
					if i >= 1 then
						exit for
					end if
					s += sColumnNames(i) &" = "& sValues(i) &" and "
				next i
				s = Left(s,len(s)-4)
				db.RunQuery(s)
				while db.InResults()
					iID = db.GetItem(sPrimaryKey)
				end while
	
			case "Edit"
				Dim sField As String
				Dim bFlag As Boolean = false
				Dim bNormalField As Boolean = true
				s = "update "& sTableName &" set "
				
				for i = 0 to UBound(sColumnNames)
					bFlag = false
					bNormalField = true
					if lcase(sColumnNames(i)) <> "sortorder" then
						for each sField In sFileFieldName
							if InStr(sField,sColumnNames(i)) > 0 then ' if the file field = the field selected
								bNormalField = false
								bFlag = true
								exit for
							else
								bNormalField = true
								bFlag = false
							end if
						next
						if bFlag then
							if (sValues(i) <> "''" and len(sValues(i)) > 0 and sValues(i) <> "NULL") and (bFlag or bNormalField) and sColumnNames(i) <> "SortOrder" then
								s += sColumnNames(i) &" = "& sValues(i) &","
							end if
						else if bNormalField then
							s += sColumnNames(i) &" = "& sValues(i) &","
						end if
					end if
				next i
				
				s = left(s, len(s) - 1)
				s += " where "& sPrimaryKey &" = "& iID
				
				db.RunOther(s)
				response.write(s)
			case "Delete"
				db.RunOther("delete from "& sTableName &" where "& sPrimaryKey &" = "& iID)
			case "Sort"
			if HAS_SORT_ORDER then
				if iSortDirection = 0 then
					db.RunQuery("select "& sPrimaryKey &", SortOrder from "& sTableName & SORT_WHERE &" order by SortOrder asc")
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
					db.RunQuery("select "& sPrimaryKey &", SortOrder from "& sTableName & SORT_WHERE &" order by SortOrder desc")
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
			end if
			case "RemoveFile"
			if HAS_FILES then
				db.RunOther("update "& sTableName &" set "& sPostFieldName &" = '' where "& sPrimaryKey &" = "& iID)
				db.CloseConnection()				
				Response.Redirect(sDetailURL &"?Action=Edit&"& sPrimaryKey &"="& iID)
			end if
		end select
		' process the bridge queries here
		Response.Write("id: "& iID)
		sBridgeQueries = Replace(sBridgeQueries,"#id#",iID)
		Dim sBridgeSplit() As String = split(sBridgeQueries,"|")
		Dim sBridgeItem As String
		for each sBridgeItem In sBridgeSplit
			if len(sBridgeItem) > 0 then
				Response.Write(sBridgeItem)
				db.RunOther(sBridgeItem)
			end if
		next
	catch e As Exception
		Response.Write("There was an error processing this request: "& e.Message())
		if DEBUG then
			Response.Write("<br/>Query: "& s) ' used for outputting the query that caused the error. use only for debugging
		end if
	end try
	'finally 
		db.CloseConnection()
		if not DEBUG then
			Response.Redirect(sRedirectURL)
		end if
	'end try	
%>