<script runat="server">

' database class
Public Class DBObject
	' DOCUMENTATION - how to use this class
	'
	' Dim db As New DBObject								- Create object (NOTE: This opens the connection and assigns the connection variable to the SQL command object)
	' db.RunQuery("select table_field from table_name")		- Run SQL Query	(NOTE: This sets the command text and assigns the query results to the SQL data reader object)
	' while db.Read()										- Loop through results (NOTE: This checks if there are rows returned. No need to call HasRows())
	' 	variable = db.GetItem("table_field", "data_type")	- Assign value to variable (NOTE: This checks for null results)
	' end while												- End the loop
	' db.EndQuery()											- Closes the data reader (NOTE: Connection remains open until it is closed at the bottom of the page
	'
	'
	' private variables
	Public sql As String
	Dim com As New System.Data.SqlClient.SqlCommand
	Dim reader As System.Data.SqlClient.SqlDataReader	
	Dim conn As System.Data.SQlClient.SqlConnection
	Dim qryRan As Boolean
	Dim isOpen As Boolean
	
	' constructor
	Public Sub New()
		'OpenConnection()
	End Sub
	
	' destructor
	Protected Overrides Sub Finalize()
		'CloseConnection()
	End Sub
	
	' functions
	Public Sub OpenConnection()
		conn = New System.Data.SQLClient.SQLConnection(ConfigurationSettings.AppSettings("ConnString"))
		conn.Open()
		com.connection = conn
		isOpen = true
	End Sub
	
	Public Sub CloseConnection()
		'if conn.State = "Open" then
		if isOpen then
			isOpen = false
			conn.Close()
			conn = nothing
		end if	
		'end if
	End Sub
	
	Public Sub InitProcedure(ByVal oObj As Object, ByVal sCommandType As String)
		com.CommandType = oObj
		com.CommandText = sCommandType	
	End Sub
	
	Public Function AddParam(ByVal oObj As Object, ByVal oType As Object, ByVal dAmt As Double)
		return com.Parameters.Add(oObj, oType, dAmt)	
	End Function
	
	Public Sub ExecuteSQL(ByVal sqlStr As String)
		
		sql = sqlStr
		com.commandText = sql
		com.ExecuteNonQuery()
		qryRan = true

	End Sub
	
	Public Sub RunQuery(ByVal sqlStr As String)
		
		sql = sqlStr
		com.commandText = sql
		reader = com.ExecuteReader()
		qryRan = true

	End Sub
	
	Public Function GetCom()
		return com
	End Function
	
	Public Sub RunOther(ByVal sqlStr As String)
		sql = sqlStr
		com.commandText = sql
		com.ExecuteNonQuery()
	End Sub
	
	Public Sub EndQuery()
		if qryRan then
			reader.Close
			qryRan = false
		end if
	End Sub
	
	Public Function Read()
		if qryRan then
			if reader.HasRows then
				Return reader.Read()
			end if
		end if
	End Function
	
	Public Function GetItem(sItem)
		if not IsDBNull(reader.Item(sItem)) then
			Return reader.Item(sItem)
		else
			Return nothing
		end if
	End Function
	
	Public Function GetItemPlainText(sItem)
		if not IsDBNull(reader.Item(sItem)) then
			Dim sValue As String = reader.Item(sItem)
			sValue = StripHTML(sValue)
			Return sValue
		else
			Return nothing
		end if
	End Function
	
	Function StripHTML(strText)
		Dim re As System.Text.RegularExpressions.Regex
		strText = re.Replace(strText, "<[^>]*>", "")
		Return strText
		
	End Function
	
End Class

</script>