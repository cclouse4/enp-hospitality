<script runat="server">
	
	dim objConn As Object, strConnectionString As String, s_ConnectionSemaphore As String, s_ConnectionSemaphore2 As String, s_ConnectionSemaphore3 As String
	dim oConn As System.Data.SQLClient.SQLConnection
	dim oConn2 As System.Data.SQLClient.SQLConnection
	dim oConn3 As System.Data.SQLClient.SQLConnection
	dim oLiveConn As System.Data.SQLClient.SQLConnection
	
	sub OpenLiveConnection()
			s_ConnectionSemaphore = 1
			oLiveConn = New System.Data.SQLClient.SQLConnection(ConfigurationSettings.AppSettings("connStringLive"))
			oLiveConn.Open()	
	end sub
	
	sub CloseLiveConnection()
			if s_ConnectionSemaphore = 1 then
				s_ConnectionSemaphore = 0
				oLiveConn.Close()
				oLiveConn = nothing
			end if
	end sub
	
	sub OpenConnection()
			s_ConnectionSemaphore = 1
			oConn = New System.Data.SQLClient.SQLConnection(ConfigurationSettings.AppSettings("connString"))
			oConn.Open()	
	end sub
	
	sub CloseConnection()
			if s_ConnectionSemaphore = 1 then
				s_ConnectionSemaphore = 0
				oConn.Close()
				oConn = nothing
			end if
	end sub
	
	sub OpenOrigConnection()
			s_ConnectionSemaphore2 = 1
			oConn2 = New System.Data.SQLClient.SQLConnection(ConfigurationSettings.AppSettings("connStringOrig"))
			oConn2.Open()	
	end sub
	
	sub CloseOrigConnection()
			if s_ConnectionSemaphore2 = 1 then
				s_ConnectionSemaphore2 = 0
				oConn2.Close()
				oConn2 = nothing
			end if
	end sub
	
	sub OpenOrigConnection2()
			s_ConnectionSemaphore3 = 1
			oConn3 = New System.Data.SQLClient.SQLConnection(ConfigurationSettings.AppSettings("connStringOrig"))
			oConn3.Open()	
	end sub
	
	sub CloseOrigConnection2()
			if s_ConnectionSemaphore3 = 1 then
				s_ConnectionSemaphore3 = 0
				oConn3.Close()
				oConn3 = nothing
			end if
	end sub
	
	function formatBoolean(ByVal p_Value) As Boolean
		if p_Value = "on" then
			formatBoolean = 1
		else
			formatBoolean = 0
		end if
	end function
	
	Function FixSQL(sql)
		dim strTempSQL
		strTempSQL = Replace(sql,"'","''")
		FixSQL = strTempSQL
	End Function
	
</script>
