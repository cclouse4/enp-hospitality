<%	
	Dim strClientName As String = ConfigurationSettings.AppSettings("applicationName")
	Dim strClientUploadDirectory As String = ""
		
	if Session("LoggedIn") <> 1 then
		Response.Redirect("index.aspx")
	end if

%>