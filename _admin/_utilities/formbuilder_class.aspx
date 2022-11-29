<script runat="server">
' form builder class -- creates form elements on the fly for detail modules
Public Class FormBuilder
	private FormElements(1) As String
	public useEditor As Boolean = false
	private leftElementWidth As Integer = 300
	private rightElementWidth As Integer = 700
	
	public function NewFormElement(Optional ByVal caption As String = "", Optional ByVal formType As String = "", Optional ByVal elName As String = "", Optional ByVal elValue As Object = Nothing, Optional ByVal iSize As Integer = 0, Optional ByVal bState As Boolean = false, Optional ByVal sString As String = "")
		Dim sReturn As String
		
		if len(sString) > 0 and formType <> "hidden" and InStr(LCase(caption),"date") = 0 then
			Dim _c() As String = split(caption,":")
			if UBound(_c) > 0 then
				sReturn += "<tr><th class='endCap' colspan='2'>Selected "& _c(0) &": "& sString &"</th></tr>"
			else	
				sReturn += "<tr><th class='endCap' colspan='2'>Selected "& caption &": "& sString &"</th></tr>"
			end if
		end if
	
		if formType <> "hidden" then
			sReturn += "<tr><th class='endCap' width='"& leftElementWidth &"'>"& caption &"</th><th class='formInput' align='left' width='"& rightElementWidth &"'>"
		end if
		
		if formType = "select" then
			Dim i As Integer
			sReturn += "<select id='"& elName &"' name='"& elName &"'><option value='-1'>- Select One -</option>"
			if isArray(elValue) then
				for i = 1 to UBound(elValue) + 1
					sReturn += "<option value='"& i &"'"
					if i = iSize then
						sReturn += "selected=selected"
					end if
					sReturn += ">"& elValue(i-1) &"</option>"
				next i
			else
				sReturn += CStr(elValue)
			end if
			sReturn += "</select>"
		else if LCase(formType) = "cell" then
			sReturn = "<tr><th class='endCap' colspan='2'>"& caption &"</th></tr>"
		else if formType = "multiCheckbox" then
			'sReturn += ""
			sReturn += "<div style='height:200px; overflow: auto; border: 1px solid #000'>"& elValue &"</div>"
		else if formType = "textarea" then
			useEditor = true
			sReturn = "<tr><th class='endCap' width='"& leftElementWidth &"'>"& caption &"</th><th class='formInput' align='left' width='"& rightElementWidth &"'><textarea name='"& elName &"' id='"& elName &"' cols=""50"" rows=""15"">"& elValue &"</textarea></th></tr>"
		else if formType = "plain" then
			sReturn = "<tr><th class='endCap' width='"& leftElementWidth &"'>"& caption &"</th><th class='formInput' align='left' width='"& rightElementWidth &"'>"& elValue &"</th></tr>"
		else
			sReturn += "<input type='"& formType &"' id='"& elName &"' name='"& elName &"'"		
			if not elValue = Nothing then
				sReturn += " value="""& elValue &""""
			end if
			if iSize > 0 then
				sReturn += " size='"& iSize &"'"
			end if
			if bState then
				if formType = "checkbox" then
					sReturn += " CHECKED "
				else if formType = "select" then
					sReturn += " selected=selected "
				end if
			end if
			sReturn += "/>"
		end if
		
		if InStr(LCase(caption),"date") > 0 and len(sString) > 0 then
			sReturn += sString
		end if
		
		if formType <> "hidden" then
			sReturn += "</th></tr>"
		end if
		
		' push it to the array
		ReDim Preserve FormElements(UBound(FormElements) + 1)
		FormElements(UBound(FormElements)) = sReturn
	end function
	
	Public Function PrintFormItems() As String
		Dim i As Integer
		Dim sReturn As String
		
		for i = 0 to UBound(FormElements)
			sReturn += FormElements(i)
		next i
		return sReturn
	End Function
	
	Public Function PrintSubmitButton(sText) As String
		return New String("<tr><th class='endCap'></th><th class='formInput'><input type='submit' value='"& sText &"' name='submit' id='submit'/></th></tr>")
	End Function
end Class
</script>