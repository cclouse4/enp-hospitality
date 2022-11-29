<script runat="server">
Dim SHOWLOGO As Boolean = true

class Homepage
	public metaKeywords As String
	public metaDescription As String
	public metaTitle As String
	public image As String
	public header1 As String
	public header2 As String
	public caption As String
	public url As String
	public urltext As String
	public tout(3) As Tout
	public pageTitle As String
	
	public sub New()
		metaKeywords = ""
		metaDescription = ""
		metaTitle = ""
		image = ""
		header1 = ""
		header2 = ""
		caption = ""
		url = ""
		urltext = ""
	end sub

end class

public class LandingTout
	public id As Integer
	public name As String
	public image As String
	public text As String
	public url As String
	
	public sub New()
	end sub
	
	public shared function getToutsByContentId(ByRef db As Chemistry.DBObject, ByRef id As Integer) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblTout a inner join brgContentTout b on a.id = b.toutId where b.contentId = "& id &" order by SortOrder asc")
		while db.InResults()
			Dim _r As New LandingTout
			_r.id = db.GetItem("id")
			_r.name = db.GetItem("name")
			_r.image = db.GetItem("image")
			_r.url = db.GetItem("url")
			_r.text = db.GetItem("text")
			r.Add(_r)
		end while
		return r
	end function
	
end class

class Tout
	public img As String
	public title As String
	public text As String
	public url As String
	public urlText As String
	
	public sub New(i As String, _t As String, t As String, u As String, ut As String)
		if len(i) > 0 then
			img = i
		else 
			img = ""
		end if
		
		if len(_t) > 0 then
			title = _t
		else
			title = ""
		end if
		
		if len(t) > 0 then
			text = t
		else
			text = ""
		end if
		
		if len(u) > 0 then
			url = u
		else
			url = ""
		end if
		
		if len(ut) > 0 then
			urltext = ut
		else
			urltext = ""
		end if
	end sub
	
end class

class Venue
	public db As Chemistry.DBObject
	public id As Integer
	public name As String
	public text As String
	public image As String
	public quickFacts As String
	public landmarks() As String
	public exclusive As Boolean
	public preferred As Boolean
	public url As String
	
	public sub New(ByRef d As Chemistry.DBObject)
		db = d
	end sub
	public sub New()
	end sub
	
	public sub getLandmarks()
		if id > 0 then
			Dim l As New System.Collections.ArrayList
			db.RunSimpleQuery("select a.* from tblLandmarks a inner join brgVenueLandmarks b on a.id = b.landmarkId where b.venueId = "& id &" order by a.SortOrder asc")
			if db.HasRows() then
				while db.ReadSimpleQuery().Read()
					l.Add(db.GetSimpleItem("landmark"))
				end while
			end if
			db.EndSimpleQuery()
			landmarks = l.toArray(Type.GetType("System.String"))
		end if
	end sub
	
	public shared function fillAll(ByVal _d As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		_d.runQuery("select * from tblVenues order by SortOrder asc")
		while _d.InResults()
			Dim v As New Venue(_d)
			v.id = _d.GetItem("id")
			v.name = _d.GetItem("name")
			v.text = _d.GetItem("text")
			v.image = _d.GetItem("img")
			v.quickFacts = _d.GetItem("quickfacts")
			if len(_d.GetItem("exclusive")) > 0 then
				v.exclusive = _d.GetItem("exclusive")
			end if
			if len(_d.GetItem("preferred")) > 0 then
				v.preferred = _d.GetItem("preferred")
			end if
			v.url = _d.GetItem("url")
			v.getLandmarks()
			r.Add(v)
		end while
		return r
	end function
	
	public shared function getSelectOptions(ByRef d As Chemistry.DBObject) As String
		Dim r As String
		d.OpenConnection()
		d.RunQuery("select * from tblVenues order by SortOrder asc")
		while d.InResults()
			Dim v As New Venue(d)
			v.id = d.GetItem("id")
			v.name = d.GetItem("name")
			r += "<option value="""& v.name &""">"& v.name &"</option>"
		end while
		d.CloseConnection()
		return r
	end function
	
	public function hasDetails() As Boolean
		if len(quickFacts) > 0 or UBound(landmarks) > 0 then
			return true
		end if
		return false
	end function

end class

class People 
	public db As Chemistry.DBobject
	public id As Integer
	public image As String
	public name(2) As String
	public title As String
	public contentTitle As String
	public contentText As String
	public email As String
	public phone As String
	public valid As Boolean
	public type As String
	public typeName() As String = { "Chefs","Director","Staff" }
	
	public sub New()
	end sub
	
	public sub New(ByRef d As Chemistry.DBObject)
		db = d
	end sub
	
	public shared function getSlideshow(ByRef d As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r as New System.Collections.ArrayList
		d.RunQuery("select * from tblPeople order by SortOrder asc, lname asc, fname asc")
		while d.InResults()
			Dim p As New People()
			p.id = d.GetItem("id")
			p.image = d.GetItem("image")
			'p.setName(d.GetItem("fname"),d.GetItem("lname"))
			'p.title = d.GetItem("title")
			'p.contentTitle = d.GetItem("contentTitle")
			'p.contentText = d.GetItem("content")
			'p.email = d.GetItem("email")
			'p.phone = d.GetItem("phone")
			r.Add(p)
		end while
		return r		
	end function
	
	public sub setName(ByRef a As String, ByRef b As String)
		name(0) = a
		name(1) = b
	end sub
	
	public function getById(ByVal _id As Integer) As People
		db.RunQuery("select * from tblPeople where id = "& _id)
		Dim p As New People()
		while db.InResults()			
			p.id = db.GetItem("id")
			p.image = db.GetItem("image")
			p.setName(db.GetItem("fname"),db.GetItem("lname"))
			p.title = db.GetItem("title")
			p.contentTitle = db.GetItem("contentTitle")
			p.contentText = db.GetItem("content")
			p.email = db.GetItem("email")
			p.phone = db.GetItem("phone")
			p.type = db.GetItem("type")
			p.valid = true
		end while
		return p
	
	end function
	
	public shared function getByType(ByRef d As Chemistry.DBObject, ByVal typeId As Integer) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		d.RunQuery("select * from tblPeople where type = "& typeId &" order by SortOrder asc")
		while d.InResults()
			Dim p As New People
			p.id = d.GetItem("id")
			p.image = d.GetItem("image")
			p.setName(d.GetItem("fname"),d.GetItem("lname"))
			p.title = d.GetItem("title")
			p.contentTitle = d.GetItem("contentTitle")
			p.contentText = d.GetItem("content")
			p.email = d.GetItem("email")
			p.phone = d.GetItem("phone")
			p.valid = true	
			r.Add(p)
		end while
		return r
	end function
	
	public function isValid() As Boolean
		if valid then
			return true
		else
			return false
		end if
	end function
	
	public function printName() As String
		return name(0) &" "& name(1)
	end function

end class

class Page 
	public db As Chemistry.DBObject
	public id As Integer
	public parentId As Integer
	public url As String
	public name As String
	public metaKeywords As String
	public metaDescription As String
	public metaTitle As String
	public pageTitle As String
	public backgroundImage As String
	public image As String
	public introText As String
	public mainText As String
	public sidebarTitle As String
	public sidebarContent As String
	public landingTouts As System.Collections.ArrayList
	public isLandingPage As Boolean
	public tout1img As String
	public tout2img As String
	public tout3img As String
	
	public sub New() 
	end sub
	
	public sub New(ByRef d As Chemistry.DBObject)
		db = d
	end sub
	
	public shared function getHomepageContent(ByRef db As Chemistry.DBObject) As Page
		Dim p As New Page
		db.RunQuery("select * from tblHomepage")
		while db.InResults()
			p.name = "About Us"
			p.mainText = db.GetItem("caption")
		end while
		return p
	end function
	
	public shared function getById(ByRef db As Chemistry.DBObject, ByRef _id As Integer, ByRef _iid As Integer) As Page
		Dim p As New Page
		db.RunQuery("select * from tblContent where id = "& _id)
		while db.InResults()
			p.id = db.GetItem("id")
			p.parentId = db.GetItem("parentId")
			p.url = db.GetItem("pageUrl")
			p.name = db.GetItem("name")
			p.metaTitle = db.GetItem("metaTitle")
			p.metaKeywords = db.GetItem("metaKeywords")
			p.metaDescription = db.GetItem("metaDescription")
			p.pageTitle = db.GetItem("pageTitle")
			p.backgroundImage = db.GetItem("backgroundImage")
			p.image = db.GetItem("image")
			p.introText = db.GetItem("introText")
			p.mainText = db.GetItem("mainText")
			p.sidebarTitle = db.GetItem("sidebarTitle")
			p.sidebarContent = db.GetItem("sidebarContent")
			p.isLandingPage = db.GetItem("landingPage")
			p.tout1img = db.GetItem("tout1img")
			p.tout2img = db.GetItem("tout2img")
			p.tout3img = db.GetItem("tout3img")
		end while
		p.landingTouts = LandingTout.getToutsByContentId(db,p.id)
		return p
	end function
	
	public shared function getSubMenu(ByRef db As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblContent order by parentId asc, SortOrder asc")
		while db.InResults()
			Dim p As New Page
			p.id = db.GetItem("id")
			p.parentId = db.GetItem("parentId")
			p.url = db.GetItem("pageUrl")
			p.name = db.GetItem("name")
			p.metaTitle = db.GetItem("metaTitle")
			p.metaKeywords = db.GetItem("metaKeywords")
			p.metaDescription = db.GetItem("metaDescription")
			p.pageTitle = db.GetItem("pageTitle")
			p.backgroundImage = db.GetItem("backgroundImage")
			p.image = db.GetItem("image")
			p.introText = db.GetItem("introText")
			p.mainText = db.GetItem("mainText")
			p.sidebarTitle = db.GetItem("sidebarTitle")
			p.sidebarContent = db.GetItem("sidebarContent")
			r.Add(p)
		end while
		return r
	end function
	
	public shared function isSelected(ByRef id As Integer) As Boolean
		if CInt(System.Web.HttpContext.Current.Request.QueryString("pid")) = id or CInt(System.Web.HttpContext.Current.Request.QueryString("ppid")) = id then
			return true
		end if
		return false	
	end function
end class

class TimelineEntry 
	public db As Chemistry.DBObject
	public id As Integer
	public year As Integer
	public name As String
	public mainText As String

	
	public sub New() 
	end sub
	
	public sub New(ByRef d As Chemistry.DBObject)
		db = d
	end sub
	
	
	public shared function getById(ByRef db As Chemistry.DBObject, ByRef _id As Integer, ByRef _iid As Integer) As TimelineEntry
		Dim p As New TimelineEntry
		db.RunQuery("select * from tblTimeline where id = "& _id)
		while db.InResults()
			p.id = db.GetItem("id")
			p.year = db.GetItem("year")
			p.name = db.GetItem("name")
			p.mainText = db.GetItem("mainText")
		end while
		return p
	end function
	
	public shared function fillAll(ByVal _d As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		_d.runQuery("select * from tblTimeline order by year asc, SortOrder asc")
		while _d.InResults()
			Dim t As New TimelineEntry(_d)
			t.id = _d.GetItem("id")
			t.name = _d.GetItem("name")
			t.year = _d.GetItem("year")
			t.mainText = _d.GetItem("mainText")
			r.Add(t)
		end while
		return r
	end function
	

end class

class Help
	public id As Integer
	public name As String
	public url As String
	
	public sub New()
	end sub
	
	public shared function getList(ByRef db As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblHelpBar order by SortOrder asc")
		while db.InResults()
			Dim h As New Help
			h.id = db.GetItem("id")
			h.name = db.GetItem("name")
			h.url = db.GetItem("url")
			r.Add(h)
		end while
		return r
	end function
	
end class

class Gallery
	public id As Integer
	public type As Integer
	public image As String
	public caption As String
	public thumbnail As String
	
	public sub New()
		image = ""
		caption = ""
	end sub
	
	public shared function getById(ByRef db As Chemistry.DBObject, ByRef id As Integer) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblGallery where type = "& id &" order by SortOrder asc")
		while db.InResults()
			Dim g As New Gallery
			g.id = db.GetItem("id")
			g.type = db.GetItem("type")
			g.image = db.GetItem("image")
			g.caption = db.GetItem("caption")
			g.thumbnail = db.GetItem("thumbnail")
			r.Add(g)
		end while
		return r
	end function
end class

class Fader
	public image As String
	
	public sub New()
	end sub
	
	public shared function getAll(ByRef db As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblFader order by SortOrder asc")
		while db.InResults()
			Dim f As New Fader
			f.image = db.GetItem("filename")
			r.Add(f)
		end while
		return r		
	end function
end class

class Glossary
	public name As String
	public text As String
	
	public sub New()
	end sub
	
	public shared function getAll(ByRef db As Chemistry.DBObject) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		db.RunQuery("select * from tblGlossary order by cast(title as nvarchar(250)) asc")
		while db.InResults()
			Dim f As New Glossary
			f.name = db.GetItem("title")
			f.text = db.GetItem("description")
			r.Add(f)
		end while
		return r		
	end function
	
	public shared function getByLetter(ByRef db As Chemistry.DBObject,ByRef l As String) As System.Collections.ArrayList
		Dim r As New System.Collections.ArrayList
		if l = "ALL" then
			return Glossary.getAll(db)
		else
			db.RunQuery("select * from tblGlossary where cast(title as nvarchar(250)) like '"& LCase(l) &"%' or cast(title as nvarchar(250)) like '"& UCase(l) &"%' order by cast(title as nvarchar(250)) asc")
			while db.InResults()
				Dim f As New Glossary
				f.name = db.GetItem("title")
				f.text = db.GetItem("description")
				r.Add(f)
			end while
			return r		
		end if
	end function
end class
</script>