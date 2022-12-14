<%@ Page Language="C#" AutoEventWireup="true" %>

<%@ Import Namespace="System.IO" %>
<%@ Import Namespace="System.Collections" %>
<%@ Import Namespace="System.Collections.Generic" %>

<script runat="server">

    private string path = HttpContext.Current.Server.MapPath("../../../images");
    private string imgPath = "../../../images";



    protected void Page_Load(object sender, EventArgs e)
    {
        if (!string.IsNullOrEmpty(Request.QueryString["del"]))
        {
            fileDel(Request.QueryString["del"]);
        }
        else if (!string.IsNullOrEmpty(Request.QueryString["fldel"]))
        {
            folderDel(Request.QueryString["fldel"]);
        }

        pageBind();
    }

    #region Function

    /// <summary>
    /// Page Load - page object bind
    /// </summary>
    private void pageBind()
    {
        lnkHome.NavigateUrl = "up.aspx";

        DirectoryInfo drInfo = new DirectoryInfo(path);

        // Main Folder
        bindFolder(drInfo);

        if (!string.IsNullOrEmpty(Request.QueryString["fl"]))
        {
            path += "\\" + Request.QueryString["fl"];

            imgPath += "/" + Request.QueryString["fl"];

            // Path Change
            drInfo = new DirectoryInfo(path);

            // Create Folder Panel  !? Next Step 
            pnlCreators.Visible = false;
        }

        // File List
        bindFile(drInfo);
    }

    /// <summary>
    /// Folder List
    /// </summary>
    /// <param name="drInfo"></param>
    private void bindFolder(DirectoryInfo drInfo)
    {
        DirectoryInfo[] drList = drInfo.GetDirectories();

        List<DirectoryInfo> lstDirectory = new List<DirectoryInfo>();

        foreach (DirectoryInfo drc in drList)
        {
            lstDirectory.Add(drc);
        }

        lvFolderList.DataSource = lstDirectory;
        lvFolderList.DataBind();
    }

    /// <summary>
    /// File List
    /// </summary>
    /// <param name="drInfo"></param>
    private void bindFile(DirectoryInfo drInfo)
    {
        FileInfo[] rgFiles = drInfo.GetFiles();

        List<FileInfo> lstFile = new List<FileInfo>();

        foreach (FileInfo fi in rgFiles)
        {
            lstFile.Add(fi);
        }

        lvFileList.DataSource = lstFile;
        lvFileList.DataBind();

    }

    #region File Delete

    private void fileDel(string fileName)
    {
        if (!string.IsNullOrEmpty(Request.QueryString["fl"]))
        {
            File.Delete(path + "\\" + Request.QueryString["fl"] + "\\" + fileName);
        }
        else
        {
            File.Delete(path + "\\" + fileName);
        }
    }

    #endregion

    #region Folder Delete

    private void folderDel(string folder)
    {
        Directory.Delete(path + "\\" + folder, true);

        Response.Redirect("up.aspx");
    }

    #endregion

    #endregion

    #region Event

    protected void btnCreateFolder_click(object sender, EventArgs e)
    {
        if (!Directory.Exists(path + "\\" + txtFolder.Text.Trim()))
            Directory.CreateDirectory(path + "\\" + txtFolder.Text.Trim());

        pageBind();
    }

    protected void btnAddFile_click(object sender, EventArgs e)
    {
        try
        {
            flUpload.SaveAs(path + "\\" + flUpload.FileName);


            if (!string.IsNullOrEmpty(Request.QueryString["fl"]))
            {
                Response.Redirect("up.aspx?fl=" + Request.QueryString["fl"]);
            }
            else
            {
                Response.Redirect("up.aspx");
            }
        }
        catch (Exception ex)
        { }
    }


    protected void lvFileList_ItemDataBound(object sender, ListViewItemEventArgs e)
    {

        if (e.Item.ItemType == ListViewItemType.DataItem)
        {
            ListViewDataItem dataItem = (ListViewDataItem)e.Item;
            Image imgList = (Image)e.Item.FindControl("imgList");

            imgList.ImageUrl = imgPath + "/" + ((FileInfo)dataItem.DataItem).Name;
        }
    }


    #endregion
    
    
</script>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>netImageBrowser</title>
    <link href="css/browser.css" rel="stylesheet" type="text/css" />
    <link href="css/jquery.contextMenu.css" rel="stylesheet" type="text/css" />

    <script type="text/javascript" src="../../tiny_mce_popup.js"></script>

    <script src="js/netbrowser.js" type="text/javascript"></script>

    <script src="http://ajax.googleapis.com/ajax/libs/jquery/1.3.2/jquery.js" type="text/javascript"></script>

    <script src="js/jquery.contextMenu.js" type="text/javascript"></script>

    <script src="js/jquery.jqURL.js" type="text/javascript"></script>

    <script type="text/javascript">
			
			$(document).ready( function() {
				
				// Show menu when File is clicked
				$('a[rel*=jmenu]').contextMenu({
					menu: 'myMenu'
				},
					function(action, el, pos) {
					
				    var queryString = $.jqURL.get("fl");
				   
				   if(action == "paste"){
				   
				    var path =  "images/";
				    var queryString = $.jqURL.get("fl");
				    
				    if(queryString != undefined ||  queryString != null){
				        path += queryString+"/";
				    }
				    
				   	javascript:netBrowserDialog.insert(path +$(el).attr('id'),$(el).attr('id'));
				   }// if end
				   else if(action == "delete"){
				   
				    if(queryString != null && queryString != undefined){
                            window.location= "up.aspx?fl="+queryString+"&del="+$(el).attr('id');
                        }
                        else{
                            window.location= "up.aspx?del="+$(el).attr('id');
                        }
				   }// else if end
				   
				});// contextMenu end
				
				// Show menu when Folder is clicked
				$('a[rel*=fmenu]').contextMenu({
					menu: 'myFolderMenu'
					
				},
					function(action, el, pos) {
				    
				    if(action == "delete"){
                        window.location= "up.aspx?fldel="+$(el).attr('id');
				   }//  if end
				   
				});// contextMenu end
				
			});// documentReady end
			
    </script>

</head>
<body style="background: url(img/fading_background_19.png) repeat;">
    <form id="form1" runat="server">
    <div>
        <table>
            <!-- Main Table -->
            <tr>
                <td style="height: 300px; width: 160px; padding-left: 5px;" valign="top">
                    <div class="Menu">
                        FOLDER</div>
                    <asp:HyperLink runat="server" ID="lnkHome" BorderStyle="None" CssClass="Home">Home</asp:HyperLink>
                    <ul class="folderListUl">
                        <asp:ListView runat="server" ID="lvFolderList">
                            <LayoutTemplate>
                                <div id="itemPlaceholder" runat="server">
                                </div>
                            </LayoutTemplate>
                            <ItemTemplate>
                                <li><a href='up.aspx?fl=<%# Eval("NAME") %>' rel="fmenu" id='<%# Eval("NAME") %>'>
                                    <%# Eval("NAME") %></a><br />
                                </li>
                            </ItemTemplate>
                        </asp:ListView>
                    </ul>
                    <asp:Panel runat="server" ID="pnlCreators">
                        <div class="Menu">
                            CREATE FOLDER</div>
                        <asp:TextBox ID="txtFolder" runat="server"></asp:TextBox><br />
                        <br />
                        <asp:Button runat="server" ID="btnCreateFolder" Text="Create Folder" OnClick="btnCreateFolder_click" />
                    </asp:Panel>
                    <div class="Menu">
                        FILE UPLOAD</div>
                    <asp:FileUpload ID="flUpload" runat="server" /><br />
                    <br />
                    <asp:Button runat="server" ID="btnAddFile" Text="Add File" OnClick="btnAddFile_click" />
                </td>
                <td style="height: 300px; width: 450px; padding-left: 10px;" valign="top">
                    <br />
                    <asp:ListView runat="server" ID="lvFileList" OnItemDataBound="lvFileList_ItemDataBound">
                        <LayoutTemplate>
                            <div id="itemPlaceholder" runat="server">
                            </div>
                        </LayoutTemplate>
                        <ItemTemplate>
                            <a href="#" id='<%# Eval("NAME") %>' rel="jmenu">
                                <asp:Image runat="server" Width="68" Height="90" ID="imgList" />
                            </a>
                        </ItemTemplate>
                    </asp:ListView>
                </td>
            </tr>
        </table>
    </div>
    <ul id="myMenu" class="contextMenu">
        <li class="paste"><a href="#paste">Select</a></li>
        <li class="delete"><a href="#delete">Delete</a></li>
    </ul>
    <ul id="myFolderMenu" class="contextMenu">
        <li class="delete"><a href="#delete">Delete</a></li>
    </ul>
    </form>
</body>
</html>
