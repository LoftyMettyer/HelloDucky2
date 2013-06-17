<%@ Page Language="VB" AutoEventWireup="false" CodeFile="TabletLogin.aspx.vb" Inherits="TabletLogin" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    
    <style type="text/css">
      /* Remove margins from the 'html' and 'body' tags, and ensure the page takes up full screen height */
      html, body {height:100%; margin:0; padding:0;}
      /* Set the position and dimensions of the background image. */
      #pagebackground {position:fixed; top:0; left:0; width:100%; height:100%;}
      /* Specify the position and layering for the content that needs to appear in front of the background image. Must have a higher z-index value than the background image. Also add some padding to compensate for removing the margin from the 'html' and 'body' tags. */
      #content {position:relative; z-index:1; padding:10px;}
    </style>

</head>
<body>
    <form id="form1" runat="server">
 	    <div id="pagebackground" runat="server" style="width:100%;height:100%"></div>

      <div id="content" style="height: 450px; width: 320px; margin: 0px auto; text-align: center; border-radius: 20px; padding: 10px;">
        <iframe style="border-radius: 20px;filter: alpha(opacity=90); -moz-opacity: 0.9; opacity: 0.9" src="MobileLogin.aspx" height="100%" width="100%" frameborder="0">
                   
        </iframe> 
      </div>

    </form>
</body>
</html>




