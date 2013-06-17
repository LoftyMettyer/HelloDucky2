<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="OpenHRWorkflow.WebForm1" %>

<!DOCTYPE html>
<html>
    <head id="Head1" runat="server">
        <title></title> 
        <style>
            /* reset */
            ul { list-style: none outside none; margin: 0; padding: 0; }
 
            /* defaults */
            body { font: 12px "Lucida Sans Unicode", helvetica, ​ sans-serif; color: #515967; }
            a { outline: 0 none; color: #515967; text-decoration: none; }
 
            /* tabstrip */
            
            .tabstrip > ul > li {
                float: left;
                border: 1px solid #DBDBDE;      /* border applied to all sides except bottom */
                border-bottom-width: 0px;
                margin-right: -1px;             /* pull the tab to the right in, to get rid of the double vertial border */
                border-radius: 4px 4px 0 0;
                background: url('Images/highlight.png') repeat-x 0 center #F3F3F4;
                background-position: 0 center;
                background-repeat: repeat-x;
            }
  
            .tabstrip ul > li > a {
                display: inline-block;          /* so vertical padding will take effect */
                padding: 6px 11px;
                border-radius: 4px 4px 0 0;     /* need to apply to this as well as li to stop this bleeding over li's' border radius */
            }           

            .tabstrip > ul > li:hover { background-color: #DBDBDE; }
            
            .tabstrip > ul > li.active {
                border: 1px solid #0879C0;
                border-bottom-width: 0;
                background-color: #FFFFFF;
                padding-bottom: 1px;            /* make the active tab 1px higher then use negative margin to pull the tab page up 1px */
                margin-bottom: -1px;            /* the tab page top border disappears behind the active tab, but not the others because it has not crossed them */       
                z-index: 1;                     /* z-index ensures the active tab comes in front of the others */
                position: relative;             /* postion required for z-index to take effect on floated elements */
            }
 
            .tabstrip > div {
                clear: both;                    /* need to clear the floats from the li's */
                border: 1px solid #0879C0;      /* border applied to all sides of the tab page */
                width: 350px;
                height: 175px;
            }
             
        </style>
    </head>
    <body>
        <form id="form1" runat="server">
            <div class='tabstrip'>
                <ul>
                    <li ><a href="#">Paris</a></li>
                    <li class="active"><a href="#">New York</a></li>
                    <li><a href="#">London</a></li>
                    <li><a href="#">Moscow</a></li>
                    <li><a href="#">Sydney</a></li>
                </ul>
                <div></div>
            </div>
        </form>
    </body>
</html>