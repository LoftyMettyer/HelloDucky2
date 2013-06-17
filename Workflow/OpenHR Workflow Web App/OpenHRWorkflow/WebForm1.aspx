<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="WebForm1.aspx.vb" Inherits="OpenHRWorkflow.WebForm1" %>

<!DOCTYPE html>
<html>
    <head runat="server">
        <title></title>
        <style>
            body { font: 12px "Lucida Sans Unicode", "​Lucida Grande", ​ arial, ​ helvetica, ​ sans-serif; color: #515967; }
            ul { list-style: none outside none; margin: 0; padding: 0; }

            .tabstrip ul li {
                float: left;
                position: relative;
                background-color: #F3F3F4;
                border: 1px solid #DBDBDE;
                border-radius: 4px 4px 0 0;
                border-bottom-width: 0px;
                margin-right: -1px;      
            }
            .tabstrip ul li a {
                display: inline-block;
                border-radius: 4px 4px 0 0;
                padding: 3px 11px;
                outline: 0 none;
                color: #515967;
                text-decoration: none;
            }
            .tabstrip li.active {
                border: 1px solid #0879C0;
                border-bottom-width: 0;
                background-color: #FFFFFF;
                margin-bottom: -1px;
                padding-bottom: 1px;
                z-index: 1;
            }
            .tabstrip div {
                clear: both;
                border: 1px solid #0879C0;
                width: 350px;
                height: 175px;
            }
        </style>
    </head>
    <body>
        <form id="form1" runat="server">
            <div class='tabstrip'>
                <ul>
                    <li class="active"><a href="#">Paris</a></li>
                    <li><a href="#">New York</a></li>
                    <li><a href="#">London</a></li>
                    <li><a href="#">Moscow</a></li>
                    <li><a href="#">Sydney</a></li>
                </ul>
                <div></div>
                <div></div>
                <div></div>
                <div></div>
                <div></div>
            </div>
        </form>
    </body>
</html>