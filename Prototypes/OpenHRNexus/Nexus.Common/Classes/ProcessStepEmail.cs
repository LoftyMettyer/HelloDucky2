using Nexus.Common.Interfaces;
using OpenHRNexus.Common.Enums;
using Nexus.Common.Enums;

namespace Nexus.Common.Classes
{
    public class ProcessStepEmail : IProcessStep
    {
        public int Id { get; set; }

        public ProcessElementType Type
        {
            get
            {
                return ProcessElementType.Email;
            }
        }

        public ProcessStepStatus Validate()
        {
            return ProcessStepStatus.Success;
        }

        public string To => "nick.gibson@advancedcomputersoftware.com";

        public string Message =>
            "<!DOCTYPE html>" +
            "<html lang='en'>" +
            "    <head>" +
            "        <meta charset='utf-8' />" +
            "    </head>" +
            "    <body>" +
            "        <p>" +
            "            <span style='color: #0094ff'>{0}</span> has requested a <span style='color:#0094ff'>{1}</span> holiday absence from <span style='color:#0094ff'>{2}</span> to <span style='color:#0094ff'>{3}.</span>" +
            "        </p>" +
            "        <p>" +
            "            Reason for absence: <span style='color: #0094ff'>{4}</span>" +
            "        </p>" +
            "        <p>" +
            "            Employee notes: <span style='color: #0094ff'>{5}</span>" +
            "        </p>" +
            "        <p>" +
            "            You can quickly approve or decline this absence request using the buttons below." +
            "        </p>" +
            "<div>" +
            "<!--[if mso]>" +
            "<style type='text/css'>" +
            ".bold {{font-weight: bold}}" +
            "</style>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='{6}' style='height:33px;v-text-anchor:middle;width:77px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#5CB85C'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      Approve" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:77px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#D9534F'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      Decline" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:130px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#5BC0DE'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      View the request" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <v:roundrect xmlns:v='urn:schemas-microsoft-com:vml' xmlns:w='urn:schemas-microsoft-com:office:word' href='http://www.EXAMPLE.com/' style='height:33px;v-text-anchor:middle;width:149px;margin-right: 5px;' arcsize='10%' stroke='f' fillcolor='#337AB7'>" +
            "    <w:anchorlock/>" +
            "    <center style='color:#ffffff;font-family:sans-serif;font-size:14px;font-weight:normal;'>" +
            "      View team calendar" +
            "    </center>" +
            "  </v:roundrect>" +
            "  <![endif]-->" +
            //"  <![if !mso]>" +
            //"  <table cellspacing='0' cellpadding='0'> <tr> " +
            //"  <td align='center' width='300' height='40' bgcolor='#d62828' style='-webkit-border-radius: 5px; -moz-border-radius: 5px; border-radius: 5px; color: #ffffff; display: block;'>" +
            //"    <a href='http://www.EXAMPLE.com/' style='font-size:16px; font-weight: bold; font-family:sans-serif; text-decoration: none; line-height:40px; width:100%; display:inline-block'>" +
            //"    <span style='color: #ffffff;'>" +
            //"      Button Text Here!" +
            //"    </span>" +
            //"    </a>" +
            //"  </td> " +
            //"  </tr> </table> " +
            //"  <![endif]>" +
            "  <!--[if !mso]>" +
            "        <span style='background: green; padding: 5px'><a style='text-decoration: none; color: white' href='{6}'>Approve</a></span>" +
            "        <span style='background: red; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>Decline</a></span>" +
            "        <span style='background: lightblue; padding: 5px'><a style='text-decoration: none; color: white' href='{7}'>View the request</a></span>" +
            "        <span style='background: blue; padding: 5px'><a style='text-decoration: none; color: white' href='{9}'>View team calendar</a></span>" +
            "  <![endif]-->" +
            "</div>" +
            "    </body>" +
            "</html>";

        public string Subject => "Nexus subject";
    }
}
