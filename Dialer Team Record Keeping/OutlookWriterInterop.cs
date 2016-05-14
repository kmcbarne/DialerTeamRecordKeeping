using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using TextBox = System.Windows.Controls.TextBox;
using Exception = System.Exception;
using System.Collections;

namespace Dialer_Team_Record_Keeping
{
    class OutlookWriterInterop
    {
        Microsoft.Office.Interop.Outlook.Application outlookApp;

        public OutlookWriterInterop()
        {
            outlookApp = new Microsoft.Office.Interop.Outlook.Application();
        }

        public void CreateIntradayUpdate(string messageBody)
        {
            MailItem outlookMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

            outlookMail.To = "Keven.McBarnes@td.com";
            outlookMail.Subject = "List Penetration Table";
            outlookMail.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
            outlookMail.HTMLBody = messageBody;
            outlookMail.Display();
        }

        public void CreateWrapEmail(List<WrapReport> report)
        {
            MailItem outlookMail = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);
            string wrapHighlight = "";
            string rowColor = "<td bgcolor='#FFFFFF'>&nbsp;";

            string msgBody = "<TABLE BORDER='4' CELLSPACING='0' CELLPADDING='0' style='border-collapse: collapse; text-align:center; width: 900px bordercolor='#111111'>" +
                             "<FONT COLOR='000000' face='ARIAL' size='2'><b><TR HEIGHT=10 >" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 100px'>&nbsp;TID</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 275px'>&nbsp;Name</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 52px'>&nbsp;Calls</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Login</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Active</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Not Ready</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Idle</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Wrap</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Hold</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 85px'>&nbsp;Hold Count</TD>" +
                             "<TD BGCOLOR='#e26b0a' style='text-align:center; width: 75px'>&nbsp;Avg Wrap</FONT></b></TD></TR>";

            foreach (WrapReport agent in report)
            {
                if (agent.AvgWrap.TotalSeconds >= 20)
                {
                    wrapHighlight = "<td bgcolor='#da9694'>&nbsp;";
                }

                msgBody += "<tr><FONT COLOR='000000' face='Arial' size='2'>" +
                    "<td bgcolor='#FFFFFF' align='center'>" + agent.Tid + "</td>" +
                    "<td bgcolor='#FFFFFF' align='center'>" + agent.Name + "</td>" +
                    rowColor + agent.Calls.ToString() + "</td>" +
                    rowColor + agent.Login.ToString() + "</td>" +
                    rowColor + agent.Active.ToString() + "</td>" +
                    rowColor + agent.NotReady.ToString() + "</td>" +
                    rowColor + agent.Idle.ToString() + "</td>" +
                    rowColor + agent.Wrap.ToString() + "</td>" +
                    rowColor + agent.Hold.ToString() + "</td>" +
                    rowColor + agent.HoldCount.ToString() + "</td>" +
                    wrapHighlight + agent.AvgWrap.ToString() + "</td>" +
                    "</FONT></tr>";

                wrapHighlight = "<td bgcolor='#ffffff'>&nbsp;";
            }

            //			If Not ((rs.Fields("TID - Name") = 0) Or (rs.Fields("TID - Name") = "Report Total")) Then
            //
            //            If rs.Fields("Avg Wrap") >= #12:00:20 AM# Then
            //                wrapHighlight = "<td bgcolor='#da9694'>&nbsp;"
            //            End If
            //
            //	        esTable = esTable & "<tr><FONT COLOR='000000' face='Arial' size='2'>" & _
            //	            "<td bgcolor='#FFFFFF' align='left'>&nbsp;" & rs.Fields("TID - Name") & "</td>" & _
            //	            rowColor & rs.Fields("Calls") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Login"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Active"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Not Ready"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Idle"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Wrap"), "h:mm:ss") & "</td>" & _
            //	            rowColor & Format(rs.Fields("Hold"), "h:mm:ss") & "</td>" & _
            //	            rowColor & rs.Fields("Hold Count") & "</td>" & _
            //	            wrapHighlight & Format(rs.Fields("Avg Wrap"), "h:mm:ss") & "</td>" & _
            //	            "</FONT></tr>"
            //
            //        End If

            outlookMail.To = "Keven.McBarnes@td.com";
            outlookMail.Subject = "Wrap Report";
            outlookMail.BodyFormat = OlBodyFormat.olFormatHTML;
            outlookMail.HTMLBody = msgBody;
            outlookMail.Display();
        }

        //		Microsoft.Office.Interop.Outlook.MailItem olkMail1 =
        //    (MailItem)olkApp1.CreateItem(OlItemType.olMailItem);
        //        olkMail1.To = txtpsnum.Text;
        //        olkMail1.CC = "";
        //        olkMail1.Subject = "Assignment note";
        //        olkMail1.Body = "Assignment note";
        //        olkMail1.Attachments.Add(AssignNoteFilePath, 
        //            Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, 1, 
        //                "Assignment_note");
        //olkMail1.Save();
        //olkMail.Send();
    }
}
