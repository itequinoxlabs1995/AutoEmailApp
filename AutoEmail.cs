using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AutoEmailApp
{
    partial class AutoEmail : ServiceBase
    {

        public AutoEmail()
        {
            InitializeComponent();
            OnStart(null);

        }

        protected override void OnStart(string[] args)
        {
            //this.WriteToFile("Simple Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            //this.WriteToFile("Simple Service stopped {0}");
             this.ScheduleService();
        }


        public void ScheduleService()
        {
            try

            {
                SendEmailAsync();
                
            }
            catch (Exception ex)
            {
               // WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("AutoEmailApp"))
                {
                    serviceController.Stop();
                }
            }

        }

        public dynamic GetEmailName()
        {
           // var toEmail = (dynamic)null;
           // var username = (dynamic)null;
           var UsernamesEmails = (dynamic)null;

            try
            {
                // AutoEmail();
                DataTable dt = new DataTable();
                string query = "Select T1.User_Name, T1.Email_ID from D_User T1 join M_Role T2 On T1.RID = T2.RID  Where T2.Role_Name ='Manager';";
               // string query = "Select User_Name, Email_ID from D_User Where UID='1';";
                string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            UsernamesEmails = dt;

                        }
                    }
                }
               /* foreach (DataRow row in dt.Rows)
                {
                    username = row["User_Name"].ToString();
                    toEmail = row["Email_ID"].ToString();
                   //WriteToFile("Trying to send email to: " + name + " " + email);

                }
                */
            }
            catch (Exception ex)
            {
                //WriteToFile("Simple Service Error on: {0} " + ex.Message + ex.StackTrace);

               
            }
            return UsernamesEmails;
        }

        public dynamic GetSummaryData()
        {
            var SummaryData = (dynamic)null;
            try
            {
                // AutoEmail();
                DataTable dt = new DataTable();
                string query = "SELECT ROW_NUMBER()OVER (ORDER BY Project_Description) As \"S.No\",T3.Project_Description AS \"Project Code\",T4.Gate_Location AS \"Gate Location\", SUM(CASE WHEN T2.Status_Name = 'Completed' THEN 1 ELSE 0 END) AS Completed,SUM(CASE WHEN T2.Status_Name = 'Pending' THEN 1 ELSE 0 END) AS Pending,COUNT(T1.SID) AS Total FROM T_Gate_Pass T1 JOIN M_Status T2 ON T1.SID = T2.SID JOIN M_Project T3 ON T1.PID = T3.PID JOIN M_Gate T4 ON T3.GID = T4.GID WHERE T1.CreatedOn >= DATEADD(day, DATEDIFF(day, 1, GETDATE()), 0)GROUP BY T3.Project_Description,T4.Gate_Location;";
                //string query = "SELECT ROW_NUMBER()OVER (ORDER BY Project_Description) As \"S.No\",T3.Project_Description AS \"Project Code\",T4.Gate_No AS \"Gate No\", SUM(CASE WHEN T2.Status_Name = 'Completed' THEN 1 ELSE 0 END) AS Completed,SUM(CASE WHEN T2.Status_Name = 'Pending' THEN 1 ELSE 0 END) AS Pending,COUNT(T1.SID) AS Total FROM T_Gate_Pass T1 JOIN M_Status T2 ON T1.SID = T2.SID JOIN M_Project T3 ON T1.PID = T3.PID JOIN M_Gate T4 ON T3.GID = T4.GID WHERE T1.CreatedOn >= DATEADD(day, DATEDIFF(day, 1, GETDATE()), 0)GROUP BY T3.Project_Description,T4.Gate_No;";
               // string query = "SELECT \r\nROW_NUMBER() OVER (ORDER BY  PO_Number) As \"S.No\",\r\n    T1.PO_Number AS \"Project Code\",\r\n    T1.Gate_ID AS \"Gate No\", \r\n    SUM(CASE WHEN T2.Status_Name = 'Completed' THEN 1 ELSE 0 END) AS Completed\r\n\t,SUM(CASE WHEN T2.Status_Name = 'Pending' THEN 1 ELSE 0 END) AS Pending\r\n\t,COUNT(T1.SID) AS Total\r\nFROM \r\n    T_Gate_Pass T1 \r\nJOIN \r\n    M_Status T2 ON T1.SID = T2.SID\r\nWHERE \r\n    T1.CreatedOn >= DATEADD(day, DATEDIFF(day, 1, GETDATE()), 0)\r\nGROUP BY \r\n    T1.PO_Number, \r\n    T1.Gate_ID";
                // string query = "Select Distinct(T2.Status_Name) As Status\r\n,Count(T1.SID)As Status_Count from T_Gate_Pass\r\n T1 join M_Status T2 on T1.SID = T2.SID \r\n Where T1.CreatedOn >= dateadd(day,datediff(day,1,GETDATE()),0)\r\n group by T1.SID, T2.Status_Name;";
                string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;
                using (SqlConnection con = new SqlConnection(constr))
                {
                    using (SqlCommand cmd = new SqlCommand(query))
                    {
                        cmd.Connection = con;
                        using (SqlDataAdapter sda = new SqlDataAdapter(cmd))
                        {
                            sda.Fill(dt);
                            SummaryData = dt;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                ex.Message.ToString();
            }
            return SummaryData;
        }
   
        public async Task SendEmailAsync()
        {
            var SummaryData = GetSummaryData();
            var NamesEmails = GetEmailName();

            var Names = new List<string>();
            var toEmails = new List<string>();

            DataTable dt = SummaryData;
            DataTable dt1 = NamesEmails;

           for (int i = 0; i < dt1.Rows.Count; i++)
             {            
               
                var Names1 = dt1.Rows[i]["User_Name"];
                var toEmail1 = dt1.Rows[i]["Email_ID"];

                Names.Add(Names1.ToString());
                toEmails.Add(toEmail1.ToString());
            }

           
            string html = "<table style='width:100%;border:1px solid black'>";
            //add header row
            html += "<tr style=' background-color: #476084;\r\n  color: white;'>";
            for (int i = 0; i < dt.Columns.Count; i++)
                html += "<td style='border:1px solid black'>" + dt.Columns[i].ColumnName + "</td>";
            html += "</tr>";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr style='background-color: #f2f2f2;'>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td style='border:1px solid black'>" + dt.Rows[i][j].ToString() + "</td>";
                html += "</tr>";
            }
            html += "</table>";


            string Htmltable = "<p>Please find below the gate entry summary for "+ DateTime.Today.Date.AddDays(-1).ToString("dd/MM/yyyy") + " <br/>"+html + "</p><p>Best regards,<br/>IT Team</p>";

            string subject, htmlBody = null;
            htmlBody = Htmltable;
            // htmlBody = "<p>Dear \" + name + \",</p><p>Please find below the gate entry summary for \"+DateTime.Now.Date+\":\\r\\n\\r\\n<br/><table style='width:100%;border:1px solid black'><tr style=' background-color: #476084;\\r\\n  color: white;'><th style='border:1px solid black'>Project Code</th><th style='border:1px solid black'>Gate No</th><th style='border:1px solid black'>Completed</th>\r\n<th style='border:1px solid black'>Pending</th>\r\n<th style='border:1px solid black'>Total</th>\r\n</tr>\r\n<tr style='background-color: #f2f2f2;'><td style='border:1px solid black'>\" + status1 + \"</td><td style='border:1px solid black'>\" + status_count1 + \"</td>\r\n<td style='border:1px solid black'>\" + status_count1 + \"</td>\r\n<td style='border:1px solid black'>\" + status_count1 + \"</td>\r\n<td style='border:1px solid black'>\" + status_count1 + \"</td></tr>\r\n</table></p><p>Best regards,<br/>IT Team</p>";
            subject = "Daily Gate Entry Summary – "+ DateTime.Today.Date.AddDays(-1).ToString("dd/MM/yyyy");


            if (string.IsNullOrEmpty(toEmails.ToString()))
            {
                Console.WriteLine("Name not found.");
                return;
            }

            /* var emailData = new
             {
                 to = new List<string> { "monapag07@gmail.com" },
                 cc = new List<string> { "monapag07@gmail.com" },
                 bcc = new List<string> { "monapag07@gmail.com" },
                 subject = subject,
                 html = htmlBody
             };
            */


            for (int i = 0; i < dt1.Rows.Count; i++)
            {


                var emailData = new
                {
                    to = new List<string> { dt1.Rows[i]["Email_ID"].ToString() },
                    cc = new List<string> { ConfigurationManager.AppSettings["CcEmailID"], ConfigurationManager.AppSettings["CcEmailID1"] },
                    bcc = new List<string> { ConfigurationManager.AppSettings["BccEmailID"], ConfigurationManager.AppSettings["BccEmailID1"] },
                    subject = subject,
                    html = "<p>Dear " + dt1.Rows[i]["User_Name"].ToString() + ",</p>" + htmlBody
                };

                try
                {
                    HttpClient _httpClient = new HttpClient();
                    var jsonContent = new StringContent(JsonSerializer.Serialize(emailData), Encoding.UTF8, "application/json");

                    if (!_httpClient.DefaultRequestHeaders.Contains("Authorization"))
                    {
                        _httpClient.DefaultRequestHeaders.Add("Authorization", ConfigurationManager.AppSettings["Authkey"]);
                    }

                    var response = await _httpClient.PostAsync(ConfigurationManager.AppSettings["EmailAPIURL"], jsonContent);

                    var responseBody = await response.Content.ReadAsStringAsync();
                    Console.WriteLine($"Response Status: {response.StatusCode}");
                    Console.WriteLine($"Response Body: {responseBody}");

                    if (!response.IsSuccessStatusCode)
                    {
                        throw new Exception($"Email API error: {response.StatusCode} - {responseBody}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Email API Failed: {ex.Message}");
                }

            }
           
        }

       /* private void WriteToFile(string text)
        {
            var filepath = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory())).Root + $@"";
            string path = filepath + "log_" + DateTime.Now.ToString("dd-MMM-yyyy") + ".txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }
        */
        

    }

}
