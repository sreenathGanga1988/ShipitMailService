using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net.Mail;
using System.Threading;
using System.Configuration;
using System.Data.SqlClient;
namespace ShipitMailService
{
    public partial class Service1 : ServiceBase
    {
        static String Consolidatedmssg = "";
        static String ConnStr = "Data Source=192.168.1.4;Initial Catalog=CourierDetails;Persist Security Info=True;User ID=sa;Password=Sreenath@123";
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            this.WriteToFile("Shipit Service started {0}");
            this.ScheduleService();
        }

        protected override void OnStop()
        {
            this.WriteToFile("Shipit Service stopped {0}");
            this.Schedular.Dispose();
        }

        private Timer Schedular;

        public void ScheduleService()
        {
            try
            {
                Schedular = new Timer(new TimerCallback(SchedularCallback));
                string mode = ConfigurationManager.AppSettings["Mode"].ToUpper();
                this.WriteToFile("ShipiT Service Mode: " + mode + " {0}");

                //Set the Default Time.
                DateTime scheduledTime = DateTime.MinValue;

                if (mode == "DAILY")
                {
                    //Get the Scheduled Time from AppSettings.
                    scheduledTime = DateTime.Parse(System.Configuration.ConfigurationManager.AppSettings["ScheduledTime"]);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next day.
                        scheduledTime = scheduledTime.AddDays(1);
                    }
                }

                if (mode.ToUpper() == "INTERVAL")
                {
                    //Get the Interval in Minutes from AppSettings.
                    int intervalMinutes = Convert.ToInt32(ConfigurationManager.AppSettings["IntervalMinutes"]);

                    //Set the Scheduled Time by adding the Interval to Current Time.
                    scheduledTime = DateTime.Now.AddMinutes(intervalMinutes);
                    if (DateTime.Now > scheduledTime)
                    {
                        //If Scheduled Time is passed set Schedule for the next Interval.
                        scheduledTime = scheduledTime.AddMinutes(intervalMinutes);
                    }
                }

                TimeSpan timeSpan = scheduledTime.Subtract(DateTime.Now);
                string schedule = string.Format("{0} day(s) {1} hour(s) {2} minute(s) {3} seconds(s)", timeSpan.Days, timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

                this.WriteToFile("ShipiT Service scheduled to run after: " + schedule + " {0}");

                //Get the difference in Minutes between the Scheduled and Current Time.
                int dueTime = Convert.ToInt32(timeSpan.TotalMilliseconds);

                //Change the Timer's Due Time.
                Schedular.Change(dueTime, Timeout.Infinite);
            }
            catch (Exception ex)
            {
                WriteToFile("ShipiT Service Error on: {0} " + ex.Message + ex.StackTrace);

                //Stop the Windows Service.
                using (System.ServiceProcess.ServiceController serviceController = new System.ServiceProcess.ServiceController("SimpleService"))
                {
                    serviceController.Stop();
                }
            }
        }

        private void SchedularCallback(object e)
        {
            Consolidatedmssg = "";
            LastCloseDateofFactory();
            this.ScheduleService();
        }

        private void WriteToFile(string text)
        {
            string path = "D:\\ServiceLog.txt";
            using (StreamWriter writer = new StreamWriter(path, true))
            {
                writer.WriteLine(string.Format(text, DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt")));
                writer.Close();
            }
        }



        public static void LastCloseDateofFactory()
        {
            DateTime today = DateTime.Now.Date;
            DataTable dt = new DataTable();

            String Datepend = "";
            using (SqlConnection con = new SqlConnection(ConnStr))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;


                String q3 = @"SELECT        MAX(ProductionDayClose_tbl.DateofProd) AS Closedate, Factory_Master.Factory_name
FROM            ProductionDayClose_tbl INNER JOIN
                         Factory_Master ON ProductionDayClose_tbl.factid = Factory_Master.Factory_ID
GROUP BY ProductionDayClose_tbl.factid, Factory_Master.Factory_name";



                //  cmd.CommandText = Query1;
                cmd.CommandText = q3;

                SqlDataReader rdr = cmd.ExecuteReader();

                dt.Load(rdr);

                if (dt != null)
                {
                    if (dt.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            Datepend = "";
                            DateTime closeddate = DateTime.Parse(dt.Rows[i]["Closedate"].ToString()).Date;

                            double daysUntilChristmas = today.Subtract(closeddate).TotalDays;

                            if (daysUntilChristmas > 1)
                            {
                                String factoryname = dt.Rows[i]["Factory_name"].ToString();
                                for (DateTime j = closeddate.AddDays(1); j < today; j = j.AddDays(1))
                                {
                                    if (j.DayOfWeek != DayOfWeek.Sunday)
                                    {

                                        if (Datepend == "")
                                        {
                                            Datepend = Datepend + j.ToString("dd/MM/yyyy");
                                        }
                                        else
                                        {
                                            Datepend = Datepend + " ****** " + j.ToString("dd/MM/yyyy");
                                        }
                                    }
                                }

                                Dayspendingtoclose(factoryname, Datepend, closeddate.ToString("dd/MM/yyyy"));
                            }
                        }


                        DataTable dt1 = GetEmailAdress("ATR", "A", "Pending Report");
                        email23("Production Day Close Status", Consolidatedmssg, "ram_santhosh@atraco.ae", "User", "Production Notification", dt1);

                    }
                }

            }

        }


        public static DataTable GetEmailAdress(String Factory, String Userlevel, String Reporttype)
        {

            DateTime today = DateTime.Now.Date;
            DataTable dt = new DataTable();


            using (SqlConnection con = new SqlConnection(ConnStr))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = con;


                String q3 = @"SELECT        Factory_Name, Userlevel, ReportType, EmailGroup
FROM            EmailGroup_tbl
WHERE        (Factory_Name = @Param1) AND (Userlevel = @Param2) AND (ReportType = @Param3)";




                cmd.CommandText = q3;
                cmd.Parameters.AddWithValue("@Param1", Factory);
                cmd.Parameters.AddWithValue("@Param2", Userlevel);
                cmd.Parameters.AddWithValue("@Param3", Reporttype);
                SqlDataReader rdr = cmd.ExecuteReader();

                dt.Load(rdr);
            }
            return dt;
        }
        public static void Dayspendingtoclose(String factoryname, String DateforClosing, String DayClosed)
        {
            String subject = "Production Entry Day Close Pending For " + factoryname + "  for Date " + DateforClosing;

            String body = "Production Entry Day Close Pending For " + factoryname + "  for Date " + DateforClosing + "  Last day Closed was " + DayClosed;


            if (factoryname == "AA2")
            {

            }
            else if (factoryname == "MA2")
            {

            }
            else if (factoryname == "MA3")
            {

            }
            else if (factoryname == "MA1")
            {

            }
            DataTable dt = GetEmailAdress(factoryname, "U", "Pending Report");

            email23(subject, body, "ram_santhosh@atraco.ae", "User", "Production Notification",dt);


            Consolidatedmssg = Consolidatedmssg + Environment.NewLine + body;

        }
        public static void email23(String subject, String body, String toadress, String toname, String Displayname, DataTable dt)
        {
            var fromAddress = new MailAddress("atracogen@gmail.com", Displayname);
            var toAddress = new MailAddress(toadress, toname);

            const string fromPassword = "8812686ba";


            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new System.Net.NetworkCredential(fromAddress.Address, fromPassword)
            };

            using (var message = new MailMessage(fromAddress, toAddress) { Subject = subject, Body = body }




                )
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    message.CC.Add(dt.Rows[i]["EmailGroup"].ToString());
                }






                smtp.Send(message);

            }
        }
    }
}
