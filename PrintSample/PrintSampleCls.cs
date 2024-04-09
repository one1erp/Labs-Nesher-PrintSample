using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Text;
using System.Windows.Forms;
using Common;
using DAL;
using LSEXT;
using LSSERVICEPROVIDERLib;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;
using XmlService;

namespace PrintSample
{

    [ComVisible(true)]
    [ProgId("PrintSample.PrintSampleCls")]
    public class PrintSampleCls : IWorkflowExtension
    {
        INautilusServiceProvider sp;
        private int _port = 9100;
        private DataLayer dal;

        private string stickType = "1";
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern SafeFileHandle CreateFile(string lpFileName, FileAccess dwDesiredAccess,
        uint dwShareMode, IntPtr lpSecurityAttributes, FileMode dwCreationDisposition,
        uint dwFlagsAndAttributes, IntPtr hTemplateFile);

        public void Execute(ref LSExtensionParameters Parameters)
        {
            try
            {
                #region params
                string tableName = Parameters["TABLE_NAME"];
                sp = Parameters["SERVICE_PROVIDER"];
                var rs = Parameters["RECORDS"];
                var sampleName = rs.Fields["NAME"].Value;
                var sampleDscription = rs.Fields["DESCRIPTION"].Value ?? "";
                var sampleID = rs.Fields["SAMPLE_ID"].Value;
                var workstationId = Parameters["WORKSTATION_ID"];
                #endregion
                ////////////יוצר קונקשן//////////////////////////
                var ntlCon = Utils.GetNtlsCon(sp);
                Utils.CreateConstring(ntlCon);
                /////////////////////////////           
                dal = new DataLayer();
                dal.Connect();
                var phrase = dal.GetPhraseByID(242);//Printout Type
                var pn = phrase.PhraseEntries.Where(x => x.PhraseDescription == "מדבקות דוגמה").FirstOrDefault();
                stickType = pn.PhraseName;
              
                Workstation ws = dal.getWorkStaitionById(workstationId);
                ReportStation reportStation = dal.getReportStationByWorksAndType(ws.NAME, stickType);
                string printerName = "";
                //string ip = GetIp(printerName);
                string goodIp = ""; //removeBadChar(ip);
                if (reportStation != null
                    )
                {
                    if (reportStation.Destination != null)
                    {
                        //                            printerName = reportStation.DESTINATION.INFO_TEXT1;
                        goodIp = reportStation.Destination.ManualIP;
                    }
                    if (reportStation.Destination != null && reportStation.Destination.RawTcpipPort != null)
                    {
                        _port = (int)reportStation.Destination.RawTcpipPort;
                    }

                    Print(sampleName.ToString(), sampleDscription.ToString(), sampleID.ToString(), goodIp);
                }

                else
                {
                    MessageBox.Show("לא הוגדרה תחנה");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("נכשלה הדפסת מדבקה");
                Logger.WriteLogFile(ex);
            }


        }


        private string removeBadChar(string ip)
        {
            string ret = "";
            foreach (var c in ip)
            {
                int ascii = (int)c;
                if ((ascii >= 48 && ascii <= 57) || ascii == 44 || ascii == 46)
                    ret += c;
            }
            return ret;
        }
        public string GetIp(string printerName)
        {
            string query = string.Format("SELECT * from Win32_Printer WHERE Name LIKE '%{0}'", printerName);
            string ret = "";
            var searcher = new ManagementObjectSearcher(query);
            var coll = searcher.Get();
            foreach (ManagementObject printer in coll)
            {
                foreach (PropertyData property in printer.Properties)
                {
                    if (property.Name == "PortName")
                    {
                        ret = property.Value.ToString();
                    }
                }
            }
            return ret;
        }
        private static string ReverseString(string s)
        {
            var str = s;
            string[] strsubs = s.Split(Convert.ToChar(" "));
            var newstr = "";
            string substr = "";
            int i;
            int c = strsubs.Count();
            for (i = 0; i < c; ++i)
            {
                substr = strsubs[i];
                if (HasHebrewChar(strsubs[i]))
                {
                    substr = Reverse(substr);
                }

                newstr += substr + " ";
            }
            return newstr;
        }

        private static string Reverse(string s)
        {
            char[] arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        public static bool HasHebrewChar(string value)
        {
            return value.ToCharArray().Any(x => (x <= 'ת' && x >= 'א'));
        }

        private void Print(string name, string description, string ID, string ip)
        {
            string ipAddress = ip;
            // ZPL Command(s)
            string ntxt = name;
            string dtxt = "";
            if (HasHebrewChar(description))
            {
                var split = ReverseString(description).Split(' ');
                split.Reverse();
                foreach (string s in split)
                {
                    dtxt = s + " " + dtxt;
                }
            }
            else
            {
                dtxt = description;
            }
            //            dtxt = split.ToString();
            string itxt = ID;
            string ZPLString =
            "^XA" +
            "^CI28" +
            "^LH13,10" +
            "^FO13,15" +
            "^A@N30,30" +
           string.Format("^FD{0}^FS", ntxt) +
            "^FO10,180" +
            "^A@N30,30" +
            string.Format("^FD{0}^FS", "dtxt") +//
            
            "^FO260,50" + "^BQN,4,5" +
                //string.Format("^FD   {0}^FS", itxt) +
             string.Format("^FDLA,{0}^FS", itxt) + //ברקוד
           "^XZ";
            try
            {
                MessageBox.Show(ntxt + " name1");
                MessageBox.Show(dtxt + " description");
                MessageBox.Show(itxt + " code");
                // Open connection
                System.Net.Sockets.TcpClient client = new System.Net.Sockets.TcpClient();
                client.Connect(ipAddress, _port);
                // Write ZPL String to connection
                StreamWriter writer = new StreamWriter(client.GetStream());
                writer.Write(ZPLString);
                writer.Flush();
                // Close Connection
                writer.Close();
                client.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.InnerException.Message);
            }
        }



    }
}
