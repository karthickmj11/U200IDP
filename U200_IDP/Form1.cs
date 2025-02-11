using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using CheckBox = System.Windows.Forms.CheckBox;
using DocumentFormat.OpenXml.Spreadsheet;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Control = System.Windows.Forms.Control;
using Point = System.Drawing.Point;
using System.Runtime.InteropServices.ComTypes;
using static System.Net.WebRequestMethods;
using File = System.IO.File;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using System.Net.Sockets;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Threading;

namespace ExceltoXML
{
    public partial class Form1 : Form
    {
        private string selectedPath;
        private List<string> Out_Filename = new List<string>();
        private List<string> files = new List<string> { };

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Visible = false;
            button1.Enabled = false;
            button3.Enabled = false;
            button4.Visible = false;
            button5.Visible = false;
            ToolTip toolTip2 = new ToolTip();
            toolTip2.SetToolTip(button2, "Click this button to choose input file");
            string subdir = @"C:\Users\491497\Documents\projects\Outputs";
            //string subdir = @"D:\U200_IDP";
            if (!Directory.Exists(subdir))
            {
                Directory.CreateDirectory(subdir);
            }
           
        }

    private void button1_Click(object sender, EventArgs e)
        {
            button2.Enabled = false;

            Thread Run_main = new Thread(new ThreadStart(Run));
            Run_main.Start();
        }




    private void Run()
        {
            //MessageBox.Show("Processing...", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            List<string> Out_Filename = new List<string>();
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
            foreach (Worksheet sheet in xlWorkbook.Sheets)
            {
                Out_Filename.Add(sheet.Name);
                files.Add(@selectedPath + @"\" + Out_Filename + ".xml");
            }
            string Out = @selectedPath + @"\All ASCVs.xml";
            FileInfo f = new FileInfo(Out);
            if (f.Exists)
            {
                f.Delete();
            }

            if (TrackCircuit.Checked == true)
            {
                
                
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                
                
                string fileName = @selectedPath + @"\"+ Out_Filename[0] +".xml";
                Console.WriteLine(fileName);
                //string fileName = @"D:\U200_IDP\SOA1_TC.xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }

                string Check(string a)
                {
                    if(a.Substring(0,2) == "TC")
                    {
                        return "T"+a.Substring(2);
                    }
                    return a;
                }

                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            


                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<SECONDARIEs>");
                                sw.WriteLine("<Secondary ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Secondary>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {

                                sw.WriteLine("<Secondary ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Secondary>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</SECONDARIEs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (Subroute.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[2];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                string fileName = @selectedPath + @"\" + Out_Filename[1] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Subroute.xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {



                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<SUBROUTEs>");
                                sw.WriteLine("<Subroute ID="+ "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Subroute>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {

                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Subroute ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Subroute>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</SUBROUTEs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }


            if (TrafficDirection.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[3];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                string fileName = @selectedPath + @"\" + Out_Filename[2] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Trafficdirection.xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                string Check(string a)
                {
                    if (a[2]!='_')
                    {
                        a.Insert(2, "_");
                        return a;
                    }
                    return a;
                }

                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {



                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<TRAFFICDIRECTIONs>");
                                sw.WriteLine("<TrafficDirection ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</TrafficDirection>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {

                                sw.WriteLine("<TrafficDirection ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</TrafficDirection>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</TRAFFICDIRECTIONs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }


            if (ESP.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[4];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                string fileName = @selectedPath + @"\" + Out_Filename[3] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_ESP.xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                string Check(string a)
                {
                    if (a[3]!='_')
                    {
                        a.Insert(3, "_");
                    }
                    return a;
                }

                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {



                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<ESPs>");
                                sw.WriteLine("<ESP ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<ESPStatus RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</ESP>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {

                                sw.WriteLine("<ESP ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + Check(xlRange.Cells[i, j + 1].Value2.ToString()) + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<ESPStatus RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</ESP>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</ESPs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }


            if (Point.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[5];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                int constant = 1;
                int DN, DR;
                int flag = 0;
                string fileName = @selectedPath + @"\" + Out_Filename[4] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }

                for (int i = 2; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, j].value2 !=null && xlRange.Cells[i, j].value2.ToString() == "SPL")
                    {
                        flag = 1;
                        i++;
                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2 && flag == 0)
                    {

                        //

                        try
                        {


                            DN = constant + Convert.ToInt32((xlRange.Cells[i, j + 3].Value2.ToString()));
                            DR = constant + Convert.ToInt32(xlRange.Cells[i, j + 5].Value2.ToString());
                            //DPNL = constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString());
                            //DPNR = constant + Convert.ToInt32(xlRange.Cells[i, j + 9].Value2.ToString());

                            using (StreamWriter sw = fi.CreateText())
                            {
                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<POINTs>");
                                sw.WriteLine("<Point ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<DetectedNormal RM_Position =" + "\"" + DN.ToString() + "\"" + "/>");
                                sw.WriteLine("<DetectedReverse RM_Position =" + "\"" + DR.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 7].Value2 != null)
                                {
                                    sw.WriteLine("<DetectedLockNormal RM_Position =" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString())).ToString() + "\"" + "/>");
                                }
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<DetectedLockReverse RM_Position =" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 9].Value2.ToString())).ToString() + "\"" + "/>");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Point>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2 && flag == 0)
                    {
                        DN = constant + Convert.ToInt32((xlRange.Cells[i, j + 3].Value2.ToString()));
                        DR = constant + Convert.ToInt32(xlRange.Cells[i, j + 5].Value2.ToString());
                        //DPNL = constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString());
                        //DPNR = constant + Convert.ToInt32(xlRange.Cells[i, j + 9].Value2.ToString());


                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Point ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<DetectedNormal RM_Position =" + "\"" + DN.ToString() + "\"" + "/>");
                                sw.WriteLine("<DetectedReverse RM_Position =" + "\"" + DR.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 7].Value2 != null)
                                {
                                    Console.WriteLine(i);
                                    sw.WriteLine("<DetectedLockNormal RM_Position =" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString())).ToString() + "\"" + "/>");
                                }
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<DetectedLockReverse RM_Position =" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 9].Value2.ToString())).ToString() + "\"" + "/>");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Point>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }
                    }
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2 && flag == 1)
                    {
                       /* int P_ISN = constant + Convert.ToInt32((xlRange.Cells[i, j + 9].Value2.ToString()));
                        int C_ISN = constant + Convert.ToInt32(xlRange.Cells[i, j + 11].Value2.ToString());
                        int P_ASN = constant + Convert.ToInt32(xlRange.Cells[i, j + 5].Value2.ToString());
                        int C_ASN = constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString());
                        int ISN   = constant + Convert.ToInt32(xlRange.Cells[i, j + 3].Value2.ToString());*/
                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Point ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SelfNormalisationInhibitionPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 9].Value2.ToString())).ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 11].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SelfNormalisationInhibitionConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 11].Value2.ToString())).ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 5].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SelfNormalisationActivationPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 5].Value2.ToString())).ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 7].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SelfNormalisationActivationConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 7].Value2.ToString())).ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                if (xlRange.Cells[i, j + 3].Value2 != null)
                                {
                                    sw.WriteLine("<DetectedLockReverse RM_Position =" + "\"" + (constant + Convert.ToInt32(xlRange.Cells[i, j + 3].Value2.ToString())).ToString() + "\"" + "/>");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Point>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                    }


                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</POINTs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

               //xlWorkbook.Close();
               //Marshal.ReleaseComObject(xlWorkbook);
               //
               //xlApp.Quit();
               //Marshal.ReleaseComObject(xlApp);
            }


            if (Signal.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[6];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                string fileName = @selectedPath + @"\" + Out_Filename[5] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Signal.xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }

                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {



                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<SIGNALs>");
                                sw.WriteLine("<HILC_TIME>\r\n    <REQUEST_TO_SR_HOUR_TIMER>5000</REQUEST_TO_SR_HOUR_TIMER>\r\n    <REQUEST_TO_CONFIRMATION_TIMER>26500</REQUEST_TO_CONFIRMATION_TIMER>\r\n    <REQUEST_TO_RETURN_CODE_TIMER>27000</REQUEST_TO_RETURN_CODE_TIMER>\r\n    <SESSION_TO_REQUEST_TIMER>1500</SESSION_TO_REQUEST_TIMER>\r\n    <REPLY_TIME_OUT>5000</REPLY_TIME_OUT>\r\n</HILC_TIME>");
                                sw.WriteLine("<Signal ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                if (xlRange.Cells[i, j + 36].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalBlockingPreparation\">");
                                    sw.WriteLine("  <Set RC_Position=\""+ xlRange.Cells[i, j + 36].Value2.ToString()+ "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 35].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 39].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalBlockingConfirmation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 39].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 38].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 42].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalUnblockingPreparation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 42].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 41].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 45].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalUnblockingConfirmation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 45].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 44].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                if (xlRange.Cells[i, j + 33].Value2 != null)
                                {
                                    sw.WriteLine("<ToONStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 32].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 3].Value2 != null)
                                {
                                    sw.WriteLine("<FreeApproachLocking RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 2].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 6].Value2 != null)
                                {
                                    sw.WriteLine("<VioletRedLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 5].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<GreenLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 8].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 12].Value2 != null)
                                {
                                    sw.WriteLine("<SignalON RM_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 11].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 15].Value2 != null)
                                {
                                    sw.WriteLine("<SignalOFFViolet RM_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 14].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 21].Value2 != null)
                                {
                                    sw.WriteLine("<ShuntOFF RM_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 20].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 18].Value2 != null)
                                {
                                    sw.WriteLine("<RouteIndicatorLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 17].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 24].Value2 != null)
                                {
                                    sw.WriteLine("<BufferStopLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 23].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 27].Value2 != null)
                                {
                                    sw.WriteLine("<OneAspects RM_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 26].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 30].Value2 != null)
                                {
                                    sw.WriteLine("<TwoAspects RM_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 29].Value2.ToString() + "-->");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Signal>");
                                

                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {

                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Signal ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                if (xlRange.Cells[i, j + 36].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalBlockingPreparation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 36].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 35].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 39].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalBlockingConfirmation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 39].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 38].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 42].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalUnblockingPreparation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 42].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 41].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 45].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SignalUnblockingConfirmation\">");
                                    sw.WriteLine("  <Set RC_Position=\"" + xlRange.Cells[i, j + 45].Value2.ToString() + "\"/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 44].Value2.ToString() + "-->");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                if (xlRange.Cells[i, j + 33].Value2 != null)
                                {
                                    sw.WriteLine("<ToONStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 32].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 3].Value2 != null)
                                {
                                    sw.WriteLine("<FreeApproachLocking RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 2].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 6].Value2 != null)
                                {
                                    sw.WriteLine("<VioletRedLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 5].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<GreenLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 8].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 12].Value2 != null)
                                {
                                    sw.WriteLine("<SignalON RM_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 11].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 15].Value2 != null)
                                {
                                    sw.WriteLine("<SignalOFFViolet RM_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 14].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 21].Value2 != null)
                                {
                                    sw.WriteLine("<ShuntOFF RM_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 20].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 18].Value2 != null)
                                {
                                    sw.WriteLine("<RouteIndicatorLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 17].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 24].Value2 != null)
                                {
                                    sw.WriteLine("<BufferStopLampProved RM_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 23].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 27].Value2 != null)
                                {
                                    sw.WriteLine("<OneAspects RM_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 26].Value2.ToString() + "-->");
                                }
                                if (xlRange.Cells[i, j + 30].Value2 != null)
                                {
                                    sw.WriteLine("<TwoAspects RM_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("<!-- Name =" + xlRange.Cells[i, j + 29].Value2.ToString() + "-->");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Signal>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</SIGNALs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }


            if (Overlap.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[7];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                
                string fileName = @selectedPath + @"\" + Out_Filename[6] + ".xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            


                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<OVERLAPs>");
                                sw.WriteLine("<Overlap ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Overlap>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {

                                sw.WriteLine("<Overlap ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Overlap>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</OVERLAPs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }


            if (PSAPR.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[8];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;

                string fileName = @selectedPath + @"\" + Out_Filename[7] + ".xml";

                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            


                            using (StreamWriter sw = fi.CreateText())
                            {

                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<PSAPRs>");
                                sw.WriteLine("<PSAPR ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<SecondaryStationPowerSupplyState RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</PSAPR>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {

                                sw.WriteLine("<PSAPR ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<SecondaryStationPowerSupplyState RM_Position =" + "\"" + xlRange.Cells[i, j + 2].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</PSAPR>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</PSAPRs>");

                    }
                }


                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (Cycle.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[9];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                //int constant = 1;
                //string CYC, CYD;

                string fileName = @selectedPath + @"\" + Out_Filename[8] + ".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            

                            

                            using (StreamWriter sw = fi.CreateText())
                            {
                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<CYCLEs>");
                                sw.WriteLine("<Cycle ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC REPLY_TIME_OUT=\"0\" Name=\"Setting\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC REPLY_TIME_OUT=\"0\" Name=\"Cancellation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<State RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Cycle>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Cycle ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC REPLY_TIME_OUT=\"0\" Name=\"Setting\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC REPLY_TIME_OUT=\"0\" Name=\"Cancellation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<State RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Cycle>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</CYCLEs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (MBL.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[10];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                //int constant = 1;
                //string CYC, CYD;

                string fileName = @selectedPath + @"\"+ Out_Filename[9] +".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            



                            using (StreamWriter sw = fi.CreateText())
                            {
                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<MBLs>");
                                sw.WriteLine("<MaintenanceBlock ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockControlPreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockControlConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockReleasePreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockReleaseConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</MaintenanceBlock>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<MaintenanceBlock ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockControlPreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockControlConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockReleasePreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"MaintenanceBlockReleaseConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<Status RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</MaintenanceBlock>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</MBLs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (Route.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[11];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                //int constant = 1;
                //string CYC, CYD;

                string fileName = @selectedPath + @"\"+ Out_Filename[10] +".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            



                            using (StreamWriter sw = fi.CreateText())
                            {
                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<ROUTEs>");
                                sw.WriteLine("<Route ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC Name=\"RouteControlConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteReleaseConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteControlPreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteReleasePreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteControl\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteCancellation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                if (xlRange.Cells[i, j + 29].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"PermanentRouteControl\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                    sw.WriteLine("<NoHILC Name=\"PermanentRouteCancellation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<RouteBlocking RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("<RouteSetLocked RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 9].Value2!= null)
                                {
                                    sw.WriteLine("<RoutePermanent RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Route>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Route ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                sw.WriteLine("<NoHILC Name=\"RouteControlConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteReleaseConfirmation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteControlPreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteReleasePreparation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteControl\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                sw.WriteLine("<NoHILC Name=\"RouteCancellation\">");
                                sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</NoHILC>");
                                if (xlRange.Cells[i, j + 29].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"PermanentRouteControl\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                    sw.WriteLine("<NoHILC Name=\"PermanentRouteCancellation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                sw.WriteLine("<RouteBlocking RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("<RouteSetLocked RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                {
                                    sw.WriteLine("<RoutePermanent RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                }
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Route>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }
                    

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</ROUTEs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (Switch.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[12];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int j = 1;
                //int constant = 1;
                //string CYC, CYD;

                string fileName = @selectedPath + @"\"+ Out_Filename[11] +".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                for (int i = 2; i <= rowCount; i++)
                {

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i == 2)
                    {

                        //

                        try
                        {

                            



                            using (StreamWriter sw = fi.CreateText())
                            {
                                sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                                sw.WriteLine("Author: ATS Integration Team");
                                sw.WriteLine("");
                                sw.WriteLine("<SWITCHs>");
                                sw.WriteLine("<HILC_TIME>\r\n    <REQUEST_TO_SR_HOUR_TIMER>5000</REQUEST_TO_SR_HOUR_TIMER>\r\n    <REQUEST_TO_CONFIRMATION_TIMER>26500</REQUEST_TO_CONFIRMATION_TIMER>\r\n    <REQUEST_TO_RETURN_CODE_TIMER>27000</REQUEST_TO_RETURN_CODE_TIMER>\r\n    <SESSION_TO_REQUEST_TIMER>1500</SESSION_TO_REQUEST_TIMER>\r\n    <REPLY_TIME_OUT>5000</REPLY_TIME_OUT>\r\n</HILC_TIME>");
                                sw.WriteLine("<Switch ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                if (xlRange.Cells[i, j + 21].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchBlockingConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 27].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchUnblockingConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 18].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchBlockingPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 24].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchUnblockingPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }if (xlRange.Cells[i, j + 30].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"CallingNormal\" REPLY_TIME_OUT=\"5000\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 33].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"CallingReverse\" REPLY_TIME_OUT=\"5000\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 36].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"ManualAuthorizationControl\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 36].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 39].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"ManualAuthorizationRelease\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 39].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                if(xlRange.Cells[i, j + 9].Value2 !=null)
                                sw.WriteLine("<BlockingStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 3].Value2 != null)
                                sw.WriteLine("<CalledNormal RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 6].Value2 != null)
                                sw.WriteLine("<CalledReverse RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 12].Value2 != null)
                                sw.WriteLine("<AuthorisationStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 15].Value2 != null)
                                sw.WriteLine("<KeyStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Switch>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }


                        //

                    }

                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null && i > 2)
                    {



                        try
                        {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                sw.WriteLine("<Switch ID=" + "\"" + xlRange.Cells[i, j].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, j + 1].Value2.ToString() + "\"" + ">");
                                if (xlRange.Cells[i, j + 21].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchBlockingConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 21].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 27].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchUnblockingConfirmation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 27].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 18].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchBlockingPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 18].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 24].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"SwitchUnblockingPreparation\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 24].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 30].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"CallingNormal\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 30].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 33].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"CallingReverse\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 33].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 36].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"ManualAuthorizationControl\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 36].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                if (xlRange.Cells[i, j + 39].Value2 != null)
                                {
                                    sw.WriteLine("<NoHILC Name=\"ManualAuthorizationRelease\">");
                                    sw.WriteLine("<Set RC_Position=" + "\"" + xlRange.Cells[i, j + 39].Value2.ToString() + "\"" + "/>");
                                    sw.WriteLine("</NoHILC>");
                                }
                                sw.WriteLine("<RM>");
                                if (xlRange.Cells[i, j + 9].Value2 != null)
                                    sw.WriteLine("<BlockingStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 9].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 3].Value2 != null)
                                    sw.WriteLine("<CalledNormal RM_Position=" + "\"" + xlRange.Cells[i, j + 3].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 6].Value2 != null)
                                    sw.WriteLine("<CalledReverse RM_Position=" + "\"" + xlRange.Cells[i, j + 6].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 12].Value2 != null)
                                    sw.WriteLine("<AuthorisationStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 12].Value2.ToString() + "\"" + "/>");
                                if (xlRange.Cells[i, j + 15].Value2 != null)
                                    sw.WriteLine("<KeyStatus RM_Position=" + "\"" + xlRange.Cells[i, j + 15].Value2.ToString() + "\"" + "/>");
                                sw.WriteLine("</RM>");
                                sw.WriteLine("</Switch>");
                            }


                        }
                        catch (Exception Ex)
                        {
                            Console.WriteLine(Ex.ToString());
                        }

                    }

                }
                if (xlRange.Cells[2, 1].Value2 != null)
                {
                    using (StreamWriter sw = fi.AppendText())
                    {

                        sw.WriteLine("</SWITCHs>");

                    }
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (ASCV.Checked == true)
            {
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[13];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                string fileName = @selectedPath + @"\"+ Out_Filename[12] +".xml";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                if (fi.Exists)
                {
                    fi.Delete();
                }
                int j=1;
                string s = "";
                string t = "";
                string l = "";
                int id = 1;
                //
                for (int i = 2; i <= rowCount; i++)
                {
                    if (xlRange.Cells[i, j + 1] != null && xlRange.Cells[i, j + 1].Value2 != null)
                    {
                        if (xlRange.Cells[i, 4].Value2 != null && i == 6)
                            s += "<SettingAck RM_Position=" + "\"" + xlRange.Cells[6, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 7)
                            s += "<ReleaseAck RM_Position=" + "\"" + xlRange.Cells[7, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 4)
                            s += "<RouteBlockingAck RM_Position=" + "\"" + xlRange.Cells[4, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 5)
                            s += "<RouteUnblockingAck RM_Position=" + "\"" + xlRange.Cells[5, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 2)
                            s += "<PointBlockingAck RM_Position=" + "\"" + xlRange.Cells[2, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 3)
                            s += "<PointUnblockingAck RM_Position=" + "\"" + xlRange.Cells[3, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 14)
                            s += "<UninterruptedPowerSupply1State RM_Position=" + "\"" + xlRange.Cells[14, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 15)
                            s += "<UninterruptedPowerSupply2State RM_Position=" + "\"" + xlRange.Cells[15, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 16)
                            s += "<PowerSupply1State RM_Position=" + "\"" + xlRange.Cells[16, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 17)
                            s += "<PowerSupply2State RM_Position=" + "\"" + xlRange.Cells[17, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 13)
                            s += "<PowerSupplyAlarmState RM_Position=" + "\"" + xlRange.Cells[13, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 18)
                            s += "<VDUModeControlState RM_Position=" + "\"" + xlRange.Cells[18, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 19)
                            s += "<ATSModeControlState RM_Position=" + "\"" + xlRange.Cells[19, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 8)
                            s += "<GlobalSignalSettingAck RM_Position=" + "\"" + xlRange.Cells[8, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 9)
                            s += "<GlobalSignalReleaseAck RM_Position=" + "\"" + xlRange.Cells[9, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 10)
                            s += "<GlobalSignalBlock RM_Position=" + "\"" + xlRange.Cells[10, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 11)
                            s += "<ActivationAck RM_Position=" + "\"" + xlRange.Cells[11, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 12)
                            s += "<InhibitionAck RM_Position=" + "\"" + xlRange.Cells[12, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 20)
                            s += "<GlobalPointUnblockingAck RM_Position=" + "\"" + xlRange.Cells[20, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 21)
                            s += "<GlobalRouteUnblockingAck RM_Position=" + "\"" + xlRange.Cells[21, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 22)
                            s += "<SignalBlockingAck RM_Position=" + "\"" + xlRange.Cells[22, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 23)
                            s += "<SignalReleaseAck RM_Position=" + "\"" + xlRange.Cells[23, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i == 24)
                            s += "<TDR RM_Position=" + "\"" + xlRange.Cells[24, 4].Value2.ToString() + "\"" + "/>\n";
                        if (xlRange.Cells[i, 4].Value2 != null && i >= 25)
                        {
                            l +="<ASCV ID=" + "\"" + (++id).ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[i, 3].Value2.ToString().Substring(3) + "\"" + ">\n";
                            l += "<RM>\n";
                            l += "<AccessToNonInterlocking RM_Position=" + "\"" + xlRange.Cells[i, 4].Value2.ToString() + "\"" + "/>\n";
                            l += "</RM>\n";
                            l += "</ASCV>\n";
                        }

                        if (xlRange.Cells[i, 7].Value2 != null && i == 2)
                        {
                            t += "<NoHILC Name=\"GlobalSignalBlockingPreparation\">\n";
                            t += "<Set RC_Position=" + "\"" + xlRange.Cells[2, 7].Value2.ToString() + "\"" + "/>\n";
                            t += "</NoHILC>\n";
                        }
                        if (xlRange.Cells[i, 7].Value2 != null && i == 3)
                        {
                            t += "<NoHILC Name=\"GlobalSignalBlockingConfirmation\">\n";
                            t += "<Set RC_Position=" + "\"" + xlRange.Cells[3, 7].Value2.ToString() + "\"" + "/>\n";
                            t += "</NoHILC>\n";
                        }
                        if (xlRange.Cells[i, 7].Value2 != null && i == 4)
                        {
                            t += "<NoHILC Name=\"GlobalSignalUnblockingPreparation\">\n";
                            t += "<Set RC_Position=" + "\"" + xlRange.Cells[4, 7].Value2.ToString() + "\"" + "/>\n";
                            t += "</NoHILC>\n";
                        }
                        if (xlRange.Cells[i, 7].Value2 != null && i == 5)
                        {
                            t += "<NoHILC Name=\"GlobalSignalUnblockingConfirmation\">\n";
                            t += "<Set RC_Position=" + "\"" + xlRange.Cells[5, 7].Value2.ToString() + "\"" + "/>\n";
                            t += "</NoHILC>\n";
                        }
                    }
                        
                    
                    }
                try
                {

                    



                    using (StreamWriter sw = fi.CreateText())
                    {

                        sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                        sw.WriteLine("Author: ATS Integration Team");
                        sw.WriteLine("");
                        sw.WriteLine("<ASCVs>");
                        sw.WriteLine("<ASCV ID=" + "\"" + xlRange.Cells[2, 1].Value2.ToString() + "\"" + " " + "Name=" + "\"" + xlRange.Cells[2, 2].Value2.ToString() + "\"" + ">");
                        sw.WriteLine("<RM>");
                        sw.WriteLine(s);
                        sw.WriteLine("</RM>");
                        sw.WriteLine(t);
                        sw.WriteLine("</ASCV>");
                        sw.WriteLine(l);
                        sw.WriteLine("</ASCVs>");


                    }
                }





                catch (Exception Ex)
                {
                    Console.WriteLine(Ex.ToString());

                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                /*using (FileStream outputStream = new FileStream(Out, FileMode.Create))
                {
                    foreach (string file in files)
                    {
                        using (FileStream inputStream = new FileStream(files, FileMode.Open))
                        {
                            inputStream.CopyTo(outputStream);
                        }
                    }
                }*/

                //xlWorkbook.Close();
                //Marshal.ReleaseComObject(xlWorkbook);

                //xlApp.Quit();
                //Marshal.ReleaseComObject(xlApp);
            }

            if (Interlocking.Checked == true)
            {
                string fileName = @selectedPath + @"\"+ Out_Filename[13] +".XML";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[14];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                int Address;
                if (fi.Exists)
                {
                    fi.Delete();
                }
                if(xlRange.Cells[1, 1].Value2==null)
                {
                    MessageBox.Show("Cell[1,1] is empty in Sheet Interlocking. Please provide Sector name", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Environment.Exit(0);
                }
                String Sector = xlRange.Cells[1, 1].Value2.ToString();
                if (xlRange.Cells[1, 2].Value2 == null)
                {
                    MessageBox.Show("Cell[1,2] is empty in Sheet Interlocking. Please provide Sector code ", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Environment.Exit(0);
                }
                String Sector_Code = xlRange.Cells[1, 2].Value2.ToString();
                Dictionary<string, List<string>> equip = new Dictionary<string, List<string>>();
                equip.Add("ALS", new List<string> { "SIGNALs", "FreeApproachLocking","S","RM" });
                equip.Add("AMC", new List<string> { "ASCVs", "ATSModeControlState","","RM" });
                equip.Add("AU_ACC", new List<string> { "ASCVs", "AccessToNonInterlocking", "ACC","RM" });
                equip.Add("CONF_ASN", new List<string> { "POINTs", "SelfNormalisationActivationConfirmation", "P", "RC" });
                equip.Add("CONF_ISN", new List<string> { "POINTs", "SelfNormalisationInhibitionConfirmation", "P", "RC" });
                equip.Add("CONF_MBL", new List<string> { "MBLs", "MaintenanceBlockControlConfirmation", "MBL", "RC" });
                equip.Add("CONF_MBUL", new List<string> { "MBLs", "MaintenanceBlockReleaseConfirmation", "MBL", "RC" });
                equip.Add("CONF_MPL", new List<string> { "SWITCHs", "SwitchBlockingConfirmation","SW","RC" });
                equip.Add("CONF_MPLD", new List<string> { "SWITCHs", "SwitchUnblockingConfirmation","SW","RC" });
                equip.Add("CONF_RBL", new List<string> { "ROUTEs", "RouteControlConfirmation", "R", "RC" });
                equip.Add("CONF_RUBL", new List<string> { "ROUTEs", "RouteReleaseConfirmation", "R", "RC" });
                equip.Add("CONF_SBL", new List<string> { "SIGNALs", "SignalBlockingConfirmation", "S", "RC"});
                equip.Add("CONF_SUBL", new List<string> { "SIGNALs", "SignalUnblockingConfirmation", "S", "RC" });
                equip.Add("CY", new List<string> { "CYCLEs", "State","C","RM" });
                equip.Add("CYC", new List<string> { "CYCLEs", "Setting", "C", "RC" });
                equip.Add("CYD", new List<string> { "CYCLEs", "Cancellation", "C", "RC" });
                equip.Add("DPN", new List<string> { "POINTs", "DetectedNormal","P","RM" });
                equip.Add("DPNL", new List<string> { "POINTs", "DetectedLockNormal","P","RM" });
                equip.Add("DPR", new List<string> { "POINTs", "DetectedReverse","P","RM"});
                equip.Add("DPRL", new List<string> { "POINTs", "DetectedLockReverse","P","RM" });
                equip.Add("ESP", new List<string> { "ESPs", "ESPStatus","ESP_","RM" });
                equip.Add("FR", new List<string> { "ROUTEs", "RoutePermanent","R","RM" });
                equip.Add("FRC", new List<string> { "ROUTEs", "PermanentRouteControl", "R", "RC" });
                equip.Add("FRD", new List<string> { "ROUTEs", "PermanentRouteCancellation", "R", "RC" });
                equip.Add("GSBL", new List<string> { "ASCVs", "GlobalSignalBlock","","RM" });
                equip.Add("ISN", new List<string> { "POINTs", "InhibitionState","P","RM" });
                equip.Add("LDBS", new List<string> { "SIGNALs", "BufferStopLampProved","S","RM" });
                equip.Add("LDRI", new List<string> { "SIGNALs", "RouteIndicatorLampProved","S","RM" });
                equip.Add("LDS_G", new List<string> { "SIGNALs", "GreenLampProved","S","RM" });
                equip.Add("LDS_VR", new List<string> { "SIGNALs", "VioletRedLampProved","S","RM" });
                equip.Add("LDSS_1", new List<string> { "SIGNALs", "OneAspects", "S","RM" });
                equip.Add("LDSS_2", new List<string> { "SIGNALs", "TwoAspects", "S","RM" });
                equip.Add("MAGP", new List<string> { "SWITCHs", "AuthorisationStatus","SW","RM" });
                equip.Add("MAP", new List<string> { "SWITCHs", "ManualAuthorizationControl","SW","RC" });
                equip.Add("MARP", new List<string> { "SWITCHs", "ManualAuthorizationRelease", "SW", "RC" });
                equip.Add("MBL", new List<string> { "MBLs", "status","MBL","RM" });
                equip.Add("MPL", new List<string> { "SWITCHs", "BlockingStatus","SW","RM" });
                equip.Add("MPS1", new List<string> { "ASCVs", "PowerSupply1State","","RM" });
                equip.Add("MPS2", new List<string> { "ASCVs", "PowerSupply2State","","RM" });
                equip.Add("OL", new List<string> { "OVERLAPs", "Status","OL","RM" });
                equip.Add("PCN", new List<string> { "SWITCHs", "CallingNormal","SW","RC" });
                equip.Add("PCR", new List<string> { "SWITCHs", "CallingReverse","SW","RC" });
                equip.Add("PMCKS", new List<string> { "SWITCHs", "KeyStatus","SW","RM" });
                equip.Add("PMCK", new List<string> { "SWITCHs", "KeyStatus","SW","RM" });
                equip.Add("PPN", new List<string> { "SWITCHs", "CalledNormal","SW","RM" });
                equip.Add("PPR", new List<string> { "SWITCHs", "CalledReverse","SW","RM" });
                equip.Add("PREP_ASN", new List<string> { "POINTs", "SelfNormalisationActivationPreparation", "P", "RC" });
                equip.Add("PREP_ASNA", new List<string> { "ASCVs", "ActivationAck","","RM" });
                equip.Add("PREP_ISN", new List<string> { "POINTs", "SelfNormalisationInhibitionPreparation","P","RC" });
                equip.Add("PREP_ISNA", new List<string> { "ASCVs", "InhibitionAck","","RM" });
                equip.Add("PREP_GMPLD", new List<string> { "ASCVs", "GlobalPointUnblockingPreparation", "","RC" });
                equip.Add("PREP_GRUBL", new List<string> { "ASCVs", "GlobalRouteUnblockingPreparation","R","RC" });
                equip.Add("CONF_GRUBL", new List<string> { "ASCVs", "GlobalRouteUnblockingConfirmation", "R","RC" });
                equip.Add("PREP_GRUBLA", new List<string> { "ASCVs", "GlobalRouteUnblockingAck", "","RM" });
                equip.Add("CONF_GMBUL", new List<string> { "ASCVs", "GlobalMaintenanceBlockUnblockingConfirmation", "","RC" });
                equip.Add("PREP_GMBUL", new List<string> { "ASCVs", "GlobalMaintenanceBlockUnblockingPreparation", "","RC" });
                equip.Add("CONF_GMPLD", new List<string> { "ASCVs", "GlobalPointUnblockingConfirmation", "","RC" });
                equip.Add("PREP_GMPLDA", new List<string> { "ASCVs", "GlobalPointUnblockingAck", "","RM" });
                equip.Add("PREP_GSBLA", new List<string> { "ASCVs", "SignalBlockingAck","","RM" });
                equip.Add("PREP_GSUBLA", new List<string> { "ASCVs", "SignalReleaseAck","", "RM" });
                equip.Add("PREP_GSUBL", new List<string> { "ASCVs", "GlobalSignalUnblockingPreparation", "", "RC" });
                equip.Add("CONF_GSUBL", new List<string> { "ASCVs", "GlobalSignalUnblockingConfirmation", "", "RC" });
                equip.Add("CONF_GSBL", new List<string> { "ASCVs", "GlobalSignalBlockingConfirmation", "", "RC" });
                equip.Add("PREP_GSBL", new List<string> { "ASCVs", "GlobalSignalBlockingPreparation", "", "RC" });
                equip.Add("PREP_MBL", new List<string> { "MBLs", "MaintenanceBlockControlPreparation","MBL","RC" });
                equip.Add("PREP_MBLA", new List<string> { "ASCVs", "SettingAck","", "RM" });
                equip.Add("PREP_MBUL", new List<string> { "MBLs", "MaintenanceBlockReleasePreparation", "MBL", "RC" });
                equip.Add("PREP_MBULA", new List<string> { "ASCVs", "ReleaseAck","", "RM" });
                equip.Add("PREP_MPL", new List<string> { "SWITCHs", "SwitchBlockingPreparation","SW","RC" });
                equip.Add("PREP_MPLA", new List<string> { "ASCVs", "PointBlockingAck","", "RM" });
                equip.Add("PREP_MPLD", new List<string> { "SWITCHs", "SwitchUnblockingPreparation","SW","RC" });
                equip.Add("PREP_MPLDA", new List<string> { "ASCVs", "PointUnblockingAck","", "RM" });
                equip.Add("PREP_RBL", new List<string> { "ROUTEs", "RouteControlPreparation", "R", "RC" });
                equip.Add("PREP_RBLA", new List<string> { "ASCVs", "RouteBlockingAck","", "RM" });
                equip.Add("PREP_RUBL", new List<string> { "ROUTEs", "RouteReleasePreparation", "R", "RC" });
                equip.Add("PREP_RUBLA", new List<string> { "ASCVs", "RouteUnblockingAck","", "RM" });
                equip.Add("PREP_SBL", new List<string> { "SIGNALs", "SignalBlockingPreparation","S","RC" });
                equip.Add("PREP_SBLA", new List<string> { "ASCVs", "SignalBlockingAck","", "RM" });
                equip.Add("PREP_SUBL", new List<string> { "SIGNALs", "SignalUnblockingPreparation", "S", "RC" });
                equip.Add("PREP_SUBLA", new List<string> { "ASCVs", "SignalReleaseAck","", "RM" });
                equip.Add("PSAPR", new List<string> { "PSAPRs", "SecondaryStationPowerSupplyState","SST_", "RM" });
                equip.Add("PSC", new List<string> { "ASCVs", "PowerSupplyAlarmState","", "RM" });
                equip.Add("RBL", new List<string> { "ROUTEs", "RouteBlocking","R", "RM" });
                equip.Add("RC", new List<string> { "ROUTEs", "RouteControl","R","RC" });
                equip.Add("RD", new List<string> { "ROUTEs", "RouteCancellation", "R", "RC" });
                equip.Add("RL_S", new List<string> { "ROUTEs", "RouteSetLocked","R", "RM" });
                equip.Add("S_OFF_V", new List<string> { "SIGNALs", "SignalOFFViolet","", "RM" });
                equip.Add("SS_OFF", new List<string> { "SIGNALs", "ShuntOFF","S", "RM" });
                equip.Add("S_ON", new List<string> { "SIGNALs", "SignalON","S", "RM" });
                equip.Add("SBL", new List<string> { "SIGNALs", "ToONStatus", "S", "RM" });
                equip.Add("TC", new List<string> { "SECONDARIEs", "Status", "T", "RM" });
                equip.Add("TD", new List<string> { "TRAFFICDIRECTIONs", "Status","TD_", "RM" });
                equip.Add("TDR", new List<string> { "ASCVs", "TDR_"+ Sector_Code, "TDR", "RM" });
                equip.Add("U", new List<string> { "SUBROUTEs", "Status","SB", "RM" });
                equip.Add("UPS1", new List<string> { "ASCVs", "UninterruptedPowerSupply1State","", "RM" });
                equip.Add("UPS2", new List<string> { "ASCVs", "UninterruptedPowerSupply2State","", "RM" });
                equip.Add("VMC", new List<string> { "ASCVs", "VDUModeControlState","", "RM" });
                string Access_Dict(string key, int index)
                {
                    if (equip.ContainsKey(key) && index >= 0 && index < equip[key].Count)
                    {
                        return equip[key][index];
                    }
                    return null;

                }
                string Get_String(string input)
                {
                    int numberIndex = -1;
                    if (input.Split('_')[0] == "PSAPR")
                        return input.Split('_')[1];
                    if(input.Substring(0,4)=="LDSS")
                    {
                        return (input.Substring(4, 3));
                    }
                    if (input.Substring(0,3) == "TDR")
                        return "";
                    if(input.Substring(0,2)=="SS")
                    {
                        return (input.Substring(2, 3));
                    }
                    if(input.Substring(0, 3) == "MAP")
                    {
                        return input.Substring(3, input.Length - 3);
                    }
                    if (input.Substring(0, 4) == "MARP")
                    {
                        return input.Substring(4, input.Length - 4);
                    }
                    if (input.Substring(0, 4) == "MAGP")
                    {
                        return input.Substring(4, input.Length - 4);
                    }
                    if (input.Substring(0, 4) == "PMCK")
                    {
                        return input.Substring(4, input.Length - 4);
                    }
                    if(input.Substring(0, 3) == "UPS" || input.Substring(0, 3) == "MPS")
                    {
                        return input[input.Length - 1].ToString();
                    }

                    for (int i = 0; i < input.Length; i++)
                    {
                        if (char.IsNumber(input[i]))
                        {
                            numberIndex = i;
                            break;
                        }
                    }
                    if (numberIndex == -1)
                    {
                        return input[input.Length - 1].ToString();
                    }
                    int firstUnderscore = input.IndexOf('_', numberIndex);
                    if (firstUnderscore != -1)
                    {
                        if (input.Split('_')[1] == "ON")
                        {
                            return input.Split('_')[0];
                        }
                    }
                    int secondUnderscore = input.IndexOf('_', firstUnderscore + 1);

                    if (secondUnderscore == -1)
                    {
                        return input.Substring(numberIndex);
                    }
                    
                    else
                    {
                        if (input.Split('_')[1] == "OFF")
                        {
                            return input.Split('_')[0];
                        }
                        if (secondUnderscore != -1 && input[input.Length - 1] == 'S' && input.Substring(2)=="RL")
                            return input.Substring(numberIndex, secondUnderscore - numberIndex);
                        else
                            return input.Substring(numberIndex);
                    }
                }
                string check(string a)
                {
                    //Console.WriteLine(a);
                    //string[] l = a.Split('_');
                    if (a.Substring(0, 3) == "MBL")
                    {
                        a = "MBL_" + a;
                    }
                    else if (a.Substring(0, 2) == "OL")
                    {
                        a = "OL_" + a;
                    }
                    int Find_US = -1;
                    Find_US = a.IndexOf('_');
                    if (Find_US != -1)
                    {
                        if (a.Split('_')[1] == "ON")
                        {
                            a = a.Split('_')[0] + "_ON";
                        }
                    }
                    return a;
                }
                    using (StreamWriter sw = fi.AppendText())
                    {
                        sw.WriteLine("<?xml version=\"1.0\" ?>");
                        sw.WriteLine("<ObjectPropertyModule Generation_Date=\""+DateTime.Now.ToString()+"\" Project=\"DELHI\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xsi:noNamespaceSchemaLocation=\"Interlocking.xsd\">");
                        sw.WriteLine("  <Versions>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#Comment\" Value=\"XML file containing the System Data\"/>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#Date\" Value=\"2011-03-19\"/>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#SyPD_Version\" Value=\"_none_\"/>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#UEVOLtoIDP_Version\" Value=\"2.1.7\"/>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#Writer_Name\" Value=\"IconisDigesterCore 0.0.4091.29806\"/>");
                        sw.WriteLine("    <Version Name=\"ATS_SYSTEM_PARAMETERS#XML_File_Version\" Value=\"1.6.6.1078\"/>");
                        sw.WriteLine("    <Version Name=\"ICONIS_ATS_Equipment#SyID_Version\" Value=\"Y3-64 A427945-J1\"/>");
                        sw.WriteLine("    <Version Name=\"ICONIS_IDP#Customisation_SwRSAD_Version\" Value=\"none\"/>");
                        sw.WriteLine("    <Version Name=\"ICONIS_IDP#Product_SwRSAD_Version\" Value=\"Y3-64 A427875-A\"/>");
                        sw.WriteLine("    <Version Name=\"ICONIS_IDP#Product_Version\" Value=\"10.3.4 (revision 3010)\"/>");
                        sw.WriteLine("    <Version Name=\"ICONIS_IDP#Project_Version\" Value=\"BLR#V10.3.4 (revision 3060)\"/>");
                        sw.WriteLine("    <Version Name=\"Project#ATS_Custom_Reference_Database_Version\" Value=\"0.0.3.565\"/>");
                        sw.WriteLine("    <Version Name=\"Project#Database_Version\" Value=\"0.0.3.565\"/>");
                        sw.WriteLine("  </Versions>");
                        sw.WriteLine(" <Classes> ");
                        sw.WriteLine(" <Class name=\"Interlocking\">");
                        sw.WriteLine("      <Objects>");
                        sw.WriteLine("        <Object name=\"ASCV_BDLD\" rules=\"update_or_create\">");
                        sw.WriteLine("          <Properties>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"ID\">ASCV_"+ xlRange.Cells[1, 1].Value2.ToString()+ "</Property>");
                        sw.WriteLine("            <Property dt=\"boolean\" name=\"LinkAlarm_EnableInstance\">1</Property>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"OPCClientID\">OPCClient_ASCV_"+xlRange.Cells[1, 1].Value2.ToString()+ "</Property>");
                        sw.WriteLine("            <MultiLingualProperty name=\"Name\">");
                        sw.WriteLine("              <MultiLingualValue localeId=\"1033\" roleId=\" - 1\">ASCV"+ xlRange.Cells[1, 2].Value2.ToString() + "</MultiLingualValue>");
                        sw.WriteLine("              <MultiLingualValue localeId=\"1036\" roleId=\" - 1\">ASCV"+ xlRange.Cells[1, 2].Value2.ToString() + "</MultiLingualValue>");
                        sw.WriteLine("            </MultiLingualProperty>");
                        sw.WriteLine("          </Properties>");
                        sw.WriteLine("        </Object>");
                        sw.WriteLine("      </Objects>");
                        sw.WriteLine("    </Class>");
                        sw.WriteLine("    <Class name=\"OPCClient\">");
                        sw.WriteLine("      <Objects>");
                        sw.WriteLine("        <Object name=\"OPCClient_ASCV_BDLD\" rules=\"update_or_create\">");
                        sw.WriteLine("          <Properties>");
                        sw.WriteLine("            <Property dt=\"i4\" name=\"ConnectAttemptsPeriod\">10</Property>");
                        sw.WriteLine("            <Property dt=\"boolean\" name=\"DoubleLinks\">0</Property>");
                        sw.WriteLine("            <Property dt=\"i4\" name=\"MonitoringPeriod\">30</Property>");
                        sw.WriteLine("            <Property dt=\"boolean\" name=\"MultiActive\">1</Property>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"ServerNodeName1\">localhost</Property>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"ServerNodeName2\">localhost</Property>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"ServerProgID1\">RDD.1.RDDASCV_" + xlRange.Cells[1, 1].Value2.ToString()+ "</Property>");
                        sw.WriteLine("            <Property dt=\"string\" name=\"ServerProgID2\">RDD.1.RDDASCV_" + xlRange.Cells[1, 1].Value2.ToString()+ "</Property>");
                        sw.WriteLine("          </Properties>");
                        sw.WriteLine("        </Object>");
                        sw.WriteLine("      </Objects>");
                        sw.WriteLine("    </Class>");
                        sw.WriteLine("    <Class name=\"Variable\">");
                        sw.WriteLine("      <Objects>");
                    }
                
                for (int i = 2; i <= rowCount; i++)
                {
                    
                    if (xlRange.Cells[i, 1] != null && xlRange.Cells[i, 1].Value2 != null)
                    {
                        int.TryParse(xlRange.Cells[i, 3].Value2.ToString(), out Address);
                        Address++;

                        try
                            {

                                
                                using (StreamWriter sw = fi.AppendText())
                                {
                                    sw.WriteLine("        <Object name=\"" + check(xlRange.Cells[i, 2].Value2.ToString()) + "\" rules=\"update_or_create\">");
                                    sw.WriteLine("          <Properties>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"Address\">" + Address.ToString() + "</Property>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"InterlockingID\">ASCV_"+Sector+"</Property>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"OPCGroup\">MyGroup</Property>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"OPCTag\">" + Access_Dict(xlRange.Cells[i, 1].Value2.ToString(),0) +"." + Access_Dict(xlRange.Cells[i, 1].Value2.ToString(),2) + Get_String(xlRange.Cells[i, 2].Value2.ToString()) + "."+ Access_Dict(xlRange.Cells[i, 1].Value2.ToString(), 3) + "." + Access_Dict(xlRange.Cells[i, 1].Value2.ToString(),1) +"</Property>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"OPCType\">VT_BOOL</Property>");
                                    sw.WriteLine("            <Property dt=\"string\" name=\"Type\">"+ Access_Dict(xlRange.Cells[i, 1].Value2.ToString(), 3) + "</Property>");
                                    sw.WriteLine("          </Properties>");
                                    sw.WriteLine("        </Object>");

                                }
                            }
                            catch (Exception Ex)
                            {
                                Console.WriteLine(Ex.ToString());

                            }
                        

                        
                            
                        

                    }
                }

                using (StreamWriter sw = fi.AppendText())
                {
                    sw.WriteLine("</Objects>");
                    sw.WriteLine("</Class>");
                    sw.WriteLine("</Classes>");
                    sw.WriteLine("</ObjectPropertyModule>");
                }







                GC.Collect();
                        GC.WaitForPendingFinalizers();

                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);

                        //xlWorkbook.Close();
                        //Marshal.ReleaseComObject(xlWorkbook);

                        //xlApp.Quit();
                        //Marshal.ReleaseComObject(xlApp);

                    
                
            }
            if(S2KFunctional.Checked == true)
            {

                string fileName = @selectedPath + @"\S2KFunctional.XML";
                //string fileName = @"D:\U200_IDP\SOA1_Point.xml";
                FileInfo fi = new FileInfo(fileName);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[14];
                Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                List<List<string>> data = new List<List<string>>();
                for (int row = 1; row <= rowCount; row++)
                {
                    List<string> rowData = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        if ((xlRange.Cells[row, col] as Range).Value2 == null)
                        {
                            rowData.Add("");
                        }
                        else
                        {
                            rowData.Add((xlRange.Cells[row, col] as Range).Value2.ToString());
                        }
                    }
                    data.Add(rowData);
                }
                data = data.OrderBy(rowData => rowData[0]).ToList();
                if (fi.Exists)
                {
                    fi.Delete();
                }
                Dictionary<string,List<string>> Obj = new Dictionary<string, List<string>>();
                Obj.Add("CY", new List<string> {"Cycle","C","CY"});
                Obj.Add("ESP", new List<string> { "ESP","ESP_", "ESP_" });
                Obj.Add("MBL", new List<string> { "MaintenanceBlock", "MBL","MBL"});
                Obj.Add("DPN", new List<string> {"Point","P","P"});
                Obj.Add("PSAPR", new List<string> {"PSAPR","SST_","SST_"});
                Obj.Add("RL_S", new List<string> {"Route","R","RT"});
                Obj.Add("S_ON", new List<string> {"Signal","",""});
                Obj.Add("U", new List<string> { "Sub_Route", "SB","SB"});
                Obj.Add("PPN", new List<string> { "Switch", "SW","SW"});
                Obj.Add("TC", new List<string> { "TrackCircuit", "T","TC"});

                List<string> lst = new List<string>();

                string Access_Dict(string key, int index)
                {
                    if (Obj.ContainsKey(key) && index >= 0 && index < Obj[key].Count)
                    {
                        return Obj[key][index];
                    }
                    return null;

                }
                string Get_String(string input)
                {
                    int numberIndex = -1;
                    if (input.Split('_')[0] == "PSAPR")
                        return input.Split('_')[1];

                    for (int i = 0; i < input.Length; i++)
                    {
                        if (char.IsNumber(input[i]))
                        {
                            numberIndex = i;
                            break;
                        }
                    }
                    if (numberIndex == -1)
                    {
                        return input[input.Length - 1].ToString();
                    }
                    int firstUnderscore = input.IndexOf('_', numberIndex);
                    if (firstUnderscore != -1)
                    {
                        if (input.Split('_')[1] == "ON")
                        {
                            return input.Split('_')[0];
                        }
                    }
                    int secondUnderscore = input.IndexOf('_', firstUnderscore + 1);

                    if (secondUnderscore == -1)
                    {
                        return input.Substring(numberIndex);
                    }

                    else
                    {
                        if (input.Split('_')[1] == "OFF")
                        {
                            return input.Split('_')[0];
                        }
                        if (secondUnderscore != -1 && input[input.Length - 1] == 'S')
                            return input.Substring(numberIndex, secondUnderscore - numberIndex);
                        else
                            return input.Substring(numberIndex);
                    }
                }
                using (StreamWriter sw = fi.CreateText())
                {
                    
                }

                for (int i = 0; i <= rowCount-2; i++)
                {
                    if (data[i][0] != null)
                    {
                       

                            try
                            {


                            using (StreamWriter sw = fi.AppendText())
                            {
                                
                                if (Obj.TryGetValue(data[i][0], out List<string> value))
                                {
                                    if (lst.Count == 0)
                                    {
                                        sw.WriteLine("<Grp GrpName=\"Function\">");
                                        sw.WriteLine("<Grp GrpObj=\"Function/Signalling\" GrpName=\"ASCV\">");
                                        sw.WriteLine("<Properties GrpObj=\"Function/Signalling/ASCV\">");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"ASCV\" LocaleID=\"1033\" RoleID=\"-1\"/>");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"ASCV\" LocaleID=\"1036\" RoleID=\"-1\"/>");
                                        sw.WriteLine("</Properties>");
                                        sw.WriteLine("<Grp GrpObj=\"Function/Signalling/ASCV\" GrpName=\"" + xlRange.Cells[1, 2].Value2.ToString() + "\">");
                                        sw.WriteLine("<Properties GrpObj=\"Function/Signalling/ASCV/" + xlRange.Cells[1, 2].Value2.ToString() + "\">");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + xlRange.Cells[1, 2].Value2.ToString() + "\" LocaleID=\"1033\" RoleID=\"-1\"/>");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + xlRange.Cells[1, 2].Value2.ToString() + "\" LocaleID=\"1036\" RoleID=\"-1\"/>");
                                        sw.WriteLine("</Properties>");
                                        sw.WriteLine("</Grp>");
                                        sw.WriteLine("</Grp>");
                                    }

                                    if (lst.Contains(data[i][0]) ==false)
                                    {
                                        sw.WriteLine("<Grp GrpObj=\"Function/Signalling\" GrpName=\"" + Access_Dict(data[i][0], 0) + "\">");
                                        sw.WriteLine("<Properties GrpObj=\"Function/Signalling/" + Access_Dict(data[i][0], 0) + "\">");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + Access_Dict(data[i][0], 0) + "\" LocaleID=\"1033\" RoleID=\"-1\"/>");
                                        sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + Access_Dict(data[i][0], 0) + "\" LocaleID=\"1036\" RoleID=\"-1\"/>");
                                        sw.WriteLine("</Properties>");
                                    }
                                    sw.WriteLine("<Grp GrpObj=\"Function/Signalling/" + Access_Dict(data[i][0], 0) + "\" GrpName=\"" + Access_Dict(data[i][0], 1) + Get_String(data[i][1]) + "\">");
                                    sw.WriteLine("<Properties GrpObj=\"Function/Signalling/" + Access_Dict(data[i][0], 0) +"/" + Access_Dict(data[i][0], 1) + Get_String(data[i][1]) + "\">");
                                    sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + Access_Dict(data[i][0], 2) + Get_String(data[i][1]) + "\" LocaleID=\"1033\" RoleID=\"-1\"/>");
                                    sw.WriteLine("<Property PrpName=\"Name\" PrpValue=\"" + Access_Dict(data[i][0], 2) + Get_String(data[i][1]) + "\" LocaleID=\"1036\" RoleID=\"-1\"/>");
                                    sw.WriteLine("</Properties>");
                                    sw.WriteLine("</Grp>");
                                    if (data[i][0] != data[i+1][0] && lst.Contains(data[i+1][0]) ==false)
                                    {
                                        Console.WriteLine("1");
                                        sw.WriteLine("</Grp>");
                                    }
                                    if(i==rowCount-3)
                                    {
                                        Console.WriteLine("2");
                                        sw.WriteLine("</Grp>");
                                    }
                                    lst.Add(data[i][0]);

                                }
                            }
                            }
                            catch (Exception Ex)
                            {
                                Console.WriteLine(Ex.ToString());

                            }
                        

                        
                            
                        

                    }

                }









                GC.Collect();
                        GC.WaitForPendingFinalizers();

                        Marshal.ReleaseComObject(xlRange);
                        Marshal.ReleaseComObject(xlWorksheet);

                        //xlWorkbook.Close();
                        //Marshal.ReleaseComObject(xlWorkbook);

                        //xlApp.Quit();
                        //Marshal.ReleaseComObject(xlApp);

                    
                
            }
            string ASCVf = @selectedPath + @"\All ASCVs.xml";
            FileInfo ASC = new FileInfo(ASCVf);
            if (ASC.Exists) { ASC.Delete();}
            int im = 0;


            using (StreamWriter writer = new StreamWriter(ASCVf))
            {
                

                foreach (string file in Out_Filename)
                {
                    if(file == "Interlocking" || file == "S2KFunctional")
                    {
                        continue;
                    }
                    FileInfo Afile = new FileInfo(@selectedPath + @"\" + file + ".xml");
                    if (Afile.Exists)
                    {
                        using (StreamReader reader = new StreamReader(@selectedPath + @"\" + file + ".xml"))
                        {
                            string line;
                            while ((line = reader.ReadLine()) != null)
                            {
                                // Skip lines starting with "New file created:" or "Author:"
                                if (line.StartsWith("New file created:") || line.StartsWith("Author:"))
                                {
                                    continue;
                                }
                                if(im++ == 0)
                                {
                                    writer.WriteLine("<CBIS>");
                                }
                                writer.WriteLine(line);
                            }
                        }

                    }

                }
            }
            using (StreamWriter sw = ASC.AppendText())
            {
                sw.WriteLine("</CBIS>");
            }
            MessageBox.Show("Program Completed.", "INFO", MessageBoxButtons.OK, MessageBoxIcon.Information);
            GC.Collect();
            GC.WaitForPendingFinalizers();


            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlWorkbook = null;

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            xlApp = null;




        }



        private void button2_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            DialogResult result = openFileDialog1.ShowDialog(); 
            if (result == DialogResult.OK) 
            {
                
                string file = openFileDialog1.FileName;
                try
                {
                    textBox1.Text = file;
                }
                catch (IOException Ex)
                {
                    Console.WriteLine(Ex.ToString());
                }
                

            }
            else
            {
                MessageBox.Show("Select input file to continue", "CAUTION", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
            foreach (Worksheet sheet in xlWorkbook.Sheets)
            {
                Out_Filename.Add(sheet.Name);
                files.Add(@selectedPath + @"\" + Out_Filename + ".xml");
            }

            this.TrackCircuit.Text = Out_Filename[0];
            //this.S2KFunctional.Text = Out_Filename[0];
            this.Interlocking.Text = Out_Filename[13];
            this.ASCV.Text = Out_Filename[12];
            this.Switch.Text = Out_Filename[11];
            this.Route.Text = Out_Filename[10];
            this.Cycle.Text = Out_Filename[8];
            this.Overlap.Text = Out_Filename[6];
            this.MBL.Text = Out_Filename[9];
            this.PSAPR.Text = Out_Filename[7];
            this.Signal.Text = Out_Filename[5];
            this.Point.Text = Out_Filename[4];
            this.ESP.Text = Out_Filename[3];
            this.TrafficDirection.Text = Out_Filename[2];
            this.Subroute.Text = Out_Filename[1];
            button3.Enabled = true;

            ToolTip toolTip1 = new ToolTip();
            toolTip1.SetToolTip(button3, "Click this button to choose output folder");

            /*Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
            Excel.Sheets sheets = xlApp.Worksheets;
            Excel.Worksheet worksheet = (Excel.Worksheet)sheets.get_Item(2);//Get the reference of second worksheet
            string strWorksheetName = worksheet.Name;//Get the name of worksheet.
            int sheetcount = xlWorkbook.Sheets.Count;

            CheckBox box;
            for(int i = 1; i<=sheetcount; i++)
            {
                worksheet = worksheet = (Excel.Worksheet)sheets.get_Item(i);
                strWorksheetName = worksheet.Name;
                box = new CheckBox();
                box.Tag = i.ToString();
                box.Text = strWorksheetName;
                box.AutoSize = true;
                //box.Location = new Point(10, i*50); //vertical
                box.Location = new Point(i * 120, 155); //horizontal
                this.Controls.Add(box);
            }*/



        }


        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.Description = "Select a folder to save the file.";

            DialogResult result = folderDialog.ShowDialog();

            if (result == DialogResult.OK)
            {
                selectedPath = folderDialog.SelectedPath;
                panel1.Visible = true;
                
                button4.Visible = true;
                button5.Visible = true;

            }
            else
            {
                MessageBox.Show("You should select output file location to continue", "CAUTION", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            
                

            
            


        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void TrackCircuit_CheckedChanged(object sender, EventArgs e)
        {
            
            if(TrackCircuit.Checked==true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false &&ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }

        }

        private void Subroute_CheckedChanged(object sender, EventArgs e)
        {
            if (Subroute.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }

        private void TrafficDirection_CheckedChanged(object sender, EventArgs e)
        {
            if (TrafficDirection.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }

        }

        private void ESP_CheckedChanged(object sender, EventArgs e)
        {
            if (ESP.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }

        private void Point_CheckedChanged(object sender, EventArgs e)
        {
            if (Point.Checked == true)
            { 
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (Control ctl in this.panel1.Controls)

            {

                if (ctl.GetType() == typeof(CheckBox))

                {

                    ((CheckBox)ctl).Checked = true;



                }

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            foreach (Control ctl in this.panel1.Controls)

            {

                if (ctl.GetType() == typeof(CheckBox))

                {

                    ((CheckBox)ctl).Checked = false;



                }

            }
        }

        private void Overlap_CheckedChanged(object sender, EventArgs e)
        {
            if (Overlap.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
                button5.Enabled = false;
                button5.Enabled = true;
            }
        }
        private void PSAPR_CheckedChanged(object sender, EventArgs e)
        {
            if (PSAPR.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }
        private void MBL_CheckedChanged(object sender, EventArgs e)
        {
            if (MBL.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }
        private void Route_CheckedChanged(object sender, EventArgs e)
        {
            if (Route.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }
        private void Switch_CheckedChanged(object sender, EventArgs e)
        {
            if (Switch.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }

        }
        private void ASCV_CheckedChanged(object sender, EventArgs e)
        {
            if (ASCV.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }

        }
        private void Cycle_CheckedChanged(object sender, EventArgs e)
        {
            if (Cycle.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }

        private void Signal_CheckedChanged(object sender, EventArgs e)
        {
            if(Signal.Checked==true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }

        private void Interlocking_CheckedChanged(object sender, EventArgs e)
        {
            if (Interlocking.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }
        
        private void S2KFunctional_CheckedChanged(object sender, EventArgs e)
        {
            if (S2KFunctional.Checked == true)
            {
                button1.Enabled = true;
            }
            if (TrackCircuit.Checked == false && Subroute.Checked == false && TrafficDirection.Checked == false && ESP.Checked == false && Point.Checked == false && Signal.Checked == false && PSAPR.Checked == false && Overlap.Checked == false && Cycle.Checked == false && MBL.Checked == false && Route.Checked == false && Switch.Checked == false && ASCV.Checked == false && Interlocking.Checked == false && S2KFunctional.Checked == false)
            {
                button1.Enabled = false;
            }
        }
    }
}

