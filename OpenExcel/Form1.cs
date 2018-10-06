using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Runtime.InteropServices.ComTypes;

namespace OpenExcel
{
    public partial class Form1 : Form
    {
        // Contains the path to the workbook file
        private string m_ExcelFileName = @"C:\Users\Linh\source\repos\OpenExcel\OpenExcel\bin\Debug\DanhSachGiaoVien.xls"; // Replace here with an existing file

        // Contains a reference to the hosting application
        private Microsoft.Office.Interop.Excel.Application m_XlApplication = null;
        // Contains a reference to the active workbook
        private Workbook m_Workbook = null;

        public Form1()
        {
            InitializeComponent();
        }

        public void OpenFile(string filename)
        {
            // Check the file exists
            if (!System.IO.File.Exists(filename)) throw new Exception();
            m_ExcelFileName = filename;
            // Load the workbook in the WebBrowser control
            this.webBrowser1.Navigate(filename, false);
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            // Creation of the workbook object
            if ((m_Workbook = RetrieveWorkbook(m_ExcelFileName)) == null) return;
            // Create the Excel.Application
            m_XlApplication = (Microsoft.Office.Interop.Excel.Application)m_Workbook.Application;
        }

        [DllImport("ole32.dll")]
        static extern int GetRunningObjectTable
                (uint reserved, out IRunningObjectTable pprot);
        [DllImport("ole32.dll")] static extern int CreateBindCtx(uint reserved, out IBindCtx pctx);

        public Workbook RetrieveWorkbook(string xlfile)
        {
            IRunningObjectTable prot = null;
            IEnumMoniker pmonkenum = null;
            try
            {
                IntPtr pfetched = IntPtr.Zero;
                // Query the running object table (ROT)
                if (GetRunningObjectTable(0, out prot) != 0 || prot == null) return null;
                prot.EnumRunning(out pmonkenum); pmonkenum.Reset();
                IMoniker[] monikers = new IMoniker[1];
                while (pmonkenum.Next(1, monikers, pfetched) == 0)
                {
                    IBindCtx pctx; string filepathname;
                    CreateBindCtx(0, out pctx);
                    // Get the name of the file
                    monikers[0].GetDisplayName(pctx, null, out filepathname);
                    // Clean up
                    Marshal.ReleaseComObject(pctx);
                    // Search for the workbook
                    if (filepathname.IndexOf(xlfile) != -1)
                    {
                        object roval;
                        // Get a handle on the workbook
                        prot.GetObject(monikers[0], out roval);
                        return roval as Workbook;
                    }
                }
            }
            catch
            {
                return null;
            }
            finally
            {
                // Clean up
                if (prot != null) Marshal.ReleaseComObject(prot);
                if (pmonkenum != null) Marshal.ReleaseComObject(pmonkenum);
            }
            return null;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFile(m_ExcelFileName);
        }

        //protected override void OnClosed(object sender, EventArgs e)
        //{
        //    try
        //    {
        //        // Quit Excel and clean up.
        //        if (m_Workbook != null)
        //        {
        //            m_Workbook.Close(true, Missing.Value, Missing.Value);
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject
        //                                    (m_Workbook);
        //            m_Workbook = null;
        //        }
        //        if (m_XlApplication != null)
        //        {
        //            m_XlApplication.Quit();
        //            System.Runtime.InteropServices.Marshal.ReleaseComObject
        //                                (m_XlApplication);
        //            m_XlApplication = null;
        //            System.GC.Collect();
        //        }
        //    }
        //    catch
        //    {
        //        MessageBox.Show("Failed to close the application");
        //    }
        //}
    }
}
