using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;

namespace ExcelAddInDemo
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            #region Demo code
            /******** Demo code for understanding events associated with the Excel Application ********/
            // this.Application.WorkbookBeforeSave += new Microsoft.Office.Interop.Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave); 
            #endregion
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Demo code
        /******** Demo code for understanding events associated with the Excel Application ********/
        //void Application_WorkbookBeforeSave(Microsoft.Office.Interop.Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        //{
        //    Excel.Worksheet activeWorksheet = ((Excel.Worksheet)Application.ActiveSheet);
        //    Excel.Range firstRow = activeWorksheet.get_Range("A1");
        //    firstRow.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
        //    Excel.Range newFirstRow = activeWorksheet.get_Range("A1");
        //    newFirstRow.Value2 = "This text was added by using code";
        //} 
        #endregion

        #region XmlRibbon
        //protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        //{
        //    return new XmlRibbon();
        //}
        #endregion

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
