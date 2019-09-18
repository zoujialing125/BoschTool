using System;
using System.IO;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace BoschTool
{
    public partial class Sheet1
    {
        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            Cells.Delete();
            Cells[1, 1] = "Available Files";
            Cells[1, 2] = "CI No.";
            Cells[1, 3] = "Shipment Name";
            Cells[1, 4] = "Trade Type";
            Cells[1, 5] = "Process Status";
            Columns["A:E"].ColumnWidth = 20;
            Columns["A:E"].NumberFormat = "@";
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(this.Sheet1_Startup);
            this.Shutdown += new System.EventHandler(this.Sheet1_Shutdown);

        }

        #endregion

    }
}
