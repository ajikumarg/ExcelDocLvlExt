using System;
using System.Data;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWrkBk
{
    public partial class SheetM61CF
    {

        //Fields
        Dictionary<string, string> CFAssumptions;


        private void Sheet45_Startup(object sender, System.EventArgs e)
        {
            CFAssumptions = new Dictionary<string, string>();
            //ReadDocument();
            read_Excel_File();
            //Check the  NoteID and the Closing Date


        }

        private void Sheet45_Shutdown(object sender, System.EventArgs e)
        {
        }

        public void ReadDocument()
        {
            int colFldName = 1, colValue = 2;
            object rngStart, rngEnd; 
            rngStart = this.M61FieldName.Row; 
            rngEnd = this.M61RateSpreadScheduleList.Row;

            //for (int rownum = rngStart + 1; rownum < rngEnd; rownum++)
            //{

            //    CFAssumptions.Add(this.Range[rownum, colFldName].Value, this.Range[rownum, colValue].Value);

            //}
        }


        public void read_Excel_File()
        {
            //----------------< read_Excel_File_into_DataGridView() >------------
            //</ init >

            Excel.Range usedRange = this.UsedRange;
            int nColumnsMax = 0;

            if (usedRange.Rows.Count > 0)
            {
                //----< Read_Header >----
                for (int iColumn = 1; iColumn <= usedRange.Columns.Count; iColumn++)
                {
                    Excel.Range cell = usedRange.Cells[1, iColumn] as Excel.Range;
                    String sValue = cell.Value2.ToString();

                    if (sValue == "") break;
                }
                //----</ Read_Header >----

                //----< Read_DataRows >----
                for (int iRow = 2; iRow <= usedRange.Rows.Count; iRow++)
                {
                    for (int iColumn = 1; iColumn <= nColumnsMax; iColumn++)
                    {
                        Microsoft.Office.Interop.Excel.Range cell = usedRange.Cells[iRow, iColumn] as Excel.Range;
                        String sValue = cell.Value2.ToString();
                    }
                }
                //----</ Read_DataRows >----
            }

            //----------------</ read_Excel_File_into_DataGridView() >------------
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet45_Startup);
            this.Shutdown += new System.EventHandler(Sheet45_Shutdown);
        }

        #endregion

    }
}
