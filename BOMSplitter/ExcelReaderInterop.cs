using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Data;

namespace BOMSplitter
{
    class ExcelReaderInterop
    {
        /// Store the Application object we can use in the member functions.
        private Application m_InputBOMExcelApp;
        private Workbook m_WorkBook;
        private object[,] m_ValueArray; //Original BOM Explosion data

        public object[,] ValueArray
        {
            get { return m_ValueArray; }
        }
        /// <summary>
        /// Initialize a new Excel reader. Must be integrated
        /// with an Excel interface object.
        /// </summary>
        public ExcelReaderInterop()
        {
            m_InputBOMExcelApp = new Application();
        }

        /// <summary>
        /// Open the file path received in Excel. Then, open the workbook
        /// within the file. Send the workbook to the next function, the internal scan
        /// function. Will throw an exception if a file cannot be found or opened.
        /// </summary>
        public void ExcelOpenSpreadsheets(string thisFileName)
        {
            try
            {
                m_WorkBook = m_InputBOMExcelApp.Workbooks.Open(thisFileName,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);

                ExcelScanIntenal();

                // Clean up.
                m_WorkBook.Close(false, thisFileName, null);
                Marshal.ReleaseComObject(m_WorkBook);
            }
            catch (Exception ex)
            {
                m_WorkBook.Close(false, thisFileName, null);
                Marshal.ReleaseComObject(m_WorkBook);
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void ExcelScanIntenal()
        {
            int sheetNum = 1; //should only be 1 sheet in a BOM Explosion
            Worksheet sheet = (Worksheet)m_WorkBook.Sheets[sheetNum];

            //
            // Take the used range of the sheet. Finally, get an object array of all
            // of the cells in the sheet (their values). Store it in m_ValueArray.
            //
            Range excelRange = sheet.UsedRange;
            m_ValueArray = (object[,])excelRange.get_Value(
                XlRangeValueDataType.xlRangeValueDefault);

        }
    
    }
}
