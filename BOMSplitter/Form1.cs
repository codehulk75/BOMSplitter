using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;

namespace BOMSplitter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string m_BOMExplosionFileName = null;
        private string m_SplitFileName = null;
        private System.Data.DataTable m_BOMData; //BOM Explosion converted to from valarray to datatable, work with this copy, valarray should stay constant
        private List<string> m_FoundPrevSplits = new List<string>(); //will contain a list of any part numbers found to be already split on original BOM, therefore not processed
        //m_Splits is read in from split file and holds all the Info needed to edit the orginal BOM
        private Dictionary<string, List<string>> m_Splits = new Dictionary<string, List<string>>(); //key = pn, value=2 strings containing top and bot ref des
        private List<BOMItem> m_BOMParts = new List<BOMItem>(); //just the lines in Parts category from the BOM, this will be edited with splits

        //Export stuff
        Microsoft.Office.Interop.Excel.Application m_ExportBOMExcelApp;
        Workbook m_ExpBook;

        private void GetBOMParts()
        {
            foreach (DataRow row in m_BOMData.Rows)
            {
                if (string.IsNullOrEmpty(row[8].ToString()))
                {
                    //No ref des's??  Don't bother...
                    continue;
                }
                if(row[1].ToString() == "Part")
                {
                    try
                    {
                        string lvl = row[0].ToString();
                        string pn = row[2].ToString();
                        string rev = row[3].ToString();
                        string desc = row[4].ToString();
                        int fn = Convert.ToInt32(row[5].ToString());
                        int qty = Convert.ToInt32(row[6].ToString());
                        string unit = row[7].ToString();
                        string rd = row[8].ToString();
                        string comments = row[9].ToString();
                        m_BOMParts.Add(new BOMItem(lvl, "Part", pn, rev, desc, qty, unit, fn, rd, comments));
                    }
                    catch(Exception ex)             
                    {
                        MessageBox.Show(ex.Message);
                        continue;
                    }
                }
            }
        }
        private void openFileButton_Click(object sender, EventArgs e)
        {
            //
            //TO DO: Make sure Data is cleared out every time before adding a new BOM???
            //

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls";

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                m_BOMExplosionFileName = openDialog.FileName;
                bomFileTextBox.Text = m_BOMExplosionFileName;
                ExcelReaderInterop rdr = new ExcelReaderInterop();
                rdr.ExcelOpenSpreadsheets(m_BOMExplosionFileName);
                m_BOMData = ArrayToDataTable(rdr.M_ValueArray); //populate the editable copy of the BOM, include all lines from BOM Explosion
                m_BOMData.TableName = "BOM";
                bomGridView.DataSource = m_BOMData;
                GetBOMParts(); //Populate list of BOMItems with only lines from BOM that are in the 'Part' category
            }
        }

        private System.Data.DataTable ArrayToDataTable(object[,] sheetdata)
        {
            //this needs to work perfectly to return every single cell in the BOM in the used range, including blank,
            //so the BOM can be restored to its original form, plus edits, in a reliable manner
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("Level");
            dt.Columns.Add("SubClass");
            dt.Columns.Add("BEInum");
            dt.Columns.Add("RevECO");
            dt.Columns.Add("Description");
            dt.Columns.Add("FindNum");
            dt.Columns.Add("Qty");
            dt.Columns.Add("UnitOfMeasure");
            dt.Columns.Add("RefDes");
            dt.Columns.Add("Notes");

            for (int i = 1; i < sheetdata.GetLength(0)+1; i++)
            {
                DataRow row = dt.NewRow();
                for (int j = 1; j < sheetdata.GetLength(1)+1; j++)
                {
                    row[j - 1] = sheetdata[i, j];
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        private void closeButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        public void ExportSplits()
        {
            try
            {              
                m_ExportBOMExcelApp = new Microsoft.Office.Interop.Excel.Application();
                m_ExpBook = m_ExportBOMExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet exportSheet = m_ExpBook.Worksheets[1];

                for (int col = 0; col < m_BOMData.Columns.Count; col++)
                {
                    for (int row = 0; row < m_BOMData.Rows.Count; row++)
                    {
                        exportSheet.Cells[row + 1, col + 1] = m_BOMData.Rows[row].ItemArray[col];
                    }
                }
                //microsoft.office.interop.excel.range firstcell = (microsoft.office.interop.excel.range)exportsheet.cells["a1"];
                //microsoft.office.interop.excel.range lastcell = (microsoft.office.interop.excel.range)exportsheet.cells[m_bomdata.rows.count, m_bomdata.columns.count];
                //microsoft.office.interop.excel.range targetrange = (microsoft.office.interop.excel.range)exportsheet.range[firstcell, lastcell];
                //targetrange.value = m_bomdata;

                // Clean up.
                m_ExpBook.Close(true);
                Marshal.ReleaseComObject(m_ExpBook);
            }
            catch (Exception ex)
            {
                m_ExpBook.Close(true);
                Marshal.ReleaseComObject(m_ExpBook);
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void splitFileButton_Click(object sender, EventArgs e)
        {
            //user chooses split file created by Methods Dept.
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Split Files (*.txt)|*.txt|All Files (*.*)|*.*";

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                m_SplitFileName = openDialog.FileName;
                splitFileTextBox.Text = m_SplitFileName;
            }
        }

        private void doSplitsButton_Click(object sender, EventArgs e)
        {
            //If the BOM file and the Split file aren't loaded, get out of here!!
            if (string.IsNullOrEmpty(m_BOMExplosionFileName) || string.IsNullOrEmpty(m_SplitFileName))
            {
                MessageBox.Show("Please Load the BOM and Split files first.", "Not Enough Data");
                return;
            }
            try
            {
                using (System.IO.StreamReader splitreader = new System.IO.StreamReader(m_SplitFileName))
                {
                    string line = null;
                    string pn = null;
                    Regex botreg = new Regex(@"BOTTOM");
                    Regex topreg = new Regex(@"TOP");
                    List<string> refs = new List<string>(); // temp container for ref des's, gets cleared every part number
                    while((line = splitreader.ReadLine()) != null)
                    {
                        if(!string.IsNullOrEmpty(line))
                        {
                            string[] lst = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);//(char[])null == split on whitespace
                            if (lst.Length == 4)
                            {
                                Match botMatch = botreg.Match(lst[1]);
                                if (botMatch.Success)
                                {
                                    pn = lst[0]; // get part number
                                    line = splitreader.ReadLine();
                                    string[] botrefs = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);//read all ref des's as 1 string
                                    refs.Add(botrefs[0]);
                                }
                            } else if(lst.Length == 3)
                            {
                                Match topMatch = topreg.Match(lst[0]);//If 'TOP' is the first non-whitespace word on the line, top ref des's follow next line
                                if(topMatch.Success)
                                {
                                    line = splitreader.ReadLine();
                                    string[] toprefs = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                                    refs.Add(toprefs[0]);
                                    m_Splits.Add(pn, new List<string>(refs));
                                    refs.Clear();
                                }
                            }
                        }                  
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            SplitBOM();
            ExportSplits();
        }
   
        private void SplitBOM()
        {
            foreach (var splitPNData in m_Splits)
            {               
                List<BOMItem> foundParts = m_BOMParts.FindAll(x => x.PartNumber == splitPNData.Key);
                if(foundParts.Count > 1)
                {
                    //if previous split is found, add it to a list for notification later, and do not process new split
                    m_FoundPrevSplits.Add(foundParts[0].PartNumber);
                    continue;
                }
                else if (foundParts.Count < 1)
                {
                    continue;
                }
                if (foundParts[0].SplitPart(splitPNData.Key, splitPNData.Value) == true)
                {
                    //make a new datatable or does datatable have decent find/replace option?
                    string expression = "BEInum='" + foundParts[0].PartNumber + "'";
                    DataRow[] foundRows = m_BOMData.Select(expression);
                    foreach (DataRow row in foundRows)
                    {
                        int index = m_BOMData.Rows.IndexOf(row);
                        m_BOMData.Rows[index]["FindNum"] = foundParts[0].FirstNewFNum;
                        m_BOMData.Rows[index]["RefDes"] = foundParts[0].FirstSplitLine;
                        m_BOMData.Rows[index]["Qty"] = foundParts[0].QtySplitOne;
                        DataRow splitLine2 = m_BOMData.NewRow();
                        splitLine2["Level"] = m_BOMData.Rows[index]["Level"];
                        splitLine2["SubClass"] = m_BOMData.Rows[index]["SubClass"];
                        splitLine2["BEInum"] = m_BOMData.Rows[index]["BEInum"];
                        splitLine2["RevECO"] = m_BOMData.Rows[index]["RevECO"];
                        splitLine2["Description"] = m_BOMData.Rows[index]["Description"];
                        splitLine2["FindNum"] = foundParts[0].SecondNewFNum;
                        splitLine2["Qty"] = foundParts[0].QtySplitTwo;
                        splitLine2["UnitOfMeasure"] = m_BOMData.Rows[index]["UnitOfMeasure"];
                        splitLine2["RefDes"] = foundParts[0].SecondSplitLine;
                        splitLine2["Notes"] = m_BOMData.Rows[index]["Notes"];
                        m_BOMData.Rows.InsertAt(splitLine2, index+1);
                    }
                }
            }
        }
    }
}
