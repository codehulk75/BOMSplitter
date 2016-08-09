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
        private System.Data.DataTable m_BOMData; //BOM Explosion converted to from valarray to datatable, this is for GUI display of the pre/post split BOM
        private System.Data.DataTable m_OutputBOM; //Tracks m_BOMData, but without original 'pre-split' lines, this is final user output
        private List<string> m_FoundPrevSplits = new List<string>(); //will contain a list of any part numbers found to be already split on original BOM, therefore not processed
        //m_Splits is read in from split file and holds all the Info needed to edit the orginal BOM
        private Dictionary<string, List<string>> m_Splits = new Dictionary<string, List<string>>(); //key = pn, value=2 strings containing top and bot ref des
        private List<BOMItem> m_BOMParts = new List<BOMItem>(); //just the lines in Parts category from the BOM, this will be edited with splits
        private Workbook m_ExpBook;
        private string m_AssemblyNumber = null;
        private int mergeFlag = -1;
        private void ClearAllData()
        {
            //Every time user choose a new BOM file, call this routine to reset all the data  
            m_AssemblyNumber = null;         
            m_SplitFileName = null;
            splitFileTextBox.Clear();
            if(m_BOMData != null)
                m_BOMData.Clear();
            if(m_OutputBOM != null)
                m_OutputBOM.Clear();
            if (m_FoundPrevSplits.Count > 1)
                m_FoundPrevSplits.Clear();
            if(m_Splits.Count > 1)
                m_Splits.Clear();
            if(m_BOMParts.Count > 1)
                m_BOMParts.Clear();
            m_ExpBook = null;
            mergeFlag = -1;
        }
    
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
                        string qty = row[6].ToString();
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
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel Files (*.xlsx, *.xls)|*.xlsx;*.xls";

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                //Reset all variables every time user picks a new BOM file.
                ClearAllData();

                m_BOMExplosionFileName = openDialog.FileName;
                bomFileTextBox.Text = m_BOMExplosionFileName;
                ExcelReaderInterop rdr = new ExcelReaderInterop();
                rdr.ExcelOpenSpreadsheets(m_BOMExplosionFileName);
                m_BOMData = ArrayToDataTable(rdr.ValueArray); //populate the editable copy of the BOM, include all lines from BOM Explosion
                m_BOMData.TableName = "BOM";
                bomGridView.DataSource = m_BOMData;

                //set output bom to same data as m_BOMData
                m_OutputBOM = ArrayToDataTable(rdr.ValueArray);
                m_OutputBOM.TableName = "BOM";
                SplitAndRejoinBOMNotes(m_OutputBOM);
                m_AssemblyNumber = m_OutputBOM.Rows[1][1].ToString();
                GetBOMParts(); //Populate list of BOMItems with only lines from BOM that are in the 'Part' category
            }
        }

        private void SplitAndRejoinBOMNotes(System.Data.DataTable dt)
        {
            //
            //There was an issue importing BOM Notes into Agile from the exported spreadsheet.
            //It was due to formatting that is not visible in Excel but nevertheless was there in the Notes field where every space is between words.
            //This split and rejoin routine fixes that issue.  Splits up BOM Notes on whitespace, and rejoins them with spaces (sans invisible formatting.)
            //
            int notesCol = dt.Columns.Count-1;
            foreach( DataRow dr in dt.Rows)
            {
                try
                {
                    string notestr = dr[notesCol].ToString();
                    if (notestr.Length > 1)
                    {
                        string[] temparr = notestr.Split();
                        notestr = string.Join(" ", temparr);
                        dr[notesCol] = notestr;
                    }
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    continue;
                }
            }
        }

        private System.Data.DataTable ArrayToDataTable(object[,] sheetdata)
        {
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

        private object[,] DataTableToArray(System.Data.DataTable dt)
        {
            int dataStart = 7;
            int dataEnd = dt.Rows.Count - 5;
            int headAndFoot = 12;
            int importColCount = 7;
            int rowDelta = 6;
            object[,] arr = new object[dt.Rows.Count - headAndFoot, importColCount];
            arr[0, 0] = "PARENT NO";
            arr[0, 1] = "CHILD NO";
            arr[0, 2] = "FIND NO";
            arr[0, 3] = "QTY";
            arr[0, 4] = "REF DES";
            arr[0, 5] = "NOTES";
            arr[0, 6] = "Description";
            try
            {
                for (int i = dataStart; i < dataEnd; i++)
                {
                    DataRow dr = dt.Rows[i];
                    string firstField = dt.Rows[i].ItemArray[0].ToString();
                    if (firstField.Equals("0")) //if this line is the parent number line(Level == 0), skip it             
                    {
                        ++rowDelta;
                        continue;
                    }
                    if (string.IsNullOrEmpty(firstField))
                        break;
                    arr[i - rowDelta, 0] = m_AssemblyNumber;
                    arr[i - rowDelta, 1] = dr[2];
                    arr[i - rowDelta, 2] = dr[5];
                    arr[i - rowDelta, 3] = dr[6];
                    arr[i - rowDelta, 4] = dr[8];
                    arr[i - rowDelta, 5] = dr[9];
                    arr[i - rowDelta, 6] = dr[4];
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message + "\nDataTableToArray()");
            }
            return arr;
        }
        private void closeButton_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        public void ExportSplits()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                m_ExpBook = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                Worksheet exportSheet = m_ExpBook.Worksheets[1];
                exportSheet.Cells.NumberFormat = "@";
                
                //copy to new Excel workbook and prompt user to save
                object[,] arr = DataTableToArray(m_OutputBOM);
                Range firstcell = exportSheet.Cells[1,1];
                Range lastcell = exportSheet.Cells[arr.GetLength(0), arr.GetLength(1)];
                Range targetrange = exportSheet.Range[firstcell, lastcell];
                targetrange.Value = arr;

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
                    Regex postreg = new Regex(@"POST");
                    List<string> refs = new List<string>(); // temp container for ref des's, gets cleared every part number
                    while((line = splitreader.ReadLine()) != null)
                    {
                        if(string.IsNullOrEmpty(line) == false)
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
                                    do
                                    {
                                        if((line = splitreader.ReadLine()) == null)
                                            break;
                                    } while (line == string.Empty);
                                    if (line != null)
                                    {
                                        string[] postattempt = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                                        if (postattempt.Length > 0)
                                        {
                                            Match postmatch = postreg.Match(postattempt[0]);
                                            if (postattempt.Length == 3 && postmatch.Success)
                                            {
                                                line = splitreader.ReadLine();
                                                string[] postrefs = line.Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                                                refs.Add(postrefs[0]);
                                            }
                                        }
                                    }
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
                MessageBox.Show("DoSplitsButtonCLick() => "+ex.Message);
            }
            SplitBOM();
            ExportSplits();
        }
   
        private void SplitBOM()
        {
            List<BOMItem> mergedParts = new List<BOMItem>();
            foreach (var splitPNData in m_Splits)
            {               
                List<BOMItem> foundParts = m_BOMParts.FindAll(x => x.PartNumber == splitPNData.Key);
                if(foundParts.Count > 1)
                {
                    if(MergeFlag() == true)
                    {
                        m_FoundPrevSplits.Add(foundParts[0].PartNumber);
                        string pts = string.Join("\n", foundParts);
                        mergedParts.AddRange(foundParts);
                        BOMItem mergedItem = MergeItems(foundParts);
                        if (mergedItem.SplitPart(splitPNData.Key, splitPNData.Value) == true)
                        {
                            UpdateGUIBOM(mergedItem);
                            UpdateOutputBOM(mergedItem);
                        }                      
                    }
                    continue;
                }
                else if (foundParts.Count < 1)
                {
                    MessageBox.Show("Split Part #" + splitPNData.Key + " not found in the BOM Explosion Report!!", "Something Isn't Right?!",MessageBoxButtons.OK, MessageBoxIcon.Warning);             
                    continue;
                }
                if (foundParts[0].SplitPart(splitPNData.Key, splitPNData.Value) == true)
                {
                    UpdateGUIBOM(foundParts[0]);
                    UpdateOutputBOM(foundParts[0]);                                
                }
            }
            if(mergedParts.Count == 0)
                return;
            MergedItemsMB mergeMessage = new MergedItemsMB(mergedParts);
            mergeMessage.ShowDialog();
        }

        private bool MergeFlag()
        {
            bool bmerge = true;
            switch(mergeFlag)
            {
                case -1:                    
                    var result = MessageBox.Show("This split file list parts that are already split.\n"
                        + "Click 'Yes' to merge the split and re-split with file data.\n"
                        + "Click 'No' to leave parts as is and ignore the new ones in the split file.", "Merge BOM?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        mergeFlag = 1;
                        bmerge = true;
                    }                         
                    else if (result == DialogResult.No)
                    {
                        mergeFlag = 0;
                        bmerge = false;
                    }                          
                    break;
                case 1:
                    bmerge =  true;
                    break;
                case 0:
                    bmerge = false;
                    break;
                default:
                    bmerge = true;
                    break;                 
            }
            return bmerge;
        }

        private BOMItem MergeItems(List<BOMItem> dupItems)
        {

            try
            {
                dupItems.Sort(delegate (BOMItem x, BOMItem y)
                {
                    if (x.OldFindNum < y.OldFindNum) return -1;
                    else return 1;
                });
                BOMItem firstBOMItem = PopAt(dupItems, 0);
                m_BOMParts.RemoveAll(x => x.OldFindNum.Equals(firstBOMItem.OldFindNum));
                foreach (BOMItem item in dupItems)
                {
                    firstBOMItem.AddRefDes(item.RefDes[item.OldFindNum], item.Qty);
                    m_BOMParts.RemoveAll(x => x.OldFindNum.Equals(item.OldFindNum));

                    string expression = "FindNum='" + item.OldFindNum + "'";
                    DataRow[] foundRowsGUI = m_BOMData.Select(expression);
                    foreach (DataRow row in foundRowsGUI)
                    {
                        m_BOMData.Rows.Remove(row);
                    }
                    DataRow[] foundRowsData = m_OutputBOM.Select(expression);
                    foreach (DataRow row in foundRowsData)
                    {
                        m_OutputBOM.Rows.Remove(row);

                    }
                }
                DataRow[] originalRow = m_BOMData.Select("FindNum='" + firstBOMItem.OldFindNum + "'");
                originalRow[0]["Qty"] = firstBOMItem.Qty;
                originalRow[0]["RefDes"] = firstBOMItem.RefDes[firstBOMItem.OldFindNum];
                m_BOMParts.Add(firstBOMItem);
                return firstBOMItem;
            }
            catch(Exception ex)
            {
                MessageBox.Show("Something went wrong when trying to merge the BOM :(\n"+ex.Message);
                return dupItems[0];
            }                     
        }

        public static T PopAt<T>(List<T> list, int index)
        {
            //helper method to remove and return an element froma list
            //I'm using it here to have like a 'Pop()' but from the beginning of the list (my obj = PopAt(myList, 0))
            T r = list[index];
            list.RemoveAt(index);
            return r;
        }

        private void UpdateGUIBOM(BOMItem bomItem)
        {
            string expression = "BEInum='" + bomItem.PartNumber + "'";
            DataRow[] foundRows = m_BOMData.Select(expression);
            foreach (DataRow row in foundRows)
            {
                int splitLineNum = 1;
                int index = m_BOMData.Rows.IndexOf(row);
                bomGridView.Rows[index].DefaultCellStyle.BackColor = Color.Red;
                foreach (KeyValuePair<int, string> entry in bomItem.RefDes)
                {
                    DataRow splitLine = m_BOMData.NewRow();
                    splitLine["Level"] = m_BOMData.Rows[index]["Level"];
                    splitLine["SubClass"] = m_BOMData.Rows[index]["SubClass"];
                    splitLine["BEInum"] = m_BOMData.Rows[index]["BEInum"];
                    splitLine["RevECO"] = m_BOMData.Rows[index]["RevECO"];
                    splitLine["Description"] = m_BOMData.Rows[index]["Description"];
                    splitLine["FindNum"] = entry.Key;
                    splitLine["Qty"] = entry.Value.Split(new char[] { ',' }).Length;
                    splitLine["UnitOfMeasure"] = m_BOMData.Rows[index]["UnitOfMeasure"];
                    splitLine["RefDes"] = entry.Value;
                    splitLine["Notes"] = m_BOMData.Rows[index]["Notes"];
                    m_BOMData.Rows.InsertAt(splitLine, index + splitLineNum);
                    ++splitLineNum;
                }                 
            }
        }
        private void UpdateOutputBOM(BOMItem bomItem)
        {
            string expression = "BEInum='" + bomItem.PartNumber + "'";
            DataRow[] foundRows = m_OutputBOM.Select(expression);            
            foreach (DataRow row in foundRows)
            {
                int index = m_OutputBOM.Rows.IndexOf(row);
                string level = m_OutputBOM.Rows[index]["Level"].ToString();
                string subclass = m_OutputBOM.Rows[index]["SubClass"].ToString();
                string beinum = m_OutputBOM.Rows[index]["BEInum"].ToString();
                string reveco = m_OutputBOM.Rows[index]["RevECO"].ToString();
                string desc = m_OutputBOM.Rows[index]["RevECO"].ToString();
                string unit = m_OutputBOM.Rows[index]["RevECO"].ToString();
                string notes = m_OutputBOM.Rows[index]["Notes"].ToString();
                m_OutputBOM.Rows.RemoveAt(index);
                foreach (KeyValuePair<int, string> entry in bomItem.RefDes)
                {
                    DataRow splitLine = m_OutputBOM.NewRow();
                    splitLine["Level"] = level;
                    splitLine["SubClass"] = subclass;
                    splitLine["BEInum"] = beinum;
                    splitLine["RevECO"] = reveco;
                    splitLine["Description"] = desc;
                    splitLine["FindNum"] = entry.Key;
                    splitLine["Qty"] = entry.Value.Split(new char[] {','}).Length;
                    splitLine["UnitOfMeasure"] = unit;
                    splitLine["RefDes"] = entry.Value;
                    splitLine["Notes"] = notes;
                    m_OutputBOM.Rows.InsertAt(splitLine, index);
                    ++index;
                }
            }
        }
    }
}
