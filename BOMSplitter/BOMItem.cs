using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOMSplitter
{
    public class BOMItem
    {
        private string m_Level;
        private string m_SubClass;
        private string m_PartNumber;
        private string m_RevEco;
        private string m_Description;
        private string m_UnitOfMeasure;
        private string m_Notes;
        private string m_OrigRefDes; //may get merged if it's found to be pre-split already, for untouched original use 'PreMergeOriginalRefDes'
        private string m_PreMergeOriginalRefDes;
        private string m_Qty;
        private int m_OrigFindNum;
        private Dictionary<int, string> m_RefDes; // key = FindNum , value = refdes string



        public BOMItem(string level, string subclass, string partno, string rev, string desc, string qty, string unit, int findnum, string refdes, string notes)
        {
            m_Level = level;
            m_SubClass = subclass;
            m_PartNumber = partno;
            m_RevEco = rev;
            m_Description = desc;
            m_Qty = qty;
            m_UnitOfMeasure = unit;
            m_OrigFindNum = findnum;
            m_OrigRefDes = refdes;
            m_PreMergeOriginalRefDes = refdes;
            m_RefDes = new Dictionary<int, string>();
            m_RefDes.Add(m_OrigFindNum, m_OrigRefDes);
            m_Notes = notes;
        }
        public int OldFindNum
        {
            get { return m_OrigFindNum; }
            set { m_OrigFindNum = value; }
        }
        public Dictionary<int, string> RefDes
        {
            get { return m_RefDes; }
        }
        public void AddRefDes(string newRefDeses, string newQty)
        {
            int qty = Convert.ToInt32(m_Qty) + Convert.ToInt32(newQty);
            m_Qty = qty.ToString();
            m_OrigRefDes += "," + newRefDeses;
            m_RefDes[m_OrigFindNum] = m_OrigRefDes;
        }
        public string PartNumber
        {
            get { return m_PartNumber; }
        }
        public string OriginalRefs
        {
            get { return m_PreMergeOriginalRefDes; }
        }
        public string Qty
        {
            get { return m_Qty; }
            set { m_Qty = value; }
        }

        public override string ToString()
        {
            string result = m_PartNumber;
            foreach(var entry in m_RefDes)
            {
                result += ("\n" + entry.Key.ToString() + "\t" + entry.Value + "\n");
            }
            return result;
        }
        
        private int ReGenFindNums(List<int> fNums)
        {
            fNums.Sort();
            int highest = fNums.Last();
            int nParentFN = m_OrigFindNum - m_OrigFindNum % 10;
            int newFNstart = 0;
            if (!fNums.Contains(nParentFN))
            {
                newFNstart = nParentFN + 1;
            }
            else
            {
                newFNstart = highest - highest % 10 + 11;
            }           
            List<int> oldkeys = new List<int>(RefDes.Keys);
            foreach(int key in oldkeys)
            {
                string tempRDs = RefDes[key];
                RefDes.Remove(key);
                RefDes.Add(newFNstart, tempRDs);
                ++newFNstart;
                ++highest;
            }
            return highest > newFNstart ? highest:newFNstart;
        }
        public int SplitPart(string partnum, List<string> splits, List<int> fNums)
        {
            int NewHighFindNum = -1;
            if (partnum == m_PartNumber)
            {
                try
                {
                    int index = 0;
                    int offset = 1;                 
                    foreach (string split in splits)
                    {                       
                        int newFindNum = m_OrigFindNum + offset;
                        if (newFindNum % 10 == 5)
                        {
                            //FindNums ending in 5 are reserved for Alt. PN's
                            ++newFindNum;
                            ++offset;
                        }
                        if (fNums.Contains(newFindNum))
                        {
                            NewHighFindNum = 0;
                        }
                        if(NewHighFindNum != 0)
                        {
                            NewHighFindNum = newFindNum;
                        }
                        m_RefDes.Add(newFindNum, splits[index]);
                        ++index;
                        ++offset;                                                             
                    }
                    m_RefDes.Remove(m_OrigFindNum);
                    if (NewHighFindNum == 0)
                    {
                        NewHighFindNum = ReGenFindNums(fNums);       
                    }                   
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Error in BOMItem->SplitPart(). PartNum = " + partnum + "\nError = " + ex.Message);
                    return 0;
                }
            }
            else
                return 0;
            return NewHighFindNum;
        }

    }
}
