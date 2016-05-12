﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOMSplitter
{
    class BOMItem
    {
        private string m_Level;
        private string m_SubClass;
        private string m_PartNumber;
        private string m_RevEco;
        private string m_Description;
        private string m_UnitOfMeasure;
        private string m_Notes;
        private string m_OrigRefDes;
        private int m_Qty;
        private int m_OrigFindNum;
        private Dictionary<int, string> m_RefDes; // key = FindNum , value = refdes string



        public BOMItem(string level, string subclass, string partno, string rev, string desc, int qty, string unit, int findnum, string refdes, string notes)
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
            m_RefDes = new Dictionary<int, string>();
            m_RefDes.Add(m_OrigFindNum, m_OrigRefDes);
            m_Notes = notes;
        }

        public BOMItem(int FindNum)
        {
            m_OrigFindNum = FindNum;
        }

        public int QtySplitOne
        {
            get
            {
                string str = m_RefDes[m_OrigFindNum + 1];
                string[] rds = str.Split(new char[] { ',' });
                return rds.Length;
            }
        }
        public int QtySplitTwo
        {
            get
            {
                string str = m_RefDes[m_OrigFindNum + 2];
                string[] rds = str.Split(new char[] { ',' });
                return rds.Length;
            }
        }
        public int OldFindNum
        {
            get { return m_OrigFindNum; }
        }
        public int FirstNewFNum
        {
            get { return m_OrigFindNum + 1; }
        }
        public int SecondNewFNum
        {
            get { return m_OrigFindNum + 2; }
        }
        public string FirstSplitLine
        {
            get { return m_RefDes[m_OrigFindNum + 1]; }
        }
        public string SecondSplitLine
        {
            get { return m_RefDes[m_OrigFindNum + 2]; }
        }
        public Dictionary<int, string> RefDes
        {
            get { return m_RefDes; }
        }
        public string PartNumber
        {
            get { return m_PartNumber; }
        }
        public int Qty
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

        public bool SplitPart(string partnum, List<string> splits)
        {
            try
            {
                if (partnum == m_PartNumber)
                {
                    m_RefDes.Add(m_OrigFindNum + 1, splits[0]);
                    m_RefDes.Add(m_OrigFindNum + 2, splits[1]);
                    m_RefDes.Remove(m_OrigFindNum);
                }
                else
                {
                    return false;
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
            return true;
        }

    }
}
