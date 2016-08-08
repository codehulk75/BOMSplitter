using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BOMSplitter
{
    public partial class MergedItemsMB : Form
    {
        public MergedItemsMB(List<BOMItem> dups)
        {
            InitializeComponent();
            string displayDups = null;
            foreach (BOMItem item in dups)
            {
                displayDups += item.PartNumber + "\n" + item.OldFindNum + " => " + item.OriginalRefs + "\n\n";
            }
            richTextBox1.Text = displayDups;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
