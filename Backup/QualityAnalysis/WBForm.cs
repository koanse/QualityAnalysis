using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QualityAnalysis
{
    public partial class WBForm : Form
    {
        public WBForm(string title, string s)
        {
            InitializeComponent();
            Text = title;
            wb.DocumentText = s;
        }
    }
}