using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QualityAnalysis
{
    public partial class ImportForm : Form
    {
        public double[,] matrX, matrY;
        public string[] arrXName, arrYName;
        public double[] arrYMin, arrYMax, arrYAv, arrXAv;
        public ImportForm()
        {
            InitializeComponent();
            tbF.Text = "Данные.xls";
            tbSx.Text = "Лист1";
            tbSy.Text = "Лист1";
            tbSmin.Text = "Лист2";
            tbSmax.Text = "Лист2";
            tbRx.Text = "A1:N635";
            tbRy.Text = "O1:Q635";
            tbRmin.Text = "B1:B4";
            tbRmax.Text = "C1:C4";
        }
        private void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                Quality.ImportData(tbF.Text, tbSx.Text, tbSy.Text, tbRx.Text, tbRy.Text,
                    tbSmin.Text, tbSmax.Text, tbRmin.Text, tbRmax.Text,
                    out matrX, out matrY, out arrXName, out arrYName, out arrYMin, out arrYMax);
                Quality.Norm(matrY, out arrYAv);
                Quality.Norm(matrX, out arrXAv);
                for (int i = 0; i < arrYMin.Length; i++)
                {
                    arrYMin[i] /= arrYAv[i];
                    arrYMax[i] /= arrYAv[i];
                }
                DialogResult = DialogResult.OK;
            }
            catch
            {
                MessageBox.Show("Ошибка импорта данных");
            }
        }
        void btnOpen_Click(object sender, EventArgs e)
        {
            try
            {
                if (openFileDialog1.ShowDialog() != DialogResult.OK)
                    return;
                tbF.Text = openFileDialog1.FileName;
            }
            catch
            {
                MessageBox.Show("Ошибка отрытия файла");
            }
        }
        void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}