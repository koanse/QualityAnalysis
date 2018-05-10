using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using ZedGraph;

namespace QualityAnalysis
{
    public partial class IdForm : Form
    {
        double[,] matrX, matrY;
        double[] arrXAv;
        string[] arrXName, arrYName;
        int[] arrI;
        public double[,] matrXM, matrYM;
        public string repSmp, repHSmp, repId, repHId, repCorrXY, repCorr, repHCorr, repGen;
        public IdForm(double[,] matrX, double[,] matrY, double[] arrXAv, string[] arrXName, string[] arrYName)
        {
            InitializeComponent();
            this.matrX = matrX;
            this.matrY = matrY;
            this.arrXAv = arrXAv;
            this.arrXName = arrXName;
            this.arrYName = arrYName;
            object[] arrObj = new object[arrYName.Length + 1];
            for (int i = 0; i < arrYName.Length; i++)
            {
                dgv.Columns.Add(new DataGridViewComboBoxColumn());
                dgv.Columns[i + 1].HeaderText = arrYName[i];
                (dgv.Columns[i + 1] as DataGridViewComboBoxColumn).Items.AddRange(new string[] {
                    "x", "x^2", "x^3", "1/x", "1/x^2", "sqrt(x)", "1/sqrt(x)", "ln(x)", "exp(x)" });
                arrObj[i + 1] = "x";
            }
            for (int i = 0; i < arrXName.Length; i++)
            {
                arrObj[0] = arrXName[i];
                dgv.Rows.Add(arrObj);
            }
        }
        void btnOk_Click(object sender, EventArgs e)
        {
            try
            {
                Regression[] arrReg = new Regression[matrY.GetLength(1)];
                repSmp = repHSmp = repId = repHId = repCorr = repHCorr = repGen = "";
                Sample[] arrSX = new Sample[matrX.GetLength(1)];
                for (int j = 0; j < arrSX.Length; j++)
                {
                    double[] arrX = new double[matrX.GetLength(0)];
                    for (int i = 0; i < matrX.GetLength(0); i++)
                        arrX[i] = matrX[i, j];
                    arrSX[j] = new Sample(arrXName[j], string.Format("x<sub>{0}</sub>", j + 1), arrX);
                    arrSX[j].Calculate();
                    repSmp += arrSX[j].GetReport();
                    repHSmp += arrSX[j].GetHypReport();
                }
                repId = "ПАРАМЕТРИЧЕСКАЯ ИДЕНТИФИКАЦИЯ<br>";
                Sample[] arrSY = new Sample[matrY.GetLength(1)];
                for (int j = 0; j < matrY.GetLength(1); j++)
                {
                    Sample[] arrTSX = new Sample[matrX.GetLength(1)];
                    for (int k = 0; k < matrX.GetLength(1); k++)
                        arrTSX[k] = new TranSample(arrSX[k],
                            dgv.Rows[k].Cells[j + 1].Value.ToString());
                    double[] arrY = new double[matrY.GetLength(0)];
                    for (int i = 0; i < matrY.GetLength(0); i++)
                        arrY[i] = matrY[i, j];
                    Sample smpY = new Sample(arrYName[j], string.Format("y<sub>{0}</sub>", j + 1), arrY);
                    smpY.Calculate();
                    arrSY[j] = smpY;
                    repSmp += smpY.GetReport();
                    repHSmp += smpY.GetHypReport();
                    arrReg[j] = new Regression(smpY, arrTSX);
                    repId += arrReg[j].GetRegReport() + "<br>";
                    repCorr += arrReg[j].GetCorrReport();
                    if (cbH.Checked)
                    {
                        arrReg[j].CheckHypothesises(double.Parse(tbAlpha.Text));
                        repHId += arrReg[j].GetHypRegrReport();
                        repHCorr += arrReg[j].GetHypCorrReport();
                    }
                }
                Sample[] arrSYX = new Sample[matrX.GetLength(1) + matrY.GetLength(1) - 1];
                for (int i = 0; i < arrSY.Length - 1; i++)
                    arrSYX[i] = arrSY[i + 1];
                for (int i = 0; i < arrSX.Length; i++)
                    arrSYX[arrSY.Length - 1 + i] = arrSX[i];
                Regression reg = new Regression(arrSY[0], arrSYX);
                repCorrXY = reg.GetCorrReport().Replace("<P>ПОКАЗАТЕЛЬ: y<sub>1</sub><BR>", "");
                if (cbG.Checked)
                {
                    Quality.GenErr(matrX, matrY, arrReg, int.Parse(tbN.Text), double.Parse(tbA.Text),
                        double.Parse(tbB.Text), out arrI, out matrXM, out matrYM);
                    repGen += string.Format("ПАРАМЕТРЫ ГЕНЕРАЦИИ<br>n = {0}<br>A = {1}<br>B = {2}<br><br>",
                        tbN.Text, tbA.Text, tbB.Text);
                    repGen += "ПЕРЕЧЕНЬ ИЗМЕНЕННЫХ НАБЛЮДЕНИЙ<br>";
                    foreach (int i in arrI)
                        repGen += string.Format("{0} ", i + 1);
                }
                else
                {
                    matrXM = matrX;
                    matrYM = matrY;
                }
                DialogResult = DialogResult.OK;
                Close();
            }
            catch
            {
                MessageBox.Show("Ошибка генерации");
            }
        }
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}