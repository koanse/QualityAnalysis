using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace QualityAnalysis
{
    public partial class MainForm : Form
    {
        double[,] matrX, matrY;
        string[] arrXName, arrYName;
        double[] arrYMin, arrYMax, arrYAv, arrXAv;
        string[] arrRep;
        public MainForm()
        {
            InitializeComponent();
        }
        void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }        
        void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImportForm form = new ImportForm();
            if (form.ShowDialog() != DialogResult.OK)
                return;
            matrX = form.matrX;
            matrY = form.matrY;
            arrXName = form.arrXName;
            arrYName = form.arrYName;
            arrYMin = form.arrYMin;
            arrYMax = form.arrYMax;
            arrXAv = form.arrXAv;
            arrYAv = form.arrYAv;
            string[] arrG, arrGU;
            int[][] arrArrI;
            int N, N1;
            Quality.SetID(matrY, arrYMin, arrYMax, out arrG, out arrGU, out arrArrI, out N, out N1);
            arrRep = new string[19];
            arrRep[0] = Quality.DataTable(arrXName, arrYName, matrX, matrY, arrG);
            arrRep[8] = "";
            wb.DocumentText = arrRep[0];
        }
        void genToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                IdForm form = new IdForm(matrX, matrY, arrXAv, arrXName, arrYName);
                if (form.ShowDialog() != DialogResult.OK)
                    return;
                matrX = form.matrXM;
                matrY = form.matrYM;
                string[] arrG, arrGU;
                int[][] arrArrI;
                int N, N1;
                Quality.SetID(matrY, arrYMin, arrYMax, out arrG, out arrGU, out arrArrI, out N, out N1);
                arrRep[0] = Quality.DataTable(arrXName, arrYName, matrX, matrY, arrG);
                arrRep[1] = form.repSmp;
                arrRep[2] = form.repHSmp;
                arrRep[3] = form.repId;
                arrRep[4] = form.repHId;
                arrRep[5] = form.repCorrXY;
                arrRep[6] = form.repCorr;
                arrRep[7] = form.repHCorr;
                arrRep[8] += form.repGen;
                wb.DocumentText = arrRep[0];
            }
            catch { }
        }
        void lb_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                wb.DocumentText = arrRep[lb.SelectedIndex];
            }
            catch { }
        }
        void analizeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                string[] arrG, arrGU;
                int[][] arrArrI;
                int N, N1;
                Quality.SetID(matrY, arrYMin, arrYMax, out arrG, out arrGU, out arrArrI, out N, out N1);
                int[] arrIS, arrFreq;
                string[] arrGS;
                int[][] arrArrIS;
                Quality.Sort(arrGU, arrArrI, out arrIS, out arrFreq, out arrGS, out arrArrIS);
                string[] arrGC;
                double[] arrP, arrP1, arrPM;
                int[] arrGFreq;
                double H;
                Quality.Clustering(arrGS, arrArrIS, N, N1, out arrGC, out arrGFreq, out arrP, out arrP1, out arrPM, out H);
                arrRep[9] = string.Format("Общее количество экземпляров: {0}<br>" +
                    "Количество качественных экземпляров: {1}<br>" +
                    "Количество бракованных экземпляров: {2}<br>", N, N - N1, N1) +
                    Quality.PieChart("Соотношение качественных и бракованных экземпляров", "pc1",
                    new double[] { N - N1, N1 }, new string[] { "Качественные", "Брак" });
                arrRep[10] = Quality.GTable(arrGU, arrArrI);
                double[] arrF = new double[arrYName.Length];
                string[] arrStr = new string[arrYName.Length];
                for (int i = 0; i < arrYName.Length; i++)
                {
                    arrF[i] = arrFreq[i];
                    arrStr[i] = arrYName[arrIS[i]];
                }
                arrRep[11] = Quality.ITable(arrIS, arrYName, arrFreq) + "<br>" +
                    Quality.PieChart("Частота недостижения показателей качества",
                    "pc2", arrF, arrStr);
                arrRep[12] = Quality.GTable(arrGS, arrArrIS);
                arrF = new double[arrGS.Length];
                for (int i = 0; i < arrGS.Length; i++)
                    arrF[i] = arrArrIS[i].Length;
                arrRep[13] = Quality.PieChart("Диаграмма состояния качества продукции", "pc3", arrF, arrGS);
                arrRep[14] = Quality.PTable(arrGC, arrGFreq, arrP, arrP1, arrPM);
                arrRep[15] = Quality.ID2dTable(arrGC, arrP) + "<br>" +
                    Quality.PieChart("Распределение дефектов", "pc4", arrP, arrGC);
                arrRep[16] = Quality.ID2dTable(arrGC, arrP1);
                arrRep[17] = Quality.ID2dTable(arrGC, arrPM) + "<br>" +
                    Quality.PieChart("Распределение дефектов\n(при условии независимости показателей качества)",
                    "pc5", arrPM, arrGC);
                arrRep[18] = string.Format("H = {0:F3}", H);
            }
            catch
            {
                MessageBox.Show("Ошибка анализа");
            }
        }
    }
}