using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using MathNet.Numerics.LinearAlgebra;
using MathNet.Numerics.Distributions;
using System.Collections;
using System.Drawing;
using System.Reflection;
using System.IO;
using ZedGraph;

namespace QualityAnalysis
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            /*double[,] matrX, matrY;
            string[] arrXName, arrYName;
            double[] arrYMin, arrYMax, arrYAv, arrXAv;
            Quality.ImportData("Данные.xls", "Лист1", "Лист1", "A1:N635", "O1:Q635",
                "Лист2", "Лист2", "B1:B4", "C1:C4",
                out matrX, out matrY, out arrXName, out arrYName, out arrYMin, out arrYMax);
            Quality.Norm(matrY, out arrYAv);
            Quality.Norm(matrX, out arrXAv);
            for (int i = 0; i < arrYMin.Length; i++)
            {
                arrYMin[i] /= arrYAv[i];
                arrYMax[i] /= arrYAv[i];
            }
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
            string s = Quality.DataTable(arrXName, arrYName, matrX, matrY, arrG);
            s += Quality.GTable(arrGU, arrArrI);
            s += Quality.ITable(arrIS, arrYName, arrFreq);
            s += Quality.GTable(arrGS, arrArrIS);
            s += Quality.PTable(arrGC, arrGFreq, arrP, arrP1, arrPM);
            s += Quality.ID2dTable(arrGC, arrP);*/
            Application.Run(new MainForm());            
        }
    }
    static class Quality
    {
        public static void ImportData(string file, string sX, string sY, string rX, string rY,
            string sYMin, string sYMax, string rYMin, string rYMax,
            out double[,] matrX, out double[,] matrY, out string[] arrXName, out string[] arrYName,
            out double[] arrYMin, out double[] arrYMax)
        {
            string s = "provider = Microsoft.Jet.OLEDB.4.0;" +
                "data source = " + file + ";" +
                "extended properties = Excel 8.0;";
            OleDbConnection conn = new OleDbConnection(s);
            conn.Open();
            s = string.Format("SELECT * FROM [{0}${1}]", sX.Replace('.', '#'), rX);
            OleDbDataAdapter da = new OleDbDataAdapter(s, conn);
            DataTable tX = new DataTable();
            da.Fill(tX);
            s = string.Format("SELECT * FROM [{0}${1}]", sY.Replace('.', '#'), rY);
            da = new OleDbDataAdapter(s, conn);
            DataTable tY = new DataTable();
            da.Fill(tY);
            conn.Close();
            s = string.Format("SELECT * FROM [{0}${1}]", sYMin.Replace('.', '#'), rYMin);
            da = new OleDbDataAdapter(s, conn);
            DataTable tYMin = new DataTable();
            da.Fill(tYMin);
            s = string.Format("SELECT * FROM [{0}${1}]", sYMax.Replace('.', '#'), rYMax);
            da = new OleDbDataAdapter(s, conn);
            DataTable tYMax = new DataTable();
            da.Fill(tYMax);
            if (tY.Rows.Count != tX.Rows.Count || tYMin.Rows.Count != tYMax.Rows.Count ||
                tY.Columns.Count != tYMin.Rows.Count)
                throw new Exception();


            matrX = new double[tX.Rows.Count, tX.Columns.Count];
            arrXName = new string[tX.Columns.Count];
            for (int i = 0; i < tX.Columns.Count; i++)
                arrXName[i] = tX.Columns[i].Caption;
            for (int i = 0; i < tX.Rows.Count; i++)
                for (int j = 0; j < tX.Columns.Count; j++)
                    matrX[i, j] = (double)tX.Rows[i][j];
            matrY = new double[tY.Rows.Count, tY.Columns.Count];
            arrYName = new string[tY.Columns.Count];
            for (int i = 0; i < tY.Columns.Count; i++)
                arrYName[i] = tY.Columns[i].Caption;
            for (int i = 0; i < tY.Rows.Count; i++)
                for (int j = 0; j < tY.Columns.Count; j++)
                    matrY[i, j] = (double)tY.Rows[i][j];
            arrYMin = new double[tYMin.Rows.Count];
            for (int i = 0; i < tYMin.Rows.Count; i++)
                arrYMin[i] = (double)tYMin.Rows[i][0];
            arrYMax = new double[tYMax.Rows.Count];
            for (int i = 0; i < tYMax.Rows.Count; i++)
                arrYMax[i] = (double)tYMax.Rows[i][0];
        }
        public static void Norm(double[,] matr, out double[] arrAv)
        {
            arrAv = new double[matr.GetLength(1)];
            for (int j = 0; j < matr.GetLength(1); j++)
            {
                arrAv[j] = 0;
                for (int i = 0; i < matr.GetLength(0); i++)
                    arrAv[j] += matr[i, j];
                arrAv[j] /= matr.GetLength(0);
                for (int i = 0; i < matr.GetLength(0); i++)
                    matr[i, j] /= arrAv[j];
            }
        }
        public static void SetID(double[,] matrY, double[] arrYMin, double[] arrYMax,
            out string[] arrG, out string[] arrGU, out int[][] arrArrI, out int N, out int N1)
        {
            arrG = new string[matrY.GetLength(0)];
            List<string> listGU = new List<string>();
            List<List<int>> listListI = new List<List<int>>();
            for (int i = 0; i < matrY.GetLength(0); i++)
            {
                string g = "";
                for (int j = 0; j < matrY.GetLength(1); j++)
                    if (matrY[i, j] <= arrYMax[j] &&
                        matrY[i, j] >= arrYMin[j])
                        g += "0";
                    else
                        g += "1";
                arrG[i] = g;
                int index = listGU.BinarySearch(g);
                if (index < 0)
                {
                    listGU.Insert(~index, g);
                    List<int> l = new List<int>();
                    l.Add(i);
                    listListI.Insert(~index, l);
                }
                else
                    listListI[index].Add(i);
            }
            arrArrI = new int[listGU.Count][];
            for (int i = 0; i < arrArrI.Length; i++)
                arrArrI[i] = listListI[i].ToArray();
            arrGU = listGU.ToArray();
            N = arrG.Length;
            int Ngood = 0;
            if (!arrGU[0].Contains("1"))
                Ngood = arrArrI[0].Length;
            N1 = N - Ngood;
        }
        public static void Sort(string[] arrGU, int[][] arrArrI,
            out int[] arrIS, out int[] arrFreq, out string[] arrGS, out int[][] arrArrIS)
        {
            int n = arrGU[0].Length;
            arrIS = new int[n];
            arrFreq = new int[n];
            arrGS = new string[arrGU.Length];
            for (int j = 0; j < n; j++)
                for (int i = 0; i < arrGU.Length; i++)
                    if (arrGU[i][j] == '1')
                        arrFreq[j] += arrArrI[i].Length;
            List<XY> list = new List<XY>();
            for (int i = 0; i < n; i++)
                list.Add(new XY(arrFreq[i], i));
            list.Sort();
            list.Reverse();
            for (int i = 0; i < list.Count; i++)
            {
                arrFreq[i] = list[i].x;
                arrIS[i] = list[i].y;
            }
            for (int i = 0; i < arrGS.Length; i++)
            {
                arrGS[i] = "";
                for (int j = 0; j < n; j++)
                    arrGS[i] += arrGU[i][arrIS[j]];
            }
            List<XY2> list2 = new List<XY2>();
            for (int i = 0; i < arrGS.Length; i++)
                list2.Add(new XY2(arrGS[i], arrArrI[i]));
            list2.Sort();
            arrArrIS = new int[arrArrI.Length][];
            for (int i = 0; i < list2.Count; i++)
			{
                arrGS[i] = list2[i].x;
                arrArrIS[i] = (int[])list2[i].y.Clone();
			}
        }
        public static void Clustering(string[] arrGS, int[][] arrArrI, int N, int N1,
            out string[] arrGC, out int[] arrGFreq, out double[] arrP, out double[] arrP1,
            out double[] arrPM, out double H)
        {
            int n = arrGS[0].Length;
            List<string> listGC = new List<string>();
            string g0 = "";
            for (int j = 0; j < n; j++)
                g0 += "0";
            for (int i = 0; i < n; i++)
            {
                for (int j = 0; j < n - i; j++)
                {
                    string g = g0;
                    g = g.Remove(i + j, 1).Insert(i + j, "1");
                    listGC.Add(g);
                }
                g0 = g0.Remove(i, 1).Insert(i, "1");
            }
            arrGC = listGC.ToArray();
            arrGFreq = new int[arrGC.Length];
            arrP = new double[arrGC.Length];
            arrP1 = new double[arrGC.Length];
            arrPM = new double[arrGC.Length];
            List<string> listGS = new List<string>(arrGS);
            double p0 = (N - N1) / (double)N;
            if (p0 > 0)
                H = -p0 * Math.Log(p0);
            else
                H = 0;
            for (int i = 0; i < arrGC.Length; i++)
            {
                int index = listGS.BinarySearch(arrGC[i]);
                if (index >= 0)
                {
                    arrGFreq[i] = arrArrI[index].Length;
                    arrP[i] = arrArrI[index].Length / (double)N;
                    arrP1[i] = arrArrI[index].Length / (double)N1;
                }
                arrPM[i] = 1;
                for (int j = 0; j < n; j++)
                    if (arrGC[i][j] == '1')
                        arrPM[i] *= arrP[j];
                H -= arrP[i] * Math.Log(arrP[i]);
            }
        }
        public static void GenErr(double[,] matrX, double[,] matrY, Regression[] arrReg,
            int n, double a, double b, out int[] arrI,
            out double[,] matrXM, out double[,] matrYM)
        {
            if (n > matrX.GetLength(0))
                throw new Exception();
            double[] arrXS = new double[matrX.GetLength(1)];
            double[] arrXAv = new double[matrX.GetLength(1)];
            for (int j = 0; j < matrX.GetLength(1); j++)
            {
                double sumX = 0, sumX2 = 0;
                for (int i = 0; i < matrX.GetLength(0); i++)
                {
                    sumX += matrX[i, j];
                    sumX2 += matrX[i, j] * matrX[i, j];
                }
                arrXAv[j] = sumX / matrX.GetLength(0);
                sumX2 /= matrX.GetLength(0);
                arrXS[j] = (sumX2 - arrXAv[j] * arrXAv[j]) * matrX.GetLength(0) /
                    (matrX.GetLength(0) - 1);
            }
            List<int> listAll = new List<int>();
            List<int> listI = new List<int>();
            for (int i = 0; i < matrX.GetLength(0); i++)
                listAll.Add(i);
            ContinuousUniformDistribution distUn = new ContinuousUniformDistribution(0, 1);
            for (int i = 0; i < n; i++)
            {                
                int li = (int)(distUn.NextDouble() * listAll.Count);
                listI.Add(listAll[li]);
                listAll.RemoveAt(li);
            }
            arrI = listI.ToArray();
            NormalDistribution dist = new NormalDistribution();
            matrXM = (double[,])matrX.Clone();
            for (int j = 0; j < matrX.GetLength(1); j++)
            {
                dist.SetDistributionParameters(a * arrXAv[j], b * arrXS[j]);
                foreach (int i in arrI)
                    matrXM[i, j] += dist.NextDouble();
            }
            matrYM = (double[,])matrY.Clone();
            double[] arrX = new double[matrX.GetLength(1)];
            foreach (int i in arrI)
            {
                for (int j = 0; j < matrX.GetLength(1); j++)
                    arrX[j] = matrXM[i, j];
                for (int j = 0; j < arrReg.Length; j++)
                    matrYM[i, j] = arrReg[j].GetRegValue(arrX);
            }
        }
        public static string DataTable(string[] arrXName, string[] arrYName,
            double[,] matrX, double[,] matrY, string[] arrG)
        {
            string res = "<TABLE border = 1 cellspacing = 0 align = center><TR><TD>i";
            for (int j = 0; j < arrXName.Length; j++)
                res += string.Format("<TD>{0}", arrXName[j]);
            for (int j = 0; j < arrYName.Length; j++)
                res += string.Format("<TD>{0}", arrYName[j]);
            res += "<TD>g";
            for (int i = 0; i < matrX.GetLength(0); i++)
            {
                res += string.Format("<TR><TD>{0}", i + 1);
                for (int j = 0; j < matrX.GetLength(1); j++)
                    res += string.Format("<TD>{0:F3}", matrX[i, j]);
                for (int j = 0; j < matrY.GetLength(1); j++)
                    res += string.Format("<TD>{0:F3}", matrY[i, j]);
                res += "<TD>" + arrG[i];
            }
            return res + "</TABLE>";
        }
        public static string GTable(string[] arrG, int[][]arrArrI)
        {
            string res = "<TABLE border = 1 cellspacing = 0 align = center>";
            res += "<TR><TD>g<TD>n(g)<TD>Номера экземпляров";
            for (int i = 0; i < arrG.Length; i++)
            {
                res += string.Format("<TR><TD>{0}<TD>{1}<TD>", arrG[i], arrArrI[i].Length);
                for (int j = 0; j < arrArrI[i].Length; j++)
                    res += string.Format("{0} ", arrArrI[i][j] + 1);
            }
            return res + "</TABLE>";
        }
        public static string PTable(string[] arrGC, int[] arrGFreq, double[] arrP,
            double[] arrP1, double[] arrPM)
        {
            string res = "<TABLE border = 1 cellspacing = 0 align = center>";
            res += "<TR><TD>g<TD>n(g)<TD>P(g)<TD>P<SUB>1</SUB>(g)<TD>P'(g)";
            for (int i = 0; i < arrGC.Length; i++)
                res += string.Format("<TR><TD>{0}<TD>{1}<TD>{2:F3}<TD>{3:F3}<TD>{4:F3}",
                    arrGC[i], arrGFreq[i], arrP[i], arrP1[i], arrPM[i]);
            return res + "</TABLE>";
        }
        public static string ITable(int[] arrIS, string[] arrYName, int[] arrFreq)
        {
            string res = "<TABLE border = 1 cellspacing = 0 align = center>";
            res += "<TR><TD>i<TD>Наименование показателя<TD>n<SUB>i</SUB>";
            for (int i = 0; i < arrIS.Length; i++)
                res += string.Format("<TR><TD>{0}<TD>{1}<TD>{2}",
                    i + 1, arrYName[arrIS[i]], arrFreq[i]);
            return res + "</TABLE>";
        }
        public static string ID2dTable(string[] arrGC, double[] arrP)
        {
            int n = arrGC[0].Length, k = 0;
            double ss = 0;
            double[] arrSum = new double[n];
            string res = "<TABLE border = 1 cellspacing = 0 align = center><TR><TD>-";
            for (int j = 0; j < n; j++)
                res += string.Format("<TD>g<SUB>{0}</SUB> = 1", j + 1);
            res += "<TD>Итого";
            for (int i = 0; i < n; i++)
            {
                res += "<TR><TD>";
                string s = "";
                if (i == 0)
                    res += "-";
                for (int j = 0; j < i; j++)
                {
                    res += string.Format("g<SUB>{0}</SUB> = 1", j + 1);
                    if (j < i - 1)
                        res += " ∩ ";
                    s += "<TD>-";
                }
                res += s;
                double sum = 0;
                for (int j = 0; j < n - i; j++)
                {
                    sum += arrP[k];
                    arrSum[i + j] += arrP[i];
                    res += string.Format("<TD>{0:F3}", arrP[k++]);
                }
                ss += sum;
                res += string.Format("<TD>{0:F3}", sum);
            }
            res += "<TR><TD>Итого";
            for (int j = 0; j < n; j++)
                res += string.Format("<TD>{0:F3}", arrSum[j]);
            res += string.Format("<TD>Сумма по вертикали: {0:F3}", ss);
            return res + "</TABLE>";
        }
        public static string PieChart(string name, string filename, double[] arrX, string[] arrLab)
        {
            ZedGraphControl zgc = new ZedGraphControl();
            zgc.Size = new Size(500, 500);
            zgc.GraphPane.Title.Text = name;
            zgc.GraphPane.AddPieSlices(arrX, arrLab);
            zgc.AxisChange();
            zgc.GetImage().Save(filename + ".gif", System.Drawing.Imaging.ImageFormat.Gif);
            string exePath = Assembly.GetExecutingAssembly().Location;
            string exeDir = Path.GetDirectoryName(exePath);
            return string.Format("<img src = '{0}/{1}.gif'>", exeDir, filename);
        }
    }
    static class Html
    {
        public static string MakeTable(double[,] matr,
            string[] arrHor, string[] arrVer, int d)
        {
            string res = "<TABLE border = 1><TR><TD>-";
            string f = "<TD>{0:F" + d.ToString() + "}";
            for (int j = 0; j < matr.GetLength(1); j++)
                res += "<TD>" + arrHor[j];
            for (int i = 0; i < matr.GetLength(0); i++)
            {
                res += "<TR><TD>" + arrVer[i];
                for (int j = 0; j < matr.GetLength(1); j++)
                    res += string.Format(f, matr[i, j]);
            }
            return res + "</TABLE>";
        }
        public static string MakeTable(double[,] matr,
            string[] arrHead, int d)
        {
            string res = "<TABLE border = 1><TR>";
            string f = "<TD>{0:F" + d.ToString() + "}";
            for (int j = 0; j < matr.GetLength(1); j++)
                res += "<TD>" + arrHead[j];
            for (int i = 0; i < matr.GetLength(0); i++)
            {
                res += "<TR>";
                for (int j = 0; j < matr.GetLength(1); j++)
                    res += string.Format(f, matr[i, j]);
            }
            return res + "</TABLE>";
        }
        public static string MakeIndexSumTable(double[,] matr,
            string[] arrHead, int d)
        {
            string res = "<TABLE border = 1 cellspacing = 0 align = center><TR><TD>i";
            string f = "<TD>{0:F" + d.ToString() + "}";
            for (int j = 0; j < matr.GetLength(1); j++)
                res += string.Format("<TD>{0}", arrHead[j]);
            double[] arrSum = new double[matr.GetLength(1)];
            for (int i = 0; i < matr.GetLength(0); i++)
            {
                res += string.Format("<TR><TD>{0}", i + 1);
                for (int j = 0; j < matr.GetLength(1); j++)
                {
                    res += string.Format(f, matr[i, j]);
                    arrSum[j] += matr[i, j];
                }
            }
            res += "<TR><TD>Итого";
            for (int j = 0; j < arrSum.Length; j++)
                res += string.Format(f, arrSum[j]);
            return res + "</TABLE>";
        }
    }
    class XY : IComparable
    {
        public int x;
        public int y;
        public XY(int x, int y)
        {
            this.x = x;
            this.y = y;
        }
        public int CompareTo(object obj)
        {
            return x.CompareTo((obj as XY).x);
        }
    }
    class XY2 : IComparable
    {
        public string x;
        public int[] y;
        public XY2(string x, int[] y)
        {
            this.x = x;
            this.y = y;
        }
        public int CompareTo(object obj)
        {
            return x.CompareTo((obj as XY2).x);
        }
    }

    public class Sample : ICloneable
    {
        protected string name = "", mark = "";
        protected double[] arrX;
        int indexMin, indexMax;
        double min, max, average, averageXPow2;
        double sum, sumXPow2, dev, devStd, sigma, sigmaStd;
        double[] arrHistY, arrHistX;    // частоты и середины интервалов
        double var;                  // коэф. вариации
        double mc1, mc2, mc3, mc4;   // центр. моменты
        double mb1, mb2, mb3, mb4;   // нач. моменты
        double asym, exc;            // асим. и эксцесс
        double asymTheor, excTheor, sigmaAsym, sigmaExc;

        public Sample() { }
        public Sample(string name, string mark, double[] arrX)
        {
            this.name = name;
            this.mark = mark;
            this.arrX = arrX;
            Calculate();
        }
        public Sample(double[] arrX)
        {
            this.arrX = arrX;
            Calculate();
        }
        public double this[int index]
        {
            get { return arrX[index]; }
            set { arrX[index] = value; }
        }
        public void Calculate()
        {
            min = double.MaxValue;
            max = double.MinValue;
            sum = 0;
            sumXPow2 = 0;
            for (int i = 0; i < arrX.Length; i++)
            {
                if (arrX[i] < min)
                {
                    min = arrX[i];
                    indexMin = i;
                }
                if (arrX[i] > max)
                {
                    max = arrX[i];
                    indexMax = i;
                }
                sum += arrX[i];
                sumXPow2 += arrX[i] * arrX[i];
            }
            average = sum / arrX.Length;
            averageXPow2 = sumXPow2 / arrX.Length;
            dev = averageXPow2 - (average * average);
            devStd = dev / arrX.Length * (arrX.Length - 1);
            sigma = (double)Math.Sqrt(dev);
            sigmaStd = (double)Math.Sqrt(devStd);

            // моменты
            mb1 = average;
            mb2 = sumXPow2 / arrX.Length;
            mc2 = dev;
            mc1 = mc3 = mc4 = 0;
            mb1 = mb2 = mb3 = mb4 = 0;
            for (int i = 0; i < arrX.Length; i++)
            {
                mb3 += arrX[i] * arrX[i] * arrX[i];
                mb4 += arrX[i] * arrX[i] * arrX[i] * arrX[i];
                mc1 += arrX[i] - mb1;
                mc3 += (arrX[i] - mb1) * (arrX[i] - mb1) * (arrX[i] - mb1);
                mc4 += (arrX[i] - mb1) * (arrX[i] - mb1) * (arrX[i] - mb1) * (arrX[i] - mb1);
            }
            mb3 /= arrX.Length;
            mb4 /= arrX.Length;
            mc1 /= arrX.Length;
            mc3 /= arrX.Length;
            mc4 /= arrX.Length;

            // асим. и эксцесс
            var = sigmaStd / average;
            asym = mc3 / (sigmaStd * sigmaStd * sigmaStd);
            exc = mc4 / (sigmaStd * sigmaStd * sigmaStd * sigmaStd) - 3;

            double n = arrX.Length;
            double g1 = mc3 / (double)Math.Pow(mc2, 1.5);
            asymTheor = (double)Math.Sqrt(n * (n - 1)) / (n - 2) * g1;
            sigmaAsym = (double)Math.Sqrt((6 * n * (n - 1)) /
                ((n - 2) * (n + 1) * (n + 3)));
            double g2 = mc4 / (mc2 * mc2) - 3;
            excTheor = (n - 1) / ((n - 2) * (n - 3)) * ((n + 1) * g2 + 6);
            sigmaExc = (double)Math.Sqrt(24 * n * (n - 2) * (n - 3) /
                ((n + 1) * (n + 1) * (n + 3) * (n + 5)));
        }
        public double[] CloneArray()
        {
            return (double[])arrX.Clone();
        }
        public double[] CloneHistY()
        {
            return (double[])arrHistY.Clone();
        }
        public double[] CloneHistX()
        {
            return (double[])arrHistX.Clone();
        }
        public int[] DropoutErrorsTau(double alpha)
        {
            ArrayList listX = new ArrayList(arrX);
            ArrayList listIndexIgnore = new ArrayList();
            while (true)
            {
                double min = double.MaxValue, max = double.MinValue;
                double sum = 0, sumXPow2 = 0;
                int indexMin = -1, indexMax = -1;
                for (int i = 0; i < listX.Count; i++)
                {
                    if (listIndexIgnore.BinarySearch(i) >= 0)
                        continue;
                    double x = (double)listX[i];
                    if (x < min)
                    {
                        min = x;
                        indexMin = i;
                    }
                    if (x > max)
                    {
                        max = x;
                        indexMax = i;
                    }
                    sum += x;
                    sumXPow2 += x * x;
                }
                int n = listX.Count - listIndexIgnore.Count;
                double average = sum / n;
                double devStd = sumXPow2 / (n - 1) - (sum / (n - 1)) * (sum / (n - 1));
                double sigmaStd = (double)Math.Sqrt(devStd);
                double tauXMin, tauXMax;
                tauXMin = (double)Math.Abs(min - average) / sigmaStd;
                tauXMax = (double)Math.Abs(max - average) / sigmaStd;
                int index;
                double tauMax;
                if (tauXMin >= tauXMax)
                {
                    index = indexMin;
                    tauMax = tauXMin;
                }
                else
                {
                    index = indexMax;
                    tauMax = tauXMax;
                }
                if (tauMax > StatTables.GetTau(alpha, n))
                {
                    listIndexIgnore.Add(index);
                    listIndexIgnore.Sort();
                }
                else
                    break;
            }
            return (int[])listIndexIgnore.ToArray(typeof(int));
        }
        public int[] DropoutErrorsStud(double alphaMin, double alphaMax)
        {
            if (alphaMin == -1)
                alphaMin = 0.001f;
            if (alphaMax == -1)
                alphaMax = 0.05f;
            ArrayList listX = new ArrayList(arrX);
            ArrayList listIndexIgnore = new ArrayList();
            while (true)
            {
                double min = double.MaxValue, max = double.MinValue;
                double sum = 0, sumXPow2 = 0;
                int indexMin = -1, indexMax = -1;
                for (int i = 0; i < listX.Count; i++)
                {
                    if (listIndexIgnore.BinarySearch(i) >= 0)
                        continue;
                    double x = (double)listX[i];
                    if (x < min)
                    {
                        min = x;
                        indexMin = i;
                    }
                    if (x > max)
                    {
                        max = x;
                        indexMax = i;
                    }
                    sum += x;
                    sumXPow2 += x * x;
                }
                int n = listX.Count - listIndexIgnore.Count;
                double average = sum / n;
                double devStd = sumXPow2 / (n - 1) - (sum / (n - 1)) * (sum / (n - 1));
                double sigmaStd = (double)Math.Sqrt(devStd);
                double tauXMin, tauXMax;
                tauXMin = (double)Math.Abs(min - average) / sigmaStd;
                tauXMax = (double)Math.Abs(max - average) / sigmaStd;
                int index;
                double tauMax;
                if (tauXMin >= tauXMax)
                {
                    index = indexMin;
                    tauMax = tauXMin;
                }
                else
                {
                    index = indexMax;
                    tauMax = tauXMax;
                }

                double tMin = StatTables.GetStudDistrInv(listX.Count - 2, alphaMin);
                double tMax = StatTables.GetStudDistrInv(listX.Count - 2, alphaMax);
                double critMin = tMin * (double)Math.Sqrt(listX.Count - 1) /
                    (double)Math.Sqrt(listX.Count - 2 + tMin * tMin);
                double critMax = tMax * (double)Math.Sqrt(listX.Count - 1) /
                    (double)Math.Sqrt(listX.Count - 2 + tMin * tMin);
                if (tauMax > critMin)
                {
                    listIndexIgnore.Add(index);
                    listIndexIgnore.Sort();
                }
                else
                    break;
            }
            return (int[])listIndexIgnore.ToArray(typeof(int));
        }
        public void DoHistogram(bool useSturgess)
        {
            int k;
            if (useSturgess)
                k = 1 + (int)(3.32 * Math.Log10(arrX.Length));
            else
                k = (int)Math.Sqrt(arrX.Length);
            if (k < 6 && arrX.Length >= 6)
                k = 6;
            else if (k > 20)
                k = 20;

            arrHistY = new double[k];
            double h = (max - min) / k;
            for (int i = 0; i < arrX.Length; i++)
            {
                int j;
                for (j = 0; j < arrHistY.Length - 1; j++)
                    if (arrX[i] < min + h * (j + 1))
                        break;
                arrHistY[j]++;
            }

            arrHistX = new double[k];
            for (int i = 0; i < arrHistX.Length; i++)
                arrHistX[i] = min + h * (i + 0.5);
        }
        public void AddValue(double value)
        {
            double[] arrXNew = new double[arrX.Length + 1];
            arrX.CopyTo(arrXNew, 0);
            arrXNew[arrX.Length] = value;
            arrX = arrXNew;
        }
        public void RemoveValues(int[] arrIndex)
        {
            double[] arrXNew = new double[arrX.Length - arrIndex.Length];
            ArrayList listIndex = new ArrayList(arrIndex);
            listIndex.Sort();
            for (int i = 0, j = 0; i < arrX.Length; i++)
            {
                if (listIndex.BinarySearch(i) >= 0)
                    continue;
                arrXNew[j++] = arrX[i];
            }
            arrX = arrXNew;
        }
        public string GetName()
        {
            return name;
        }
        public virtual string GetMark()
        {
            return mark;
        }
        public double GetAverage()
        {
            return average;
        }
        public double GetSigma()
        {
            return sigma;
        }
        public double GetDevStd()
        {
            return devStd;
        }
        public int GetLength()
        {
            return arrX.Length;
        }
        public string GetReport()
        {
            int d = 3;
            string s = string.Format("<P>ВЫБОРОЧНЫЕ ХАРАКТЕРИСТИКИ {0} [={1}]<BR>" +
                "Минимум: {1}<SUB>min</SUB> = {2}<BR>" +
                "Максимум: {1}<SUB>max</SUB> = {3}<BR>" +
                "Размах выборки: w = {4}<BR>" +
                "Среднее: {1}<SUB>ср</SUB> = {5}<BR>" +
                "Средний квадрат: {1}<SUP>2</SUP><SUB>ср</SUB> = {6}<BR>" +
                "Дисперсия: s<SUP>2</SUP> = {7}<BR>" +
                "Среднее квадр. откл.: s = {8}<BR>" +
                "Испр. дисперсия: s<SUP>2</SUP><SUB>испр</SUB> = {9}<BR>" +
                "Испр. среднее квадр. откл.: s<SUB>испр</SUB> = {10}<BR>" +
                "Асимметрия: A = {11}<BR>" +
                "Эксцесс: E = {12}<BR>" +
                "Коэффициент вариации: v = {13}<BR></P>",
                name, mark, Math.Round(min, d), Math.Round(max, d),
                Math.Round(min - max, d), Math.Round(average, d),
                Math.Round(averageXPow2, d), Math.Round(dev, d), Math.Round(sigma, d),
                Math.Round(devStd, d), Math.Round(sigmaStd, d), Math.Round(asym, d),
                Math.Round(exc, d), Math.Round(var, d));
            return s;
        }
        public string GetHypReport()
        {
            int d = 3;
            string asymRepr = "|A| < 3s<SUB>A</SUB> => " +
                "гипотеза о нормальности не отвергается";
            string excRepr = "|E| < 5s<SUB>E</SUB> => " +
                "гипотеза о нормальности не отвергается";
            string varRepr = "v < 0.33 => гипотеза о нормальности не отвергается";
            if (Math.Abs(asym) >= 3 * sigmaAsym)
                asymRepr = "|A| ≥ 3s<SUB>A</SUB> => " +
                    "гипотеза о нормальности отвергается";
            if (Math.Abs(exc) >= 5 * sigmaExc)
                excRepr = "|E| ≥ 5s<SUB>E</SUB> => " +
                    "гипотеза о нормальности отвергается";
            if (var >= 0.33)
                varRepr = "v ≥ 0.33 => гипотеза о нормальности отвергается";

            string s = string.Format("<P>ПРИЗНАК: {0} [={1}]<BR>", name, mark);
            s += string.Format("<P>Показатель асимметрии:<BR>" +
            "наблюдаемый: A = {0}<BR>" +
            "ошибка: s<SUB>A</SUB> = {1}<BR>{2}</P>",
            Math.Round(asym, 3), Math.Round(sigmaAsym, d), asymRepr);
            s += string.Format("<P>Показатель эксцесса:<BR>" +
                "наблюдаемый: E = {0}<BR>" +
                "ошибка: s<SUB>E</SUB> = {1}<BR>{2}</P>",
                Math.Round(exc, d), Math.Round(sigmaExc, d), excRepr);
            s += string.Format("<P>Коэффициент вариации:<BR>" +
                "наблюдаемый: v = {0}<BR>{1}</P>",
                Math.Round(var, d), varRepr);
            return s;
        }
        public void SetName(string name)
        {
            this.name = name;
        }
        public object Clone()
        {
            return new Sample(name, mark, arrX);
        }
    }
    public class TranSample : Sample
    {
        string transform;
        public TranSample(Sample s, string transform)
        {
            this.transform = transform;
            arrX = new double[s.GetLength()];
            name = s.GetName();
            mark = s.GetMark();
            switch (transform)
            {
                case "x^2":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = s[i] * s[i];
                    break;
                case "x^3":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = s[i] * s[i] * s[i];
                    break;
                case "1/x":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = 1 / s[i];
                    break;
                case "1/x^2":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = 1 / (s[i] * s[i]);
                    break;
                case "sqrt(x)":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = Math.Sqrt(s[i]);
                    break;
                case "1/sqrt(x)":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = 1 / Math.Sqrt(s[i]);
                    break;
                case "ln(x)":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = Math.Log(s[i]);
                    break;
                case "exp(x)":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = Math.Exp(s[i]);
                    break;
                case "x":
                    for (int i = 0; i < arrX.Length; i++)
                        arrX[i] = s[i];
                    break;
                default:
                    throw new Exception();
            }
            for (int i = 0; i < arrX.Length; i++)
                if (double.IsInfinity(arrX[i]) || double.IsNaN(arrX[i]))
                {
                    this.transform = "x";
                    arrX = s.CloneArray();
                }
            Calculate();
        }
        public string GetTransform()
        {
            return transform;
        }
        public override string ToString()
        {
            return transform;
        }
        public override string GetMark()
        {
            switch (transform)
            {
                case "x^2":
                    return string.Format("{0}<sup>2</sup>", base.GetMark());
                case "x^3":
                    return string.Format("{0}<sup>3</sup>", base.GetMark());
                case "1/x":
                    return string.Format("{0}<sup>-1</sup>", base.GetMark());
                case "1/x^2":
                    return string.Format("{0}<sup>-2</sup>", base.GetMark());
                case "sqrt(x)":
                    return string.Format("{0}<sup>1/2</sup>", base.GetMark());
                case "1/sqrt(x)":
                    return string.Format("{0}<sup>-1/2</sup>", base.GetMark());
                case "ln(x)":
                    return string.Format("ln({0})", base.GetMark());
                case "exp(x)":
                    return string.Format("exp({0})", base.GetMark());
                case "x":
                    return base.GetMark();
                default:
                    throw new Exception();
            }
        }
    }
    public class Regression
    {
        double[] arrYMod;
        public double[] arrB;      // коэффициенты регрессии
        double[,] matrC;    // матр. сист. норм. ур.
        double[,] matrCInv; // обр. матр. C
        double[,] matrR;    // матрица коэф. кор.
        double[,] matrRPart;// матр. част. коэфф. кор.
        double R, sigmaR;   // коэф. множ. кор.
        double fishRegr, fishRegrTheor;    // крит. Фиш. при пров. знач. регр.
        double studR, studRTheor;   // крит. Стьюд. при пров. R
        double fishR, fishRTheor; // крит. Фиш. при пров. R
        double devEnd;      // ост. дисп.
        double[] arrSigmaB; // дисп. и ср. кв. откл. коэф. B
        double[] arrStudB;  // крит. Стьюд. при пров. знач. B
        double studBTheor;
        double[] arrStudRxy;    // крит. Стьюд. при пров. знач. парн. коэф. кор. y и xi
        double[] arrSigmaRxy;
        double studRxyTheor;
        int n, p;           // кол. набл. и кол. факторов
        bool isRegrRepr, isRReprStud, isRReprFish;  // результаты пров. гип.
        bool[] arrIsBRepr;
        bool[] arrIsRxyRepr;
        string[] arrMark;   // назв. перем. (снач. y, потом все x)

        public Regression(Sample smpY, Sample[] arrSmpX)
        {
            double[] arrY = smpY.CloneArray();
            n = arrY.Length;
            p = arrSmpX.Length;
            double[][] matrX = new double[p + 1][];
            matrX[0] = new double[n];
            for (int j = 0; j < matrX[0].Length; j++)
                matrX[0][j] = 1;
            for (int i = 1; i < p + 1; i++)
                matrX[i] = arrSmpX[i - 1].CloneArray();
            Matrix mX = Matrix.Create(matrX);
            mX.Transpose();
            Matrix mY = Matrix.Create(new double[][] { arrY });
            mY.Transpose();
            Matrix mXTr = mX.Clone();
            mXTr.Transpose();
            Matrix mB = mXTr * mX;
            matrC = mB.CopyToArray();
            mB = mB.Inverse();
            matrCInv = mB.CopyToArray();
            Matrix mXTrY = mXTr * mY;
            mB = mB * mXTrY;
            arrB = new double[mB.RowCount];
            for (int i = 0; i < mB.RowCount; i++)
                arrB[i] = mB[i, 0];

            // матр. коэф. кор.
            int smpCount = p + 1;
            ArrayList listSmp = new ArrayList(arrSmpX);
            listSmp.Insert(0, smpY);
            matrR = new double[smpCount, smpCount];
            for (int i = 0; i < smpCount; i++)
                for (int j = 0; j <= i; j++)
                {
                    if (i == j)
                    {
                        matrR[i, j] = 1;
                        continue;
                    }
                    Sample smp1 = (Sample)listSmp[i];
                    Sample smp2 = (Sample)listSmp[j];
                    double sum = 0;
                    for (int k = 0; k < n; k++)
                        sum += smp1[k] * smp2[k];
                    matrR[i, j] = (sum / n - smp1.GetAverage() *
                        smp2.GetAverage()) / (smp1.GetSigma() *
                        smp2.GetSigma());
                }
            for (int i = 0; i < smpCount; i++)
                for (int j = i + 1; j < smpCount; j++)
                    matrR[i, j] = matrR[j, i];

            // пров. знач. коэф. кор.
            arrStudRxy = new double[p];
            arrSigmaRxy = new double[p];
            for (int i = 0; i < arrStudRxy.Length; i++)
            {
                double r = matrR[0, i + 1];
                arrSigmaRxy[i] = (1 - r) / Math.Sqrt(n - 1);
                arrStudRxy[i] = r / arrSigmaRxy[i];
            }

            // матр. част. коэф. кор.
            matrRPart = new double[smpCount, smpCount];
            Matrix mR = Matrix.Create(matrR);
            double minor11 = GetMinor(mR, 0, 0);
            for (int i = 0; i < mR.RowCount; i++)
                for (int j = 0; j < i; j++)
                    matrRPart[i, j] = GetMinor(mR, 1, j) /
                        Math.Sqrt(minor11 * GetMinor(mR, j, j));
            for (int i = 0; i < smpCount; i++)
                for (int j = i + 1; j < smpCount; j++)
                    matrRPart[i, j] = matrRPart[j, i];

            // коэф. множ. кор.
            //R = Math.Sqrt(1 - devEnd / smpY.GetDevStd());
            R = Math.Sqrt(1 - mR.Determinant() / minor11);
            sigmaR = (1 - R * R) / Math.Sqrt(n - p - 1);
            fishR = R * R * (n - p - 1) /
                (1 - R * R) * p;
            studR = R / sigmaR;

            // эмп. знач. y
            arrYMod = new double[n];
            double[] arrX = new double[p];
            for (int i = 0; i < arrYMod.Length; i++)
            {
                for (int j = 0; j < arrX.Length; j++)
                    arrX[j] = arrSmpX[j][i];
                arrYMod[i] = GetRegValue(arrX);
            }

            // ост. дисп.
            devEnd = 0;
            for (int i = 0; i < n; i++)
                devEnd += (arrY[i] - arrYMod[i]) * (arrY[i] - arrYMod[i]);
            devEnd /= n - p - 1;

            // пров. знач. ур. регр. и коэф. ур. регр.
            fishRegr = smpY.GetDevStd() / devEnd;
            arrSigmaB = new double[p + 1];
            arrStudB = new double[p + 1];
            for (int i = 0; i < arrSigmaB.Length; i++)
            {
                arrSigmaB[i] = Math.Sqrt(devEnd * matrCInv[i, i]);
                arrStudB[i] = arrB[i] / arrSigmaB[i];
            }

            // названия переменных
            arrMark = new string[p + 1];
            arrMark[0] = smpY.GetMark();
            for (int i = 1; i < arrMark.Length; i++)
                arrMark[i] = arrSmpX[i - 1].GetMark();
        }
        public void CheckHypothesises(double alpha)
        {
            // знач. регр.
            fishRegrTheor = StatTables.GetFishDistrInv(n - 1, n - p - 1, alpha);
            if (fishRegr > fishRegrTheor)
                isRegrRepr = true;
            else
                isRegrRepr = false;

            // знач. коэф. регр.
            studBTheor = StatTables.GetStudDistrInv(n - p - 1, alpha);
            arrIsBRepr = new bool[p + 1];
            for (int i = 0; i < arrIsBRepr.Length; i++)
                if (arrStudB[i] > studBTheor)
                    arrIsBRepr[i] = true;
                else
                    arrIsBRepr[i] = false;

            // знач. коэф. кор.
            studRxyTheor = StatTables.GetStudDistrInv(n - 2, alpha);
            arrIsRxyRepr = new bool[p];
            for (int i = 0; i < arrIsRxyRepr.Length; i++)
                if (arrStudRxy[i] > studRxyTheor)
                    arrIsRxyRepr[i] = true;
                else
                    arrIsRxyRepr[i] = false;

            // знач. множ. коэф. кор.
            studRTheor = StatTables.GetStudDistrInv(n - p - 1, alpha);
            if (studR > studRTheor)
                isRReprStud = true;
            else
                isRReprStud = false;

            fishRTheor = StatTables.GetFishDistrInv(n - p - 1, p, alpha);
            if (fishR > fishRTheor)
                isRReprFish = true;
            else
                isRReprFish = false;

        }
        public double GetRegValue(double[] arrX)
        {
            double res = arrB[0];
            for (int i = 0; i < arrX.Length; i++)
                res += arrX[i] * arrB[i + 1];
            return res;
        }
        public double[] CloneArray()
        {
            return (double[])arrYMod.Clone();
        }
        public string GetRegReport()
        {
            int d = 3;
            string s = string.Format("<P>ПОКАЗАТЕЛЬ: {0}<BR>", arrMark[0]);
            s += string.Format("<P>ПОСТРОЕНИЕ УРАВНЕНИЯ РЕГРЕССИИ<BR>" +
                "Матрица системы нормальных уравнений:<BR>{1}<BR>" +
                "Обратная матрица системы нормальных уравнений:<BR>{2}<BR>" +
                "Уравнение регрессии:<BR> {3} = ", arrMark[0],
                MatrixToHtml(matrC, d), MatrixToHtml(matrCInv, d), arrMark[0]);
            s += Math.Round(arrB[0], d).ToString();
            for (int i = 1; i < arrB.Length; i++)
            {
                if (arrB[i] >= 0)
                    s += "+";
                s += Math.Round(arrB[i], d).ToString() + arrMark[i];
            }
            s += string.Format("<BR>Остаточная дисперсия: s<SUB>ост</SUB> = {0}<BR>" +
                "Коэффициент множественной корреляции: R = {1}</P>",
                Math.Round(devEnd, d), Math.Round(R, d));
            return s;
        }
        public string GetHypRegrReport()
        {
            int d = 3;
            string repr = "F ≤ F<SUB>теор</SUB> => уравнение регрессии не значимо";
            if (isRegrRepr)
                repr = "F > F<SUB>теор</SUB> => уравнение регрессии значимо";
            string s = string.Format("<P>ПОКАЗАТЕЛЬ: {0}<BR>", arrMark[0]);
            s += string.Format("<P>ПРОВЕРКА ЗНАЧИМОСТИ УРАВНЕНИЯ РЕГРЕССИИ<BR>" +
                "Наблюдаемое значение критерия Фишера:<BR>F = {0}<BR>" +
                "Теоретическое значение критерия Фишера:<BR>F<SUB>теор</SUB> = {1}<BR>" +
                "{2}</P>", Math.Round(fishRegr, d), Math.Round(fishRegrTheor, d), repr, arrMark[0]);
            s += string.Format("<P>ПРОВЕРКА ЗНАЧИМОСТИ КОЭФФИЦИЕНТОВ УРАВНЕНИЯ РЕГРЕССИИ<BR>" +
                "Теоретическое значение критерия Стьюдента:<BR>t<SUB>теор</SUB> = {0}<BR>",
                Math.Round(studBTheor, d));
            for (int i = 0; i < arrIsBRepr.Length; i++)
            {
                repr = "коэффициент не значим";
                if (arrIsBRepr[i])
                    repr = "коэффициент значим";
                s += string.Format("b<SUB>{0}</SUB>: t<SUB>набл</SUB> = {1} => {2}<BR>",
                    i, Math.Round(arrStudB[i], d), repr);
            }
            return s;
        }
        public string GetCorrReport()
        {
            int d = 3;
            string s = string.Format("<P>ПОКАЗАТЕЛЬ: {0}<BR>", arrMark[0]);
            s += string.Format("<P>КОРРЕЛЯЦИОННЫЙ АНАЛИЗ<BR>{0}",
                 MatrixToHtmlHeaders(matrR, d, arrMark),
                 MatrixToHtmlHeaders(matrRPart, d, arrMark),
                 Math.Round(R, d), Math.Round(sigmaR, d), arrMark[0]);
            return s;
        }
        public string GetHypCorrReport()
        {
            int d = 3;
            string repr = "t < t<SUB>теор</SUB> => R не значим";
            if (isRReprStud)
                repr = "t ≥ t<SUB>теор</SUB> => R значим";
            string s = string.Format("<P>ПОКАЗАТЕЛЬ: {0}<BR>", arrMark[0]);
            s += string.Format("<P>ПРОВЕРКА ЗНАЧИМОСТИ R ПО КРИТЕРИЮ СТЬЮДЕНТА<BR>" +
                "Наблюдаемое значение:<BR>t = {0}<BR>" +
                "Теоретическое значение:<BR>t<SUB>теор</SUB> = {1}<BR>" +
                "{2}</P>", Math.Round(studR, d), Math.Round(studRTheor, d), repr);

            repr = "F ≤ F<SUB>теор</SUB> => R не значим";
            if (isRReprFish)
                repr = "F > F<SUB>теор</SUB> => R значим";
            s += string.Format("<P>ПРОВЕРКА ЗНАЧИМОСТИ R ПО КРИТЕРИЮ ФИШЕРА<BR>" +
                "Наблюдаемое значение:<BR>F = {0}<BR>" +
                "Теоретическое значение:<BR>F<SUB>теор</SUB> = {1}<BR>" +
                "{2}</P>", Math.Round(fishR, d), Math.Round(fishRTheor, d), repr);

            s += string.Format("<P>ПРОВЕРКА ЗНАЧИМОСТИ КОЭФФИЦИЕНТОВ КОРРЕЛЯЦИИ<BR>" +
                "Теоретическое значение критерия Стьюдента: t<SUB>теор</SUB> = {0}<BR>",
                Math.Round(studRxyTheor, d));
            for (int i = 0; i < arrIsRxyRepr.Length; i++)
            {
                repr = "коэффициент не значим";
                if (arrIsRxyRepr[i])
                    repr = "коэффициент значим";
                s += string.Format("r<SUB>{0}{1}</SUB>: t<SUB>набл</SUB> = {2} => {3}<BR>",
                    arrMark[0], arrMark[i + 1], Math.Round(arrStudRxy[i], d), repr);
            }
            s += "</P>";
            return s;
        }
        public double GetR()
        {
            return R;
        }
        public double GetR_yx1()
        {
            return matrR[0, 1];
        }
        double GetMinor(Matrix m, int rowIndex, int colIndex)
        {
            Matrix mTmp = new Matrix(m.RowCount - 1, m.ColumnCount - 1);
            for (int i = 0; i < mTmp.RowCount; i++)
                for (int j = 0; j < mTmp.ColumnCount; j++)
                {
                    int i1 = i, j1 = j;
                    if (i >= rowIndex)
                        i1++;
                    if (j >= colIndex)
                        j1++;
                    mTmp[i, j] = m[i1, j1];
                }
            return mTmp.Determinant();
        }
        string MatrixToHtml(double[,] matr, int d)
        {
            string s = "";
            s += "<TABLE border = 1 cellspacing = 0 align = center>";
            for (int i = 0; i < matr.GetLength(0); i++)
            {
                s += "<TR>";
                for (int j = 0; j < matr.GetLength(1); j++)
                    s += "<TD>" + Math.Round(matr[i, j], d).ToString() + "</TD>";
                s += "</TR>";
            }
            s += "</TABLE>";
            return s;
        }
        string MatrixToHtmlHeaders(double[,] matr, int d, string[] arrHeader)
        {
            string s = "";
            s += "<TABLE border = 1 cellspacing = 0 align = center><TR><TD></TD>";
            for (int i = 0; i < arrMark.Length; i++)
                s += "<TD>" + arrMark[i] + "</TD>";
            s += "</TR>";
            for (int i = 0; i < matr.GetLength(0); i++)
            {
                s += "<TR><TD>" + arrMark[i] + "</TD>";
                for (int j = 0; j < matr.GetLength(1); j++)
                    s += "<TD>" + Math.Round(matr[i, j], d).ToString() + "</TD>";
                s += "</TR>";
            }
            s += "</TABLE>";
            return s;
        }
    }
    public class StatTables
    {
        static double epsilon = 0.0000001;
        static StudentsTDistribution distStud = new StudentsTDistribution();
        static FisherSnedecorDistribution distFish = new FisherSnedecorDistribution();
        static public double GetStudDistrInv(int n, double alpha)
        {
            distStud.DegreesOfFreedom = n;
            return GetX(1 - alpha, distStud);
        }
        static public double GetFishDistrInv(int n1, int n2, double alpha)
        {
            distFish.Alpha = n1;
            distFish.Beta = n2;
            return GetX(1 - alpha, distFish);
        }
        public static double GetTau(double p, double v)
        {
            int i, j;
            for (i = arrP.Length - 1; i >= 0; i--)
                if (arrP[i] >= p)
                    break;
            for (j = arrV.Length - 1; j >= 0; j--)
                if (arrV[j] <= v)
                    break;
            return arrTau[j, i];
        }
        static double GetX(double p, ContinuousDistribution dist)
        {
            double x, xNext = 1;
            do
            {
                x = xNext;
                xNext = x - (dist.CumulativeDistribution(x) - p) /
                    dist.ProbabilityDensity(x);
            }
            while (Math.Abs(x - xNext) > epsilon);
            return xNext;
        }

        static double[] arrP = new double[]
                {
                    0.10f, 0.05f, 0.025f, 0.01f
                };
        static int[] arrV = new int[]
                {
                    3, 4, 5, 6, 7, 8, 9, 10,
                    11, 12, 13, 14, 15, 16, 17, 18, 19, 20,
                    21, 22, 23
                };
        static double[,] arrTau = new double[,]
                {
                    { 1.41f, 1.41f, 1.41f, 1.41f },
                    { 1.65f, 1.69f, 1.71f, 1.72f },
                    { 1.79f, 1.87f, 1.92f, 1.96f },
                    { 1.89f, 2.00f, 2.07f, 2.13f },
                    { 1.97f, 2.09f, 2.18f, 2.27f },
                    { 2.04f, 2.17f, 2.27f, 2.37f },
                    { 2.10f, 2.24f, 2.35f, 2.46f },
                    { 2.15f, 2.29f, 2.41f, 2.54f },                    
                    { 2.19f, 2.34f, 2.47f, 1.41f },
                    { 2.23f, 2.39f, 2.52f, 1.72f },
                    { 2.26f, 2.43f, 2.56f, 1.96f },
                    { 2.30f, 2.46f, 2.60f, 2.13f },                    
                    { 2.33f, 2.49f, 2.64f, 2.80f },
                    { 2.35f, 2.52f, 2.67f, 2.84f },
                    { 2.38f, 2.55f, 2.70f, 2.87f },
                    { 2.40f, 2.58f, 2.73f, 2.90f },
                    { 2.43f, 2.60f, 2.75f, 2.93f },
                    { 2.45f, 2.62f, 2.78f, 2.96f },
                    { 2.47f, 2.64f, 2.80f, 2.98f },
                    { 2.49f, 2.66f, 2.82f, 3.01f },
                    { 2.50f, 2.68f, 2.84f, 3.03f },
                    { 2.52f, 2.70f, 2.86f, 3.05f },
                    { 2.54f, 2.72f, 2.88f, 3.07f }
                };
    }
}