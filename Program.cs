using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace CSharp01
{

    class ExcelSheetReader : System.IDisposable
    {
        /// <summary>
        /// エクセルファイルの指定したシートを2次元配列に読み込む.
        /// </summary>
        /// <param name="filePath">エクセルファイルのパス</param>
        /// <param name="sheetIndex">シートの番号 (1, 2, 3, ...)</param>
        /// <param name="startRow">最初の行 (>= 1)</param>
        /// <param name="startColmn">最初の列 (>= 1)</param>
        /// <param name="lastRow">最後の行</param>
        /// <param name="lastColmn">最後の列</param>
        /// <returns>シート情報を格納した2次元数値配列を返す．
        //public ArrayList Read(string filePath, int sheetIndex, int startRow, int startColmn, int lastRow, int lastColmn)
        public List<List<double>> Read(string filePath, int sheetIndex, int startRow, int startColmn, int lastRow, int lastColmn)
        {
            // ワークブックを開く
            if (!Open(filePath)) { return null; }

            mSheet = mWorkBook.Sheets[sheetIndex];
            mSheet.Select();

            //var arrOut = new ArrayList();
            var arrOut = new List<List<double>>();

            for (int r = startRow; r <= lastRow; r++)
            {
                // 一行読み込む
                //var row = new ArrayList();
                var row = new List<double>();
                for (int c = startColmn; c <= lastColmn; c++)
                {
                    var cell = mSheet.Cells[r, c];

                    //if (cell == null || cell.Value == null) { row.Add(""); }
                    if (cell == null || cell.value == null)
                    {
                        Console.WriteLine("おそらくファイルの最後に到達しました．");
                    }
                    else
                    {
                        row.Add(cell.Value);
                    }
                }

                arrOut.Add(row);
            }

            return arrOut;
        }

        protected bool Open(string filePath)
        {
            if (!System.IO.File.Exists(filePath))
            {
                return false;
            }

            try
            {
                mApp = new Microsoft.Office.Interop.Excel.Application();
                mApp.Visible = false;

                // filePath が相対パスのとき例外が発生するので fullPath に変換
                string fullPath = System.IO.Path.GetFullPath(filePath);
                mWorkBook = mApp.Workbooks.Open(fullPath);
            }
            catch
            {
                //Close();
                return false;
            }

            return true;
        }

        // インターフェースの実装。リソース解放処理をまとめる。
        public void Dispose()
        {
            if (mSheet != null)
            {
                Marshal.ReleaseComObject(mSheet);
                mSheet = null;
            }

            if (mWorkBook != null)
            {
                mWorkBook.Close();
                Marshal.ReleaseComObject(mWorkBook);
                mWorkBook = null;
            }

            if (mApp != null)
            {
                mApp.Quit();
                Marshal.ReleaseComObject(mApp);
                mApp = null;
            }
        }

        protected Microsoft.Office.Interop.Excel.Application mApp = null;
        protected Microsoft.Office.Interop.Excel.Workbook mWorkBook = null;
        protected Microsoft.Office.Interop.Excel.Worksheet mSheet = null;
    }

    class Program
    {
        private const int N = 41; // 分割数
        private const int M = 20; // 周波数
        private const double T = 0.02; // 周期
        private const double dt = T / (N - 1);

        static void Main(string[] args)
        {
            // ExcelSheetReaderのインスタンスを作成
            var reader = new ExcelSheetReader();

            // 'yomikomi.xlsx' の1番シートの 1A から 41B までを読む
            var sheet = reader.Read(@"yomikomi.xlsx", 1, 1, 1, 41, 2);

            // solvers
            new Program().solveFourierSeries(sheet); // フーリエ級数を求めたいアナタに！
            new Program().solveFourierTransformShortTerm(sheet); // 短期間フーリエ変換を求めたいアナタに！
            new Program().solveFourierTransformPeridicSignal(sheet); // 周期フーリエ変換を求めたいアナタに！
            new Program().solvePowerSpectrumDensity(sheet); // 電力スペクトル密度を求めたいアナタに！
        }

        private void solveFourierSeries(List<List<double>> sheet)
        {
            var copySheet = new List<List<double>>(sheet);

            double f0 = copySheet[0].Last();
            double fn = copySheet[40].Last();

            var arrOut = new List<List<double>>();

            for (int i = 0; i <= 100; i++)
            {
                var row = new List<double>();
                double an = 0;
                double bn = 0;
                for (int j = 1; j <= N - 1; j++)
                {
                    double v = copySheet[j][1];
                    an = an + v * Math.Cos(2.0 * j * i * Math.PI / N);
                    bn = bn + v * Math.Sin(2.0 * j * i * Math.PI / N);
                }
                an = (an + (f0 + fn) * 0.5) * (2.0 / N);
                bn = bn * (2.0 / N);
                row.Add(i);
                row.Add(an);
                row.Add(bn);
                arrOut.Add(row);
            }

            // CSVファイル出力
            try
            {
                // appendをtrueにすると，既存のファイルに追記
                // appendをfalseにすると，ファイルを新規作成する
                var append = false;

                // 出力用のファイルを開く
                int n = arrOut.Count;
                using (var sw = new System.IO.StreamWriter(@"FourierSeries.csv", append))
                {
                    //sw.WriteLine("{0},{1},{2}", "n", "an", "bn");
                    for (int i = 0; i < n; ++i)
                        sw.WriteLine("{0},{1},{2}", arrOut[i][0], arrOut[i][1], arrOut[i][2]);
                }
            }
            catch (System.Exception e)
            {
                // ファイルを開くのに失敗したときエラーメッセージを表示
                Console.WriteLine(e.Message);
            }
        }

        private void solveFourierTransformShortTerm(List<List<double>> sheet)
        {
            var copySheet = new List<List<double>>(sheet);

            var arrOut = new List<List<double>>();

            int MM = 5000;

            for (int i = MM; i > 0; i-=50)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(i * (j * T / (N - 1)));
                    Im = Im - f * Math.Sin(i * (j * T / (N - 1)));
                }

                Re = Re * T / (N - 1);
                Im = Im * T / (N - 1);

                row.Add(-i);
                row.Add(Re);
                row.Add(Im);
                row.Add(Math.Pow((Math.Pow(Re, 2) + Math.Pow(Im, 2)), 0.5));
                arrOut.Add(row);
            }

            for (int i = 0; i <= MM; i+=50)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(i * (j * T / (N - 1)));
                    Im = Im - f * Math.Sin(i * (j * T / (N - 1)));
                }

                Re = Re * T / (N - 1);
                Im = Im * T / (N - 1);

                row.Add(i);
                row.Add(Re);
                row.Add(Im);
                row.Add(Math.Pow((Math.Pow(Re, 2) + Math.Pow(Im, 2)), 0.5));
                arrOut.Add(row);
            }

            // CSVファイル出力
            try
            {
                // appendをtrueにすると，既存のファイルに追記
                // appendをfalseにすると，ファイルを新規作成する
                var append = false;

                // 出力用のファイルを開く
                int n = arrOut.Count;
                using (var sw = new System.IO.StreamWriter(@"FourierTransformShortTerm.csv", append))
                {
                    //sw.WriteLine("{0},{1},{2}", "n", "an", "bn");
                    for (int i = 0; i < n; ++i)
                        sw.WriteLine("{0},{1},{2},{3}", arrOut[i][0], arrOut[i][1], arrOut[i][2], arrOut[i][3]);
                }
            }
            catch (System.Exception e)
            {
                // ファイルを開くのに失敗したときエラーメッセージを表示
                Console.WriteLine(e.Message);
            }

        }

        private void solveFourierTransformPeridicSignal(List<List<double>> sheet)
        {
            var copySheet = new List<List<double>>(sheet);

            var arrOut = new List<List<double>>();

            for (int i = M; i > 0; i--)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(2 * Math.PI * i * j / N);
                    Im = Im - f * Math.Sin(2 * Math.PI * i * j / N);
                }

                Re = Re * dt;
                Im = Im * dt;

                row.Add(-i);
                row.Add(Re);
                row.Add(Im);
                row.Add(Math.Pow((Math.Pow(Re, 2.0) + Math.Pow(Im, 2.0)), 0.5) * 2.0 * Math.PI / (1.0 * T));
                arrOut.Add(row);
            }

            for (int i = 0; i <= M; i++)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(2 * Math.PI * i * j / N);
                    Im = Im - f * Math.Sin(2 * Math.PI * i * j / N);
                }

                Re = Re * dt;
                Im = Im * dt;

                row.Add(i);
                row.Add(Re);
                row.Add(Im);
                row.Add(Math.Pow((Math.Pow(Re, 2.0) + Math.Pow(Im, 2.0)), 0.5) * 2.0 * Math.PI / (1.0 * T));
                arrOut.Add(row);
            }

            // CSVファイル出力
            try
            {
                // appendをtrueにすると，既存のファイルに追記
                // appendをfalseにすると，ファイルを新規作成する
                var append = false;

                // 出力用のファイルを開く
                int n = arrOut.Count;
                using (var sw = new System.IO.StreamWriter(@"FourierTransformPeridicSignal.csv", append))
                {
                    //sw.WriteLine("{0},{1},{2}", "n", "an", "bn");
                    for (int i = 0; i < n; ++i)
                        sw.WriteLine("{0},{1},{2},{3}", arrOut[i][0], arrOut[i][1], arrOut[i][2], arrOut[i][3]);
                }
            }
            catch (System.Exception e)
            {
                // ファイルを開くのに失敗したときエラーメッセージを表示
                Console.WriteLine(e.Message);
            }
        }

        private void solvePowerSpectrumDensity(List<List<double>> sheet)
        {
            var copySheet = new List<List<double>>(sheet);

            var arrOut = new List<List<double>>();

            for (int i = M; i > 0; i--)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(2 * Math.PI * i * j / N);
                    Im = Im - f * Math.Sin(2 * Math.PI * i * j / N);
                }

                Re = Re * dt;
                Im = Im * dt;

                row.Add(-i);
                row.Add(Re);
                row.Add(Im);
                row.Add((Math.Pow(Re, 2.0) + Math.Pow(Im, 2.0)) * 2.0 * Math.PI / T);
                arrOut.Add(row);
            }

            for (int i = 0; i <= M; i++)
            {
                var row = new List<double>();

                double Re = 0;
                double Im = 0;

                for (int j = 0; j < N; j++)
                {
                    double f = copySheet[j][1];
                    Re = Re + f * Math.Cos(2 * Math.PI * i * j / N);
                    Im = Im - f * Math.Sin(2 * Math.PI * i * j / N);
                }

                Re = Re * dt;
                Im = Im * dt;

                row.Add(i);
                row.Add(Re);
                row.Add(Im);
                row.Add((Math.Pow(Re, 2.0) + Math.Pow(Im, 2.0)) * 2.0 * Math.PI / T);
                arrOut.Add(row);
            }

            // CSVファイル出力
            try
            {
                // appendをtrueにすると，既存のファイルに追記
                // appendをfalseにすると，ファイルを新規作成する
                var append = false;

                // 出力用のファイルを開く
                int n = arrOut.Count;
                using (var sw = new System.IO.StreamWriter(@"PowerSpectrumDensity.csv", append))
                {
                    //sw.WriteLine("{0},{1},{2}", "n", "an", "bn");
                    for (int i = 0; i < n; ++i)
                        sw.WriteLine("{0},{1},{2},{3}", arrOut[i][0], arrOut[i][1], arrOut[i][2], arrOut[i][3]);
                }
            }
            catch (System.Exception e)
            {
                // ファイルを開くのに失敗したときエラーメッセージを表示
                Console.WriteLine(e.Message);
            }
        }
    }
}
