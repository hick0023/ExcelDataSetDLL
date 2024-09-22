using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDataSetDLL
{

    /// <summary>
    /// エクセルファイルーDataSet制御（クラス）
    /// </summary>
    public class ExcelDataSet
    {
        private DataSet DataSet { get; set; }
        private String ExcelFile = null;
        private Excel.Application Application { get; }
        private Excel.Workbook Workbook { get; set; }

        /// <summary>
        /// エクセルファイルーDataSet制御（コンストラクター）
        /// </summary>
        public ExcelDataSet()
        {
            DataSet = new DataSet();
            this.Application = new Excel.Application();
            this.Application.Visible = false;
            this.Application.DisplayAlerts = false;
        }

        /// <summary>
        /// エクセルファイルーDataSet制御（コンストラクター）
        /// </summary>
        /// <param name="filename">エクセルファイル名</param>
        public ExcelDataSet(String filename)
        {
            DataSet = new DataSet();
            this.ExcelFile = filename;
            this.Application = new Excel.Application();
            this.Application.Visible = false;
            this.Application.DisplayAlerts = false;
            this.Workbook = this.Application.Workbooks.Open(this.ExcelFile);
        }

        /// <summary>
        /// デストラクター
        /// </summary>
        ~ExcelDataSet()
        {
            if (this.Workbook != null)
            {
                this.Workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Workbook);
                this.Workbook = null;
            }
            this.Application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Application);
        }

        /// <summary>
        /// エクセルブック（ファイル）を開く
        /// </summary>
        /// <param name="filename">エクセルファイル名</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool OpenWorkbook(String filename)
        {
            bool result = false;
            if (this.Workbook == null)
            {
                this.ExcelFile = filename;
                this.Workbook = this.Application.Workbooks.Open(this.ExcelFile);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// エクセルブック（ファイル）を閉じる
        /// </summary>
        /// <param name="seve">true：保存する</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool CloseWorkbook(bool seve = false)
        {
            bool result = false;
            if (this.Workbook != null)
            {
                if (seve)
                {
                    this.Workbook.Save();
                }
                this.Workbook.Close();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(this.Workbook);
                this.Workbook = null;
                result = true;
            }
            return result;
        }

        /// <summary>
        /// エクセルシート内のデータをシートとしてデータセットに格納する。
        /// </summary>
        /// <param name="SheetName">エクセルデータシート名</param>
        /// <param name="StartRow">データ開始列</param>
        /// <param name="ColumnInfo">格納するデータ列情報。[列名, [エクセル列記号, データタイプ]]のフォーマット</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool SetSheetContentsToDataSet(String SheetName, int StartRow, Dictionary<String, KeyValuePair<String, Type>> ColumnInfo)
        {
            bool result = false;
            if (this.Workbook != null) //ワークブックがOPENしている場合
            {
                if (!this.DataSet.Tables.Contains(SheetName)) //データセットに同名のTeableが存在しない場合。
                {
                    Excel.Worksheet worksheet = this.Workbook.Worksheets[SheetName] as Excel.Worksheet;
                    DataTable dt = new DataTable();
                    dt.TableName = SheetName; // テーブル名はシート名とする。
                    // カラムの情報セット //
                    foreach (KeyValuePair<String, KeyValuePair<String, Type>> keyValue in ColumnInfo)
                    {
                        dt.Columns.Add(keyValue.Key, keyValue.Value.Value);
                    }
                    int i = 0;
                    int col_num_init = ColumnStrToInt(ColumnInfo.Values.ToArray()[0].Key);
                    // データの読み込み //
                    while (worksheet.Cells[StartRow + i, col_num_init].value != null)
                    {
                        DataRow dataRow = dt.NewRow();
                        foreach (KeyValuePair<String, KeyValuePair<String, Type>> keyValue in ColumnInfo)
                        {
                            int col_num = ColumnStrToInt(keyValue.Value.Key);
                            String data = worksheet.Cells[StartRow + i, col_num].value.ToString(); //一旦全てStringで読み取る。
                            Type dtype = dataRow.Table.Columns[dataRow.Table.Columns.IndexOf(keyValue.Key.ToString())].DataType;
                            //*** カラムの情報でのデータ型に合わせてデータをTableにセットする。 ***//
                            switch (dtype.Name)
                            {
                                case "String":
                                    dataRow[keyValue.Key] = data;
                                    break;
                                case "Int32":
                                    dataRow[keyValue.Key] = Int32.Parse(data);
                                    break;
                                case "Single":
                                    dataRow[keyValue.Key] = Single.Parse(data);
                                    break;
                                case "Double":
                                    dataRow[keyValue.Key] = Double.Parse(data);
                                    break;
                                case "Boolean":
                                    dataRow[keyValue.Key] = Boolean.Parse(data);
                                    break;
                                case "DateTime":
                                    dataRow[keyValue.Key] = DateTime.Parse(data);
                                    break;
                                default:
                                    break;
                            }
                        }
                        dt.Rows.Add(dataRow);
                        i++;
                    }
                    this.DataSet.Tables.Add(dt);
                }
                result = true;
            }
            return result;
        }

        /// <summary>
        /// データセットからテーブルを削除する。
        /// </summary>
        /// <param name="TableName">テーブル名</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool DeleteTable(String TableName)
        {
            bool result = false;
            if (this.DataSet.Tables.Contains(TableName))
            {
                this.DataSet.Tables.Remove(TableName);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// エクセルシートにデータセットからテーブルを書き出す。
        /// </summary>
        /// <param name="TableName">データセットのテーブル名</param>
        /// <returns>ture：成功 / false：失敗</returns>
        /// <remarks>事前にデータセット内のテーブル名はエクセルブック（ファイル）に無い名前に変更必要</remarks>
        public bool WriteTebleToExcel(String TableName)
        {
            bool result = false;
            if (this.Workbook != null)
            {
                if (this.DataSet.Tables.Contains(TableName))
                {
                    if (!SheetExists(TableName))
                    {
                        Excel.Worksheet newSheet = this.Workbook.Sheets.Add();
                        newSheet.Name = TableName;
                        DataTable work_tabel = this.DataSet.Tables[TableName];
                        int r = 0;
                        foreach (DataRow row in work_tabel.Rows)
                        {
                            int c = 0;
                            foreach (DataColumn column in work_tabel.Columns)
                            {
                                this.Workbook.Sheets[TableName].cells[r + 1, c + 1] = row[column].ToString();
                                c++;
                            }
                            r++;
                        }
                        result = true;
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// データセット内のテーブルをコピーして新しいテーブルを追加する。
        /// </summary>
        /// <param name="InitTable">コピー元のテーブル名</param>
        /// <param name="DestinationTable">コピー先のテーブル名</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool CopyTable(String InitTable, String DestinationTable)
        {
            bool result = false;
            if (this.DataSet.Tables.Contains(InitTable))
            {
                DataTable new_dt = this.DataSet.Tables[InitTable].Copy();
                new_dt.TableName = DestinationTable;
                this.DataSet.Tables.Add(new_dt);
                result = true;
            }
            return result;
        }

        /// <summary>
        /// データセット内のテーブル名を変更する。
        /// </summary>
        /// <param name="InitTable">元のテーブル名</param>
        /// <param name="DestinationTable">変更後のテーブル名</param>
        /// <returns>ture：成功 / false：失敗</returns>
        public bool RenameTable(String InitTable, String DestinationTable)
        {
            bool result = false;
            if (this.DataSet.Tables.Contains(InitTable))
            {
                this.DataSet.Tables[InitTable].TableName = DestinationTable;
                result = true;
            }
            return result;
        }

        /// <summary>
        /// データセットをオブジェクトとして取得する。
        /// </summary>
        /// <returns>データセット</returns>
        public DataSet GetDataSetObj()
        {
            return this.DataSet;
        }

        /// <summary>
        /// エクセル列記号を番号に変換する。
        /// </summary>
        /// <param name="ColumnStr">エクセル列記号</param>
        /// <returns>エクセル列番号</returns>
        private int ColumnStrToInt(String ColumnStr)
        {
            int out_int = 0;
            int base_num = Convert.ToInt32('A') - 1;
            int digit_num = Convert.ToInt32('Z') - base_num;
            int col_str_length = ColumnStr.Length;
            int i = 0;
            foreach (char item in ColumnStr)
            {
                int item_num = Convert.ToInt32(item) - base_num;
                out_int += item_num * (int)Math.Pow(digit_num, col_str_length - 1 - i);
                i++;
            }
            return out_int;
        }

        /// <summary>
        /// エクセルブック（ファイル）に一致するシート名が存在するか確認する。
        /// </summary>
        /// <param name="SheetName">エクセルシート名</param>
        /// <returns>ture：ある / false：ない</returns>
        private bool SheetExists(String SheetName)
        {
            bool result = false;
            foreach (Excel.Worksheet sheet in this.Workbook.Sheets)
            {
                if (sheet.Name == SheetName)
                {
                    result = true;
                }
            }
            return result;
        }
    }
}
