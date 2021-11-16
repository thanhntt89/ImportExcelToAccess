using ImportExcel2Access.Business;
using ImportExcel2Access.Util;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using static ImportExcel2Access.Util.LogUtil;

namespace ImportExcel2Access
{
    public partial class Main : Form
    {
        private FileConfig fileConfig;
        private ImportDataBusiness importDataBusiness;
        private DataTable importDataTable;

        public Main()
        {
            InitializeComponent();
            LoadConfig();
        }

        /// <summary>
        /// Load config
        /// </summary>
        private void LoadConfig()
        {
            fileConfig = new FileConfig("Setting.ini");
            fileConfig.Reading();
            txtDataSource.Text = fileConfig.Read(Constant.ExcelKey, Constant.EXCEL_SECSION);
            importDataBusiness = new ImportDataBusiness();
            importDataTable = new DataTable();
        }

        private void btnFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "*.xlsx|*.xlsx|*.xls|*.xls";
            openFileDialog.ShowDialog();
            txtDataSource.Text = openFileDialog.FileName;
        }

        private void btnExecute_Click(object sender, EventArgs e)
        {
            Execute();
        }

        /// <summary>
        /// Execute import data
        /// </summary>
        private void Execute()
        {
            if (!Valid())
                return;
            try
            {
                importDataBusiness.ImportExecute(importDataTable);
                importDataTable.Rows.Clear();

                MessageBox.Show("インポートが完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = this.GetType().FullName,
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);

                MessageBox.Show(string.Format("システムエラー: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Valid parameters
        /// </summary>
        /// <returns></returns>
        private bool Valid()
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtDataSource.Text))
                {
                    MessageBox.Show("ファイルを選択してください。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                //Test excel
                if (!GetConnection.CheckExcelConnection(txtDataSource.Text))
                {
                    MessageBox.Show(string.Format("サーバー上にある{0} データベースに接続できません。データベースが存在し、サーバーが実行中であることを確認してください。", GetConnection.GetExcelPath), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    return false;
                }
                try
                {
                    importDataTable = importDataBusiness.GetDataFromExcel();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Excelファイルの読み取りに失敗しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    ErrorEntity error = new ErrorEntity()
                    {
                        ErrorMessage = ex.Message,
                        FunctionName = this.GetType().FullName,
                        FilePath = Constant.LOG_FILE_PATH
                    };

                    LogUtil.Write(error);

                    return false;
                }               

                //Check table has no data
                if (importDataTable.Rows.Count == 0)
                {
                    MessageBox.Show(string.Format("インポートのデータがありません。"), "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return false;
                }

                List<string> columns = Utils.GetColumns(importDataTable);

                // Get exception columns
                var excepColumns = columns.Except(Constant.DEFAUT_COLUMNS).ToList();

                if (excepColumns.Count > 0)
                {
                    string colums = string.Join(", ", excepColumns.ToArray()).Trim();

                    MessageBox.Show(string.Format("{0} は無効です。", colums), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // 記入日, 回答期限日,納期,完パケ期限日（初回）, 回答日
                foreach (DataRow row in importDataTable.Rows)
                {
                    #region Valid Text length
                    //記入者
                    if (!ValidUtils.IsValidLength(row["記入者"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "記入者", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //選曲番号
                    if (!ValidUtils.IsValidLength(row["選曲番号"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "選曲番号", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //発注枠
                    if (!ValidUtils.IsValidLength(row["発注枠"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "発注枠", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //曲名
                    if (!ValidUtils.IsValidLength(row["曲名"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "曲名", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //歌手名
                    if (!ValidUtils.IsValidLength(row["歌手名"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "歌手名", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //確認箇所
                    if (!ValidUtils.IsValidLength(row["確認箇所"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "確認箇所", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //回答者
                    if (!ValidUtils.IsValidLength(row["回答者"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("No {0} {1}は {2} 文字以下で入力してください。", row["No"], "回答者", 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    #endregion

                    #region Valid number
                    if (!ValidUtils.IsNullOrNumber(row["選曲番号"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は数値を入力してください。", row["No"], "選曲番号"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    #endregion

                    #region Valid datetime
                    if (!ValidUtils.IsNullOrDateTime(row["記入日"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は日付を入力してください。", row["No"], "記入日"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (!ValidUtils.IsNullOrDateTime(row["回答期限日"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は日付を入力してください。", row["No"], "回答期限日"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (!ValidUtils.IsNullOrDateTime(row["納期"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は日付を入力してください。", row["No"], "納期"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (!ValidUtils.IsNullOrDateTime(row["完パケ期限日（初回）"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は日付を入力してください。", row["No"], "完パケ期限日（初回）"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (!ValidUtils.IsNullOrDateTime(row["回答日"].ToString()))
                    {
                        MessageBox.Show(string.Format("No{0}{1}は日付を入力してください。", row["No"], "回答日"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    #endregion
                }

                return true;
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = this.GetType().FullName,
                    FilePath = Constant.LOG_FILE_PATH
                };

                LogUtil.Write(error);

                MessageBox.Show(string.Format("システムエラー: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                fileConfig.Write(Constant.ExcelKey, txtDataSource.Text, Constant.EXCEL_SECSION);
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = this.GetType().FullName,
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);
            }
        }

        private void Main_Load(object sender, EventArgs e)
        {
            if (!CheckDatabaseConnection())
            {
                this.Close();
            }
        }

        /// <summary>
        /// Check database connection
        /// </summary>
        /// <returns></returns>
        private bool CheckDatabaseConnection()
        {
            if (!GetConnection.CheckAccessConnection())
            {
                MessageBox.Show(string.Format("メッセージ：{0}　Accessファイルの接続に失敗しました。", GetConnection.GetAccessPath), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }
    }
}
