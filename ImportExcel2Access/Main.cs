using ImportExcel2Access.Business;
using ImportExcel2Access.Util;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using static ImportExcel2Access.Util.LogUtil;

namespace ImportExcel2Access
{
    public partial class Main : Form
    {
        private FileConfig fileConfig;
        private ImportDataBusiness importDataBusiness;
        private DataTable dataRawDataTable;
        private DataTable importDataTable;
        private object[,] sheetObjectArray = null;
        private BackgroundWorker bgwReadingFileExecute;


        public Main()
        {
            InitializeComponent();
            Init();
        }

        /// <summary>
        /// Load config
        /// </summary>
        private void Init()
        {
            fileConfig = new FileConfig("Setting.ini");

            importDataBusiness = new ImportDataBusiness();
            importDataTable = new DataTable();
            dataRawDataTable = new DataTable();
            bgwReadingFileExecute = new BackgroundWorker();
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
                MessageBox.Show("インポートが完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = string.Format("{0}-{1}", this.GetType().Name, MethodBase.GetCurrentMethod().Name),
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);

                MessageBox.Show(string.Format("予期せぬエラーが発生しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    txtDataSource.Focus();
                    return false;
                }

                if ((File.GetAttributes(txtDataSource.Text) & FileAttributes.ReadOnly) > 0)
                {
                    return false;
                }

                //Check file is open
                if (File.Exists(txtDataSource.Text) && Utils.FileLocked(txtDataSource.Text))
                {
                    MessageBox.Show(string.Format("インポートする前にExcelファイルを閉じてください。"), "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtDataSource.Focus();
                    return false;
                }

                //Test excel
                if (!File.Exists(txtDataSource.Text) || !GetConnection.CheckExcelConnection(txtDataSource.Text, Constant.EXCEL_TABLE_NAME))
                {
                    MessageBox.Show(string.Format("Excelファイルの読み取りに失敗しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtDataSource.Focus();
                    return false;
                }
                                
                List<string> columns = new List<string>();

                try
                {
                    sheetObjectArray = ExcelHelps.GetDataFromExcelToMatrix(GetConnection.GetExcelPath, Constant.SHEET_NAME, ref columns);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Excelファイルの読み取りに失敗しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    ErrorEntity error = new ErrorEntity()
                    {
                        ErrorMessage = ex.Message,
                        FunctionName = string.Format("{0}-{1}", this.GetType().Name, MethodBase.GetCurrentMethod().Name),
                        FilePath = Constant.LOG_FILE_PATH
                    };

                    LogUtil.Write(error);

                    return false;
                }

                //Valid columns
                if (!Utils.IsExistColumn("No", columns))
                {
                    MessageBox.Show(string.Format("【{0}】列が存在していません。", "No"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (!Utils.IsExistColumn("修正済", columns))
                {
                    MessageBox.Show(string.Format("【{0}】列が存在していません。", "修正済"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
                if (!Utils.IsExistColumn("取込", columns))
                {
                    MessageBox.Show(string.Format("【{0}】列が存在していません。", "取込"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                //Get data to valid
                try
                {
                    dataRawDataTable.Rows.Clear();
                    importDataTable.Rows.Clear();

                    dataRawDataTable = ExcelHelps.ConvertArrayToDataTable(sheetObjectArray);

                    importDataTable = importDataBusiness.GetDataFilter(dataRawDataTable, Constant.DEFAUT_COLUMNS);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Excelファイルの読み取りに失敗しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    ErrorEntity error = new ErrorEntity()
                    {
                        ErrorMessage = ex.Message,
                        FunctionName = string.Format("{0}-{1}", this.GetType().Name, MethodBase.GetCurrentMethod().Name),
                        FilePath = Constant.LOG_FILE_PATH
                    };

                    LogUtil.Write(error);

                    return false;
                }

                //Check table has no data
                if (importDataTable.Rows.Count == 0)
                {
                    MessageBox.Show(string.Format("インポートが完了しました。"), "警告", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                object rowIndex = 0;

                // 記入日, 回答期限日,納期,完パケ期限日（初回）, 回答日
                foreach (DataRow row in importDataTable.Rows)
                {
                    #region Valid Text length

                    rowIndex = row["No"];

                    //記入者
                    if (Utils.IsExistColumn("記入者", columns) && !ValidUtils.IsValidLength(row["記入者"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "記入者", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //選曲番号
                    if (Utils.IsExistColumn("選曲番号", columns) && !ValidUtils.IsValidLength(row["選曲番号"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "選曲番号", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //発注枠
                    if (Utils.IsExistColumn("発注枠", columns) && !ValidUtils.IsValidLength(row["発注枠"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "発注枠", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //曲名
                    if (Utils.IsExistColumn("曲名", columns) && !ValidUtils.IsValidLength(row["曲名"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "曲名", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //歌手名
                    if (Utils.IsExistColumn("歌手名", columns) && !ValidUtils.IsValidLength(row["歌手名"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "歌手名", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //確認箇所
                    if (Utils.IsExistColumn("確認箇所", columns) && !ValidUtils.IsValidLength(row["確認箇所"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "確認箇所", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    //回答者
                    if (Utils.IsExistColumn("回答者", columns) && !ValidUtils.IsValidLength(row["回答者"].ToString(), 255))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】は最大{2}ケタを入力してください。", "回答者", rowIndex, 255), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    #endregion

                    #region Valid number
                    if (Utils.IsExistColumn("No", columns) && !ValidUtils.IsNullOrNumber(row["No"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に数値を入力してください。", "No", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    if (Utils.IsExistColumn("選曲番号", columns) && !ValidUtils.IsNullOrNumber(row["選曲番号"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に数値を入力してください。", "選曲番号", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    if (Utils.IsExistColumn("取込", columns) && !ValidUtils.IsNullOrNumber(row["取込"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に数値を入力してください。", "取込", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    #endregion

                    #region Valid datetime
                    if (Utils.IsExistColumn("記入日", columns) && !ValidUtils.IsNullOrDateTime(row["記入日"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に日付を入力してください。", "記入日", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (Utils.IsExistColumn("回答期限日", columns) && !ValidUtils.IsNullOrDateTime(row["回答期限日"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に日付を入力してください。", "回答期限日", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (Utils.IsExistColumn("納期", columns) && !ValidUtils.IsNullOrDateTime(row["納期"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に日付を入力してください。", "納期", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (Utils.IsExistColumn("完パケ期限日（初回）", columns) && !ValidUtils.IsNullOrDateTime(row["完パケ期限日（初回）"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に日付を入力してください。", "完パケ期限日（初回）", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    if (Utils.IsExistColumn("回答日", columns) && !ValidUtils.IsNullOrDateTime(row["回答日"].ToString()))
                    {
                        MessageBox.Show(string.Format("【{0}】列の【No:{1}】に日付を入力してください。", "回答日", rowIndex), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    FunctionName = string.Format("{0}-{1}", this.GetType().Name, MethodBase.GetCurrentMethod().Name),
                    FilePath = Constant.LOG_FILE_PATH
                };

                LogUtil.Write(error);
                MessageBox.Show(string.Format("予期せぬエラーが発生しました。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
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
            if (!fileConfig.IsExist)
            {
                MessageBox.Show(string.Format("【Setting.ini】ファイルが見つかりませんでした。"), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);

                //try
                //{
                //    fileConfig.CreateSetting();
                //}
                //catch (Exception ex)
                //{
                //    ErrorEntity error = new ErrorEntity()
                //    {
                //        ErrorMessage = ex.Message,
                //        FunctionName = this.GetType().FullName,
                //        FilePath = Constant.LOG_FILE_PATH
                //    };

                //    LogUtil.Write(error);
                //    MessageBox.Show(string.Format("システムエラー: {0}", ex.Message), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //    return false;
                //}

                return false;
            }

            txtDataSource.Text = fileConfig.Read(Constant.ExcelKey, Constant.EXCEL_SECSION);

            if (!GetConnection.CheckAccessConnection())
            {
                MessageBox.Show(string.Format("Accessファイルの接続に失敗しました。", GetConnection.GetAccessPath), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void txtDataSource_TextChanged(object sender, EventArgs e)
        {
            lblWatermask.Visible = string.IsNullOrWhiteSpace(txtDataSource.Text);
        }

        private void txtDataSource_Leave(object sender, EventArgs e)
        {
            lblWatermask.Visible = string.IsNullOrWhiteSpace(txtDataSource.Text);
        }

        private void txtDataSource_Enter(object sender, EventArgs e)
        {
            lblWatermask.Visible = string.IsNullOrWhiteSpace(txtDataSource.Text);
        }

        private void txtDataSource_Click(object sender, EventArgs e)
        {
            lblWatermask.Visible = false;
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                //Write file config
                if (fileConfig.IsExist)
                    fileConfig.Write(Constant.ExcelKey, txtDataSource.Text, Constant.EXCEL_SECSION);
            }
            catch (Exception ex)
            {
                ErrorEntity error = new ErrorEntity()
                {
                    ErrorMessage = ex.Message,
                    FunctionName = string.Format("{0}-{1}", this.GetType().Name, MethodBase.GetCurrentMethod().Name),
                    FilePath = Constant.LOG_FILE_PATH
                };
                LogUtil.Write(error);
            }
        }
    }
}
