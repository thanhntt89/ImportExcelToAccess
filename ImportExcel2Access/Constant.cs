using System.Collections.Generic;
using System.Windows.Forms;

namespace ImportExcel2Access
{
    public class Constant
    {
        public const string EXCEL_SECSION = "EXCEL";
        public const string ExcelKey = "EXCEL_PATH";

        public const string ACCESS_SECSION = "ACCESS";
        public const string AccessKey = "ACCESS_PATH";

        public const string FILE_CONFIG = "Setting.ini";

        public static string LOG_FILE_PATH = string.Format("{0}\\{1}", Application.StartupPath, "system_error.log");


        /// <summary>
        /// Default columns in excel
        /// </summary>
        public static List<string> DEFAUT_COLUMNS = new List<string>()
        {
            "No","記入日","記入グループ","記入者","回答期限日","納期","完パケ期限日（初回）","発注枠","選曲番号","曲名","歌手名","確認用素材（格納場所）","発注時音源","発注時歌詞","確認箇所","問い合わせ内容","回答者","回答日","回答","対応方法","修正済","取込"
        };

        /// <summary>
        /// Start row index has data in excel table
        /// </summary>
        public const int START_INDEX = 3;

        public const string SHEET_NAME = "歌詞問い合わせ";

        /// <summary>
        /// Range of data in excel
        /// 3: Line index has data
        /// </summary>
        public static string EXCEL_TABLE_NAME = string.Format("[{0}$A2:V]", SHEET_NAME);

        public const string ACCESS_TABLE_NAME = "問い合わせ一覧";

    }
}
