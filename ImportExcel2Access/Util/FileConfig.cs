using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace ImportExcel2Access.Util
{
    public class FileConfig
    {
        public string ConfigFilePath;
        private string CurrentExePath = Assembly.GetExecutingAssembly().GetName().Name;

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern long WritePrivateProfileString(string Section, string Key, string Value, string FilePath);

        [DllImport("kernel32", CharSet = CharSet.Unicode)]
        static extern int GetPrivateProfileString(string Section, string Key, string Default, StringBuilder RetVal, int Size, string FilePath);

        /// <summary>
        /// Check file exist
        /// </summary>
        public bool IsExist
        {
            get
            {
                return File.Exists(ConfigFilePath);
            }
        }

        /// <summary>
        /// Check access file exist
        /// </summary>
        public bool IsKeyNullOrEmpty(string keyName)
        {
            string key = Read(keyName);
            return string.IsNullOrEmpty(key);
        }

        public FileConfig(string IniPath = null)
        {
            ConfigFilePath = new FileInfo(IniPath ?? CurrentExePath + ".ini").FullName;
        }

        public string Read(string Key, string Section = null)
        {
            var RetVal = new StringBuilder(255);
            GetPrivateProfileString(Section ?? CurrentExePath, Key, "", RetVal, 255, ConfigFilePath);
            return RetVal.ToString();
        }

        public void Write(string Key, string Value, string Section = null)
        {
            WritePrivateProfileString(Section ?? CurrentExePath, Key, Value, ConfigFilePath);// ConvertToShifJIS(ConfigFilePath));
        }

        public void DeleteKey(string Key, string Section = null)
        {
            Write(Key, null, Section ?? CurrentExePath);
        }

        public void DeleteSection(string Section = null)
        {
            Write(null, null, Section ?? CurrentExePath);
        }

        public bool KeyExists(string Key, string Section = null)
        {
            return Read(Key, Section).Length > 0;
        }
    }
}
