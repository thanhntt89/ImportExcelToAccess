using System;
using System.Globalization;
using System.Linq;

namespace ImportExcel2Access.Util
{
    public class ValidUtils
    {
        private static CultureInfo provider = CultureInfo.InvariantCulture;

        /// <summary>
        /// Valid date time input ex: MM/dd/yyyy
        /// </summary>
        /// <param name="dateString"></param>
        /// <returns></returns>
        public static bool IsNullOrDateTime(string dateString)
        {
            DateTime result;

            try
            {
                if (string.IsNullOrWhiteSpace(dateString))
                    return true;

                return DateTime.TryParse(dateString, out result);

                //string[] data = dateString.Split('/');
                //int yyyy = int.Parse(data[2].Split(' ')[0]);
                //int MM = int.Parse(data[0]);
                //int dd = int.Parse(data[1]);

                //result = new DateTime(yyyy, MM, dd);

                // return true;               
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Valid string is null or number
        /// </summary>
        /// <param name="numberString"></param>
        /// <returns>True if string is null or number | False if string is not number</returns>
        public static bool IsNullOrNumber(string numberString)
        {
            if (string.IsNullOrWhiteSpace(numberString))
                return true;
            return numberString.Replace("-", "").Replace("+", "").All(char.IsDigit);
        }

        /// <summary>
        /// Valid length
        /// </summary>
        /// <param name="textString"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        public static bool IsValidLength(string textString, int maxLength)
        {
            return textString.Length < maxLength;
        }
    }
}
