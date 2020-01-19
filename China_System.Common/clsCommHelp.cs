using Microsoft.Win32;
using SDZdb;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace China_System.Common
{
    public class clsCommHelp
    {
        #region NullToString
        public static string NullToString(object obj)
        {
            string strResult = "";
            if (obj != null)
            {
                strResult = obj.ToString().Trim();
            }
            return strResult;
        }
        #endregion

        #region StringToDecimal
        /// <summary>
        /// 转换字符串，将字符串转换成数字，并且将空字符串转换成0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static decimal StringToDecimal(string s)
        {
            decimal result = 0;

            if (s != null && s != "")
            {
                result = Decimal.Parse(s);
            }
            return result;
        }
        #endregion

        #region StringToInt
        /// <summary>
        /// 转换字符串，将字符串转换成数字，并且将空字符串转换成0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int StringToInt(string s)
        {
            int result = 0;

            if (s != null && s != "")
            {
                result = Convert.ToInt32(s.Trim());
            }
            return result;
        }
        #endregion

        #region 日期转换(objToDateTime)
        /// <summary>
        /// 将excel里取得的日期转化成String数据
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string objToDateTime<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("MM/dd/yyyy");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("MM/dd/yyyy");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(0, 4) + "-" +
                                                       obj.ToString().Substring(4, 2) + "-" +
                                                       obj.ToString().Substring(6, 2)).ToString("MM/dd/yyyy");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            return strResult;
        }

        public static string objToDateTime1<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("yyyy/MM/dd");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("yyyy/MM/dd");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(4, 4) + "-" +
                                                       obj.ToString().Substring(0, 2) + "-" +
                                                       obj.ToString().Substring(2, 2)).ToString("yyyy/MM/dd");
                        }
                    }
                    catch (Exception ex)
                    {
                        try
                        {
                            if (obj.ToString().Length == 8)
                            {
                                strResult = DateTime.Parse(obj.ToString().Substring(0, 4) + "-" +
                                                           obj.ToString().Substring(4, 2) + "-" +
                                                           obj.ToString().Substring(6, 2)).ToString("yyyy/MM/dd");
                            }

                        }
                        catch (Exception ex1)
                        {
                            
                            throw ex1;
                        }
                       
                    }
                }
            }

            return strResult;
        }
        public static string objToDateTime2<T>(T t)
        {
            string strResult = "";
            object obj = t;

            try
            {
                if (obj != null)
                {
                    strResult = DateTime.FromOADate((double)obj).ToString("yyyy/MM/dd/HH/mm");
                }
            }
            catch
            {
                try
                {
                    strResult = Convert.ToDateTime(obj.ToString()).ToString("yyyy/MM/dd/HH/mm");
                }
                catch
                {
                    try
                    {
                        if (obj.ToString().Length == 8)
                        {
                            strResult = DateTime.Parse(obj.ToString().Substring(4, 4) + "-" +
                                                       obj.ToString().Substring(0, 2) + "-" +
                                                       obj.ToString().Substring(2, 2)).ToString("yyyy/MM/dd/HH/mm");
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                }
            }

            return strResult;
        }
        #endregion

        #region 字符串简单加密解密

        /// <summary>
        /// 简单加密解密

        /// </summary>
        /// <param name="str">需要加密、解密的字符串</param>
        /// <returns>加密、解密后的字符串</returns>
        public static string encryptString(string str)
        {
            string strResult = "";
            char[] charMessage = str.ToCharArray();
            foreach (char c in charMessage)
            {
                char newChar = changerChar(c);
                strResult += newChar.ToString();
            }
            return strResult;
        }

        private static char changerChar(char c)
        {
            char resutlt;
            int intStrLength = 0;
            string twoString = Convert.ToString(c, 2).PadLeft(8, '0');
            if (twoString.Length > 8)
            {
                twoString = Convert.ToString(c, 2).PadLeft(16, '0');
            }
            intStrLength = twoString.Length;
            string newTwoString = twoString.Substring(intStrLength / 2) + twoString.Substring(0, intStrLength / 2);
            resutlt = Convert.ToChar(Convert.ToInt32(newTwoString, 2));
            return resutlt;
        }
        #endregion

        #region 将字符串日期转换为时间类型

        public static DateTime GetDateByString(string dateString)
        {
            return DateTime.Parse(dateString.Substring(0, 4) + "-" + dateString.Substring(4, 2) + "-" + dateString.Substring(6, 2));
        }
        #endregion

        #region 关闭打开的Excel
        public static void CloseExcel(Microsoft.Office.Interop.Excel.Application ExcelApplication, Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook)
        {
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing);
            ExcelWorkbook = null;
            ExcelApplication.Quit();
            GC.Collect();
            clsKeyMyExcelProcess.Kill(ExcelApplication);
        }
        #endregion

        #region 得到Sap连接字符串

        #endregion

        #region 判断2个日期是否是整年

        public static bool CheckThroughoutTheYear(string data1, string date2)
        {
            bool blnResult = false;
            string dtStart = "";
            string dtEnd = "";
            if (Convert.ToDateTime(date2).CompareTo(Convert.ToDateTime(data1)) > 0)
            {
                dtStart = data1;
                dtEnd = date2;
            }
            else
            {
                dtStart = date2;
                dtEnd = data1;
            }
            string strTheoryDate = Convert.ToDateTime(dtEnd).ToString("yyyy")
                                 + Convert.ToDateTime(dtStart).ToString("MMdd");
            strTheoryDate = Convert.ToDateTime(objToDateTime<string>(strTheoryDate)).AddDays(-1).ToString("MM/dd/yyyy");
            if (objToDateTime<string>(strTheoryDate) == objToDateTime<string>(dtEnd))
            {
                blnResult = true;
            }
            return blnResult;
        }

        #endregion

        #region MyRegion
         #region 排序
        public class Comp : Comparer<clCard_info>
        {
            public override int Compare(clCard_info iten1, clCard_info item)
            {
                #region 判断是否为汉字
                if (iten1.daima_gonghao != null && iten1.daima_gonghao != "")
                {
                    char[] c = iten1.daima_gonghao.ToCharArray();
                    bool ischina = false;

                    for (int i = 0; i < c.Length; i++)
                    {
                        if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                            ischina = true;
                    }

                    if (ischina == true)//|| Regex.Matches(iten1.daima_gonghao, "[a-zA-Z]").Count > 0
                    {
                        return 0;
                    }
                }
                else
                    return 0;

                if (iten1.daima_gonghao != null && iten1.daima_gonghao != "")
                {
                    char[] c = item.daima_gonghao.ToCharArray();
                    bool ischina = false;
                    for (int i = 0; i < c.Length; i++)
                    {
                        if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                            ischina = true;
                    }
                    if (ischina == true)//|| Regex.Matches(item.daima_gonghao, "[a-zA-Z]").Count > 0
                    {
                        return 0;
                    }
                }
                else
                    return 0;
                #endregion
                if (iten1.daima_gonghao.Length > 10 || item.daima_gonghao.Length > 10)
                {
                    return 0;

                }
                if (item.daima_gonghao == null && item.daima_gonghao == "")
                {
                    //  item.DO_NO = "1";
                    //  return 0;
                    if (iten1.daima_gonghao == null || !iten1.daima_gonghao.Contains("DO"))
                        return int.Parse("0") - int.Parse("0");

                    return int.Parse("0") - int.Parse("0");
                }
                return int.Parse(System.Text.RegularExpressions.Regex.Replace(item.daima_gonghao, @"[^0-9]+", "")) - int.Parse(System.Text.RegularExpressions.Regex.Replace(iten1.daima_gonghao, @"[^0-9]+", ""));
                ;

            }
        }
        public class Comp1 : Comparer<clCard_info>
        {
            public override int Compare(clCard_info iten1, clCard_info item)
            {
                #region 判断是否为汉字
                if (iten1.Order_id != null && iten1.Order_id != "")
                {
                    char[] c = iten1.Order_id.ToCharArray();
                    bool ischina = false;

                    for (int i = 0; i < c.Length; i++)
                    {
                        if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                            ischina = true;
                    }

                    if (ischina == true)//|| Regex.Matches(iten1.daima_gonghao, "[a-zA-Z]").Count > 0
                    {
                        return 0;
                    }
                }
                else
                    return 0;

                if (iten1.Order_id != null && iten1.Order_id != "")
                {
                    char[] c = item.Order_id.ToCharArray();
                    bool ischina = false;
                    for (int i = 0; i < c.Length; i++)
                    {
                        if (c[i] >= 0x4e00 && c[i] <= 0x9fbb)
                            ischina = true;
                    }
                    if (ischina == true)//|| Regex.Matches(item.daima_gonghao, "[a-zA-Z]").Count > 0
                    {
                        return 0;
                    }
                }
                else
                    return 0;
                #endregion
                if (iten1.Order_id.Length > 10 || item.Order_id.Length > 10)
                {
                    return 0;

                }
                if (item.Order_id == null && item.Order_id == "")
                {
                    //  item.DO_NO = "1";
                    //  return 0;
                    if (iten1.Order_id == null || !iten1.Order_id.Contains("DO"))
                        return int.Parse("0") - int.Parse("0");

                    return int.Parse("0") - int.Parse("0");
                }
                return int.Parse(System.Text.RegularExpressions.Regex.Replace(item.Order_id, @"[^0-9]+", "")) - int.Parse(System.Text.RegularExpressions.Regex.Replace(iten1.Order_id, @"[^0-9]+", ""));
                ;

            }
        }
   
        public class SortableBindingList<T> : BindingList<T>
        {
            private bool isSortedCore = true;
            private ListSortDirection sortDirectionCore = ListSortDirection.Ascending;
            private PropertyDescriptor sortPropertyCore = null;
            private string defaultSortItem;

            public SortableBindingList() : base() { }

            public SortableBindingList(IList<T> list) : base(list) { }

            protected override bool SupportsSortingCore
            {
                get { return true; }
            }

            protected override bool SupportsSearchingCore
            {
                get { return true; }
            }

            protected override bool IsSortedCore
            {
                get { return isSortedCore; }
            }

            protected override ListSortDirection SortDirectionCore
            {
                get { return sortDirectionCore; }
            }

            protected override PropertyDescriptor SortPropertyCore
            {
                get { return sortPropertyCore; }
            }

            protected override int FindCore(PropertyDescriptor prop, object key)
            {
                for (int i = 0; i < this.Count; i++)
                {
                    if (Equals(prop.GetValue(this[i]), key)) return i;
                }
                return -1;
            }

            protected override void ApplySortCore(PropertyDescriptor prop, ListSortDirection direction)
            {
                isSortedCore = true;
                sortPropertyCore = prop;
                sortDirectionCore = direction;
                Sort();
            }

            protected override void RemoveSortCore()
            {
                if (isSortedCore)
                {
                    isSortedCore = false;
                    sortPropertyCore = null;
                    sortDirectionCore = ListSortDirection.Ascending;
                    Sort();
                }
            }

            public string DefaultSortItem
            {
                get { return defaultSortItem; }
                set
                {
                    if (defaultSortItem != value)
                    {
                        defaultSortItem = value;
                        Sort();
                    }
                }
            }

            private void Sort()
            {
                List<T> list = (this.Items as List<T>);
                list.Sort(CompareCore);
                ResetBindings();
            }

            private int CompareCore(T o1, T o2)
            {
                int ret = 0;
                if (SortPropertyCore != null)
                {
                    ret = CompareValue(SortPropertyCore.GetValue(o1), SortPropertyCore.GetValue(o2), SortPropertyCore.PropertyType);
                }
                if (ret == 0 && DefaultSortItem != null)
                {
                    PropertyInfo property = typeof(T).GetProperty(DefaultSortItem, BindingFlags.Public | BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.IgnoreCase, null, null, new Type[0], null);
                    if (property != null)
                    {
                        ret = CompareValue(property.GetValue(o1, null), property.GetValue(o2, null), property.PropertyType);
                    }
                }
                if (SortDirectionCore == ListSortDirection.Descending) ret = -ret;
                return ret;
            }

            private static int CompareValue(object o1, object o2, Type type)
            {
                if (o1 == null) return o2 == null ? 0 : -1;
                else if (o2 == null) return 1;
                else if (type.IsPrimitive || type.IsEnum) return Convert.ToDouble(o1).CompareTo(Convert.ToDouble(o2));
                else if (type == typeof(DateTime)) return Convert.ToDateTime(o1).CompareTo(o2);
                else return String.Compare(o1.ToString().Trim(), o2.ToString().Trim());
            }
        }

        #endregion

        #endregion

        #region 判断是否是装了WPS


        public static bool isWpsInstall()
        {
            bool isInstall = false;
            RegistryKey wpsLibrary = Registry.CurrentUser.OpenSubKey(@"Software\Kingsoft\Office\6.0\common");
            if (wpsLibrary != null)
            {
                if (wpsLibrary.GetValue("InstallRoot") != null)
                {
                    string strpath = wpsLibrary.GetValue("InstallRoot").ToString();
                    if (File.Exists(strpath + @"\office6\wps.exe"))
                    {
                        isInstall = true;
                    }
                }
            }
            return isInstall;
        }
        #endregion
        #region 获取Excel版本

        public static double JongCheckExcelVer()
        {
            Type objExcelType = Type.GetTypeFromProgID("Excel.Application");
            if (objExcelType == null)
            {
                return 0;
            }
            object objApp = Activator.CreateInstance(objExcelType);
            if (objApp == null)
            {
                return 0;
            }
            object objVer = objApp.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, objApp, null);
            double iVer = Convert.ToDouble(objVer.ToString());
            objVer = null;
            objApp = null;
            objExcelType = null;
            GC.Collect();
            return iVer;
        }

        #endregion

     
    }
}
