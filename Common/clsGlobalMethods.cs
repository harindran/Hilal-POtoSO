using System;
using System.Data;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Globalization;
using System.ComponentModel;
using System.Configuration;
using Microsoft.VisualBasic;

namespace Common.Common
{
    class clsGlobalMethods
    {
        string strsql;
        SAPbobsCOM.Recordset objrs;
        public SAPbouiCOM.Form ActualForm;
        //Live
     public string SODB= "HCTTC_LIVE";
     public string PODB = "NORTHSTAR";

        //HILLALTEST

        //public string SODB = "TEST_HCTTC";
      // public string PODB = "TEST_NS";


       // public string SODB = "OEC_TEST";
      //public string PODB = "OEC_TEST";
        public string GetDocNum(string sUDOName, int Series)
        {
            string GetDocNumRet = "";
            string StrSQL;
            SAPbobsCOM.Recordset objRS;
            objRS = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            if (Series == 0)
            {
                StrSQL = " select  \"NextNumber\"  from NNM1 where \"ObjectCode\"='" + sUDOName + "'";
            }
            else
            {
                StrSQL = " select  \"NextNumber\"  from NNM1 where \"ObjectCode\"='" + sUDOName + "' and \"Series\" = " + Series;
            }
            objRS.DoQuery(StrSQL);
            objRS.MoveFirst();
            if (!objRS.EoF)
            {
                return Convert.ToInt32(objRS.Fields.Item(0).Value.ToString()).ToString();
            }
            else
            {
                GetDocNumRet = "1";
            }

            return GetDocNumRet;
        }

        public string Gridcolumnname(SAPbouiCOM.Grid grid, object Index)
        {
            SAPbouiCOM.GridColumn column = grid.Columns.Item(Index);

            return column.UniqueID;
        }

        public string GetNextCode_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";

                strsql = "select IFNULL(Max(CAST(\"Code\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                if (!clsModule.HANA)
                {
                    strsql = clsModule.objaddon.objglobalmethods.ChangeHANAtoSql(strsql);
                }

                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string GetNextDocNum_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocNum\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                if (!clsModule.HANA)
                {
                    strsql = clsModule.objaddon.objglobalmethods.ChangeHANAtoSql(strsql);
                }
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public List<string> SplitByLength(string str, int n)
        {
            List<string> substrings = new List<string>();
            int i = 0;
            while (i < str.Length)
            {
                int length = n;
                if (i + length < str.Length)
                {
                    while (length > 0 && !char.IsWhiteSpace(str[i + length - 1]))
                    {
                        length--;
                    }
                }
                else
                {
                    length = str.Length - i;
                }
                string insertvalue = str.Substring(i, length).Trim();
                if (insertvalue.Length <= 3) insertvalue = "  " + insertvalue;
                substrings.Add(insertvalue);
                i += length;
            }
            return substrings;
        }


        public string GetNextDocEntry_Value(string Tablename)
        {
            try
            {
                if (string.IsNullOrEmpty(Tablename.ToString()))
                    return "";
                strsql = "select IFNULL(Max(CAST(\"DocEntry\" As integer)),0)+1 from \"" + Tablename.ToString() + "\"";
                if (!clsModule.HANA)
                {
                    strsql = clsModule.objaddon.objglobalmethods.ChangeHANAtoSql(strsql);
                }
                objrs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objrs.DoQuery(strsql);
                if (objrs.RecordCount > 0)
                    return Convert.ToString(objrs.Fields.Item(0).Value);
                else
                    return "";
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" + ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";
            }
        }

        public string Convert_String_TimeHHMM(string str)
        {
            str = "0000" + Regex.Replace(str, @"[^\d]", "");
            return str.PadRight(4);
        }

        public string GetDuration_BetWeenTime(string strFrom, string strTo)
        {
            DateTime Fromtime, Totime;
            TimeSpan Duration;
            strFrom = Convert_String_TimeHHMM(strFrom);
            strTo = Convert_String_TimeHHMM(strTo);
            Totime = new DateTime(2000, 1, 1, Convert.ToInt32(strTo.PadLeft(2)), Convert.ToInt32(strTo.PadLeft(2)), 0);
            Fromtime = new DateTime(2000, 1, 1, Convert.ToInt32(strFrom.PadLeft(2)), Convert.ToInt32(strFrom.PadRight(2)), 0);
            if (Totime < Fromtime)
                Totime = new DateTime(2000, 1, 2, Convert.ToInt32(strTo.PadLeft(2)), Convert.ToInt32(strTo.PadLeft(2)), 0);
            Duration = Totime - Fromtime;
            return Duration.Hours.ToString() + "." + Duration.Minutes.ToString() + "00".PadLeft(2);
        }


        public string getSingleValue(string StrSQL)
        {
            try
            {
                SAPbobsCOM.Recordset rset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rset.DoQuery(StrSQL);
                return Convert.ToString((rset.RecordCount) > 0 ? rset.Fields.Item(0).Value.ToString() : "");
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " + ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return "";
            }
        }



        //esta es la version original
        public DataTable RsTODataTabla(ref SAPbobsCOM.Recordset _rs)
        {
            DataTable dt = new DataTable();
            for (int i = 0; i < _rs.Fields.Count; i++)
                dt.Columns.Add(_rs.Fields.Item(i).Description);
            while (!_rs.EoF)
            {
                object[] array = new object[_rs.Fields.Count];
                for (int i = 0; i < _rs.Fields.Count; i++)
                {
                    switch (_rs.Fields.Item(i).Type)
                    {
                        case SAPbobsCOM.BoFieldTypes.db_Date:
                            array[i] = Getdateformat(Convert.ToString(_rs.Fields.Item(i).Value));
                            break;
                        default:
                            array[i] = _rs.Fields.Item(i).Value;
                            break;
                    }
                }
                dt.Rows.Add(array);
                _rs.MoveNext();
            }
            return dt;
        }
        public string Getdateformat(string Inv_Doc_Date)
        {
            string date = Inv_Doc_Date;
            List<DateTime> RemoveDate = new List<DateTime>();

            RemoveDate.Add(new DateTime(1999, 12, 30));
            RemoveDate.Add(new DateTime(1899, 12, 30));

            if (string.IsNullOrEmpty(date))
                return "";

            DateTime dateTime = Convert.ToDateTime(date);
            if (RemoveDate.Contains(dateTime.Date))
            {
                return "";
            }

            string dtformate = dateTime.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            return dtformate;
        }
        public string Getdateformat(string Inv_Doc_Date, string format)
        {
            if (string.IsNullOrEmpty(Inv_Doc_Date)) return "";
            DateTime RemoveDate = new DateTime(1999, 12, 30);

            string date = Inv_Doc_Date;
            DateTime dateTime = DateTime.ParseExact(date, format, CultureInfo.InvariantCulture);
            if (dateTime.Date == RemoveDate.Date)
            {
                return "";
            }
            string dtformate = dateTime.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
            return dtformate;
        }


        public DataTable GetmultipleValue(string StrSQL)
        {
            DataTable dt = new DataTable();
            try
            {
                SAPbobsCOM.Recordset rset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rset.DoQuery(StrSQL);
                dt = RsTODataTabla(ref rset);
                return dt;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " + ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                return dt;
            }
        }

        public string MTrim(object PstrValue)
        {
            string MTrim = "";
            try
            {
                MTrim = Regex.Replace(Trim(PstrValue), "[^A-Za-z0-9]", "");
                MTrim = MTrim.ToUpper();
                return MTrim;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string MTrim(string PstrValue, string Replacement, bool Upper)
        {
            string MTrim = "";
            try
            {
                MTrim = Regex.Replace(PstrValue, "[^A-Za-z0-9]", Replacement);
                if (Upper)
                    MTrim = MTrim.ToUpper();
                return MTrim;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string MTrim(string PstrValue, string Regexp = "[^A-Za-z0-9]", string Replacement = "", bool Upper = true)
        {
            string MTrim = "";
            try
            {
                MTrim = Regex.Replace(PstrValue, Regexp, Replacement);
                if (Upper)
                    MTrim = MTrim.ToUpper();
                return MTrim;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }




        public string GetSplChr(string PstrValue)
        {
            string GetSplChr;
            try
            {
                GetSplChr = Regex.Replace(PstrValue, "[A-Za-z0-9]", "");
                return GetSplChr;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string Trim(object PstrValue)
        {
            string Trim = "";
            try
            {
                Trim = Convert.ToString(PstrValue);
                return Trim;
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool KEmpty(object Var)
        {
            bool KEmpty = false;
            try
            {
                string lstr = System.Convert.ToString(Var);

                if (string.IsNullOrEmpty(lstr))
                {
                    KEmpty = true;
                    return KEmpty;
                }

                KEmpty = lstr.Length == 0;
                return KEmpty;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public string DateToString(DateTime PstrDate, string DateFormat = "dd/MM/yyyy", DateTime? RemoveDate = null)
        {
            string DateToString;
            try
            {
                if (RemoveDate == null)
                {
                    RemoveDate = new DateTime(1900, 1, 1);
                }
                if (PstrDate.Date == RemoveDate.Value.Date)
                {
                    return "";
                }
                DateToString = PstrDate.ToString(DateFormat, CultureInfo.InvariantCulture);
                return DateToString;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DateTime GetDate(string PstrDate, string DateFormat = "dd/MM/yyyy")
        {
            DateTime GetDate;
            try
            {
                var LstrDate = PstrDate;
                var LintRetYear = DateTime.Now.Year;
                var LintRetMonth = DateTime.Now.Month;
                var LintRetDay = DateTime.Now.Day;
                string LSplcFormat = GetSplChr(DateFormat);
                string LSplcDate = GetSplChr(LstrDate);

                LstrDate = MTrim(LstrDate, "[^0-9" + LSplcDate + "]");

                if (LSplcFormat.Length > 0)
                {
                    for (var lint = 0; lint <= LSplcFormat.Length - 1; lint++)
                        LstrDate = MTrim(LstrDate, Trim(LSplcFormat[lint]), false);

                    if (!KEmpty(LstrDate))
                    {
                        string[] splitvalue = Regex.Split(LstrDate, Trim(LSplcFormat[0]));
                        for (var lint = 0; lint <= splitvalue.Length - 1; lint++)
                        {
                            if (KEmpty(splitvalue[lint]))
                                continue;

                            switch (lint)
                            {
                                case 0:
                                    {
                                        LintRetDay = Convert.ToInt32(splitvalue[lint]);
                                        break;
                                    }
                                case 1:
                                    {
                                        LintRetMonth = Convert.ToInt32(splitvalue[lint]);
                                        break;
                                    }
                                case 2:
                                    {
                                        LintRetYear = Convert.ToInt32(Strings.Mid(splitvalue[lint], 1, 4));
                                        break;
                                    }
                            }
                        }
                    }
                }
                LstrDate = DateToString(DateAndTime.DateSerial(LintRetYear, LintRetMonth, LintRetDay), DateFormat);
                if (DateTime.TryParseExact(LstrDate, DateFormat, CultureInfo.CurrentCulture, DateTimeStyles.None, out GetDate))
                    GetDate = DateAndTime.DateSerial(LintRetYear, LintRetMonth, LintRetDay);

                return GetDate;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public string ChangeHANAtoSql(string StrSQL)
        {
            string sql = StrSQL;
            try
            {
                sql = sql.Replace("IFNULL", "ISNULL");
                sql = sql.Replace("ifnull", "ISNULL");

                return sql;
            }
            catch (Exception ex)
            {
                return sql;
            }
        }

        public void LoadSeries(SAPbouiCOM.Form objform, SAPbouiCOM.DBDataSource DBSource, string ObjectType)
        {
            try
            {
                SAPbouiCOM.ComboBox ComboBox0;
                ComboBox0 = (SAPbouiCOM.ComboBox)objform.Items.Item("Series").Specific;
                ComboBox0.ValidValues.LoadSeries(ObjectType, SAPbouiCOM.BoSeriesMode.sf_Add);
                ComboBox0.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                DBSource.SetValue("DocNum", 0, clsModule.objaddon.objglobalmethods.GetDocNum(ObjectType, Convert.ToInt32(ComboBox0.Selected.Value)));
            }
            catch (Exception ex)
            {

            }
        }


        public void WriteErrorLog(string Str)
        {
            try
            {
                string Foldername;
                Foldername = @"Log";
                if (Directory.Exists(Foldername))
                {
                }
                else
                {
                    Directory.CreateDirectory(Foldername);
                }

                FileStream fs;
                string chatlog = Foldername + @"\Log_" + DateTime.Now.ToString("ddMMyy") + ".txt";
                if (File.Exists(chatlog))
                {
                }
                else
                {
                    fs = new FileStream(chatlog, FileMode.Create, FileAccess.Write);
                    fs.Close();
                }
                string sdate;
                sdate = Convert.ToString(DateTime.Now);
                if (File.Exists(chatlog) == true)
                {
                    var objWriter = new StreamWriter(chatlog, true);
                    objWriter.WriteLine(sdate + " : " + Str);
                    objWriter.Close();
                }
                else
                {
                    var objWriter = new StreamWriter(chatlog, false);
                }
            }
            catch (Exception)
            {


            }
        }


        public void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
        {
            try
            {
                if (omatrix.VisualRowCount == 0)
                    return;
                if (string.IsNullOrEmpty(Columname_check.ToString()))
                    return;
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific).String == "")
                {
                    omatrix.DeleteRow(omatrix.VisualRowCount);
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void SetAutomanagedattribute_Editable(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
        }

        public void SetAutomanagedattribute_Visible(SAPbouiCOM.Form oform, string fieldname, bool add, bool find, bool update)
        {

            if (add == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Add), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (find == true)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Find), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }

            if (update)
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_True);
            }
            else
            {
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, Convert.ToInt32(SAPbouiCOM.BoAutoFormMode.afm_Ok), SAPbouiCOM.BoModeVisualBehavior.mvb_False);
            }
        }

        public void Matrix_Addrow(SAPbouiCOM.Matrix omatrix, string colname = "", string rowno_name = "", bool Error_Needed = false)
        {
            try
            {
                bool addrow = false;

                if (omatrix.VisualRowCount == 0)
                {
                    addrow = true;
                    goto addrow;
                }
                if (string.IsNullOrEmpty(colname))
                {
                    addrow = true;
                    goto addrow;
                }
                if (((SAPbouiCOM.EditText)omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific).String != "")
                {
                    addrow = true;
                    goto addrow;
                }

            addrow:
                ;

                if (addrow == true)
                {
                    omatrix.AddRow(1);
                    omatrix.ClearRowData(omatrix.VisualRowCount);
                    if (!string.IsNullOrEmpty(rowno_name))
                        ((SAPbouiCOM.EditText)omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific).String = Convert.ToString(omatrix.VisualRowCount);
                }
                else if (Error_Needed == true)
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            catch (Exception ex)
            {

            }
        }

        public int GetColumnindex(SAPbouiCOM.Matrix matrix, string uniquecode)
        {
            string columnName = uniquecode;
            for (int i = 0; i <= matrix.Columns.Count - 1; i++)
            {
                if (matrix.Columns.Item(i).UniqueID == columnName)
                {
                    return i;
                }
            }
            return -1;
        }

        public int GetColumnindex(SAPbouiCOM.Grid matrix, string uniquecode)
        {
            string columnName = uniquecode;
            for (int i = 0; i <= matrix.Columns.Count - 1; i++)
            {
                if (matrix.Columns.Item(i).UniqueID == columnName)
                {
                    return i;
                }
            }
            return -1;
        }
        public void columnadd(SAPbouiCOM.Grid grid, string colname, SAPbouiCOM.BoFieldsType type = SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, SAPbouiCOM.BoGridColumnType coltype = SAPbouiCOM.BoGridColumnType.gct_EditText, Dictionary<string, string> keyValuePairs = null)
        {
            try
            {
                SAPbouiCOM.ComboBoxColumn ocomboColumn;

                grid.DataTable.Columns.Add(colname, type);
                grid.Columns.Item(colname).Type = coltype;
                if (coltype == SAPbouiCOM.BoGridColumnType.gct_ComboBox)
                {
                    ocomboColumn = (SAPbouiCOM.ComboBoxColumn)grid.Columns.Item(colname);
                    ocomboColumn.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Value;
                    foreach (var item in keyValuePairs)
                    {
                        ocomboColumn.ValidValues.Add(item.Value, item.Key);
                    }
                }
                keyValuePairs.Clear();
            }
            catch (Exception ex)
            {

                return;
            }

        }
    }
}
