using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Common.Common;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace HillalPOtoSO
{
    [FormAttribute("POReport", "POReport.b1f")]
    public class POReport : UserFormBase
    {
        private List<string> lastColumns = new List<string>();
        Dictionary<string, string> keyValuePairs = new Dictionary<string, string>();
        private int startcol;
        public static SAPbouiCOM.Form objform;
        private string curcol;
        SAPbouiCOM.ProgressBar oProgBar;

        public POReport()
        {
        }
        Recordset recordset;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_4").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.KeyDownAfter += new SAPbouiCOM._IGridEvents_KeyDownAfterEventHandler(this.Grid0_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_5").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("Lblto").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("ToDt").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_9").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_10").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("btnsav").Specific));
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("TxtFind").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("Frmtxt").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("POReport", pVal.FormTypeCount);

            }
            catch (Exception ex)
            {

                throw ex;
            }

        }

        private void OnCustomInitialize()
        {
            Recordset recordset = (Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
           //our db
            //string query2 = "Select \"FldValue\" ,\"Descr\" FROM " + clsModule.objaddon.objglobalmethods.SODB + ".UFD1 Where \"TableID\" = '@SMPR_OHEM' and \"FieldID\" = '43'";


       string query2 = "Select \"FldValue\" ,\"Descr\" FROM " + clsModule.objaddon.objglobalmethods.SODB + ".UFD1 Where \"TableID\" = 'ORDR' and \"FieldID\" = '10'";
            query2 += " union all ";
            query2 += "Select 'Purchase Order','Purchase Order'  FROM dummy ";


            DataTable dt = new DataTable();
            dt = clsModule.objaddon.objglobalmethods.GetmultipleValue(query2);

            foreach (DataRow row in dt.Rows)
            {
                string value = row["FldValue"].ToString();
                string description = row["Descr"].ToString();

                ComboBox0.ValidValues.Add(value, description);
            }

            lastColumns.Add("INTERNET NUMBER");
            lastColumns.Add("PI NUMBER");
            lastColumns.Add("ORDER NUMBER");



        }
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.Button Button0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            
            Loaddata();


            return;
        }

        private void Loaddata()
        {
            try
            {               
                Recordset recordset = (Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                if (string.IsNullOrEmpty(EditText1.Value))
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("From Date is Missing....!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                if (string.IsNullOrEmpty(EditText6.Value))
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("To Date is Missing....!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
                string query2 = "call GetInvoice_Details('" + EditText1.Value + "','" + EditText6.Value + "','" + ComboBox0.Selected.Value + "')";

                recordset.DoQuery(query2);
                this.Grid0.DataTable.Clear();
                DataTable dt = new DataTable();
                dt = clsModule.objaddon.objglobalmethods.GetmultipleValue(query2);
                objform.Freeze(true);
                foreach (DataColumn item in dt.Columns)
                {
                   switch(item.ColumnName)
                    {
                        case "OFF_ON":
                            keyValuePairs.Add("", "");
                            keyValuePairs.Add("OffLine", "Offline");
                            keyValuePairs.Add("Online", "Online");                           
                            clsModule.objaddon.objglobalmethods.columnadd(this.Grid0, item.ColumnName, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, SAPbouiCOM.BoGridColumnType.gct_ComboBox, keyValuePairs);
                            break;
                        case "PODATE":
                            clsModule.objaddon.objglobalmethods.columnadd(this.Grid0, item.ColumnName, SAPbouiCOM.BoFieldsType.ft_Date);
                            break;
                        default:
                          clsModule.objaddon.objglobalmethods.columnadd(this.Grid0,item.ColumnName, SAPbouiCOM.BoFieldsType.ft_AlphaNumeric);
                            break;
                    }
                  
                }


                this.Grid0.DataTable.Rows.Add(dt.Rows.Count);

                if (dt.Rows.Count > 0)
                {
                    oProgBar = clsModule.objaddon.objapplication.StatusBar.CreateProgressBar("Loading Please Wait", dt.Rows.Count, true);

                    foreach (DataRow row in dt.Rows)
                    {
                        foreach (DataColumn column in dt.Columns)
                        {
                            switch (column.ColumnName)
                            {
                                case "PODATE":
                                    if (!string.IsNullOrEmpty((row[column].ToString().Trim())))
                                    {

                                        this.Grid0.DataTable.SetValue(dt.Columns.IndexOf(column), dt.Rows.IndexOf(row), clsModule.objaddon.objglobalmethods.GetDate(row[column].ToString().Trim()));
                                    }
                                    break;
                                default:
                                    this.Grid0.DataTable.SetValue(dt.Columns.IndexOf(column), dt.Rows.IndexOf(row), row[column]);
                                    break;
                            }


                          
                        }
                        oProgBar.Value += 1;
                    }

                }
                // this.Grid0.DataTable.ExecuteQuery(query2);
                this.Grid0.RowHeaders.TitleObject.Caption="#";
                this.Grid0.AutoResizeColumns();
                roweditable();
                editable();
                Colsetting();
                objform.Freeze(false);
                oProgBar.Stop();
                return;

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.StackTrace.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                
                return;
            }


        }
      
        private void roweditable()
        {
            

            for (int i = 0; i < this.Grid0.Rows.Count; i++)
            {
                this.Grid0.CommonSetting.SetRowEditable(i + 1,true);
                

                if (string.IsNullOrEmpty(this.Grid0.DataTable.GetValue("PPINum", i).ToString()) || string.IsNullOrEmpty(this.Grid0.DataTable.GetValue("PODNum", i).ToString()))
                {
                    this.Grid0.CommonSetting.SetRowFontColor(i + 1, 255);
                }
                else
                {
                    this.Grid0.CommonSetting.SetRowFontColor(i + 1, 6863650);
                }

                switch (this.Grid0.DataTable.GetValue("Status", i).ToString())
                {
                    case "GRPO CREATED":
                        this.Grid0.CommonSetting.SetRowEditable(i + 1, false);
                        break;
                }
                this.Grid0.RowHeaders.SetText(i, (i + 1).ToString());
                

            }

        }

        private void editable()
        {
            int lastnoneditcol = 6;
            for (int i = 0; i < this.Grid0.Columns.Count - lastnoneditcol; i++)
            {
                SAPbouiCOM.GridColumn column = this.Grid0.Columns.Item(i);
                column.Editable = false;
                if (i < 5)
                {
                    column.Visible = false;
                }

              //  objform.EnableMenu("4870", true);


            }

            startcol = this.Grid0.Columns.Count - 3;
        }
        private void Colsetting()
        {
            for (int i = 0; i < this.Grid0.Columns.Count; i++)
            {
                this.Grid0.Columns.Item(i).TitleObject.Sortable = true;               

            }
            this.Grid0.Columns.Item(0).Editable = false;
        }


        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.Button Button1;

        private void Button1_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {


                string query;



                for (int i = 0; i < this.Grid0.Rows.Count; i++)
                {
                    clear();
                    string columnname="";
                    if (this.Grid0.DataTable.GetValue("Edit", i).ToString() == "1")
                    {
                        int j = 0;
                        switch(this.Grid0.DataTable.GetValue("Sales", i).ToString())
                        {
                            case "1":
                                query = "update "+ clsModule.objaddon.objglobalmethods.PODB + ".POR1 set ";
                                break;
                            default:
                                query = "update " + clsModule.objaddon.objglobalmethods.SODB + ".RDR1 set ";
                                break;
                        }                        
                        Recordset recordset = (Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                        foreach (var item in lastColumns)
                        {

                            string columnvalue = this.Grid0.DataTable.GetValue(clsModule.objaddon.objglobalmethods.GetColumnindex(this.Grid0,item.ToString()), i).ToString();
                            switch (item)
                            {
                                case "INTERNET NUMBER":
                                    columnname = "U_INNum";
                                    break;
                                case "PI NUMBER":
                                    columnname = "U_PINum";
                                    break;
                                case "ORDER NUMBER":
                                    columnname = "U_ODNum";
                                    break;
                            }
                            query += "\"" + columnname + "\"" + " = '" + columnvalue + "'";
                            if (lastColumns.Count != j + 1)
                            {
                                query += ", ";
                            }
                            j++;
                        }


                        query += " where \"ItemCode\" ='" + this.Grid0.DataTable.GetValue("ItemCode", i).ToString() + "'" +
                                 " and \"DocEntry\" ='" + this.Grid0.DataTable.GetValue("DocEntry", i).ToString() + "' "+
                                 " and \"LineNum\" ='" + this.Grid0.DataTable.GetValue("LineNum", i).ToString() + "' ";
                        recordset.DoQuery(query);
                        SaveLogs(this.Grid0, i);
                        clear();
                    }
                }
                clsModule.objaddon.objapplication.StatusBar.SetText("Update Successfully...!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                Loaddata();

            }

            catch (Exception ex)
            {

                clsModule.objaddon.objapplication.StatusBar.SetText(ex.ToString(), SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return;
            }
        }
        private void clear()
        {

        }


        private void SaveLogs(SAPbouiCOM.Grid grid, int row)
        {
            try
            {
                UserTable userTable;
                userTable = clsModule.objaddon.objcompany.UserTables.Item("POTOSO");

                for (int i = 0; i <= userTable.UserFields.Fields.Count - 1; i++)
                {
                    userTable.UserFields.Fields.Item(i).Value = userTable.UserFields.Fields.Item(i).DefaultValue;
                }

                userTable.UserFields.Fields.Item("U_ItemCode").Value = this.Grid0.DataTable.GetValue("ItemCode", row).ToString();
                userTable.UserFields.Fields.Item("U_SODocEntry").Value = this.Grid0.DataTable.GetValue("DocEntry", row).ToString();

                if ((!string.IsNullOrEmpty(this.Grid0.DataTable.GetValue("INTERNET NUMBER", row).ToString())) && this.Grid0.DataTable.GetValue("PINNum", row).ToString() != this.Grid0.DataTable.GetValue("INTERNET NUMBER", row).ToString())
                {
                    userTable.UserFields.Fields.Item("U_INNum").Value = this.Grid0.DataTable.GetValue("INTERNET NUMBER", row).ToString();
                    userTable.UserFields.Fields.Item("U_INNumUDt").Value = DateTime.Now;
                }
                if ((!string.IsNullOrEmpty(this.Grid0.DataTable.GetValue("PI NUMBER", row).ToString())) && this.Grid0.DataTable.GetValue("PPINum", row).ToString() != this.Grid0.DataTable.GetValue("PI NUMBER", row).ToString())
                {
                    userTable.UserFields.Fields.Item("U_PINum").Value = this.Grid0.DataTable.GetValue("PI NUMBER", row).ToString();
                    userTable.UserFields.Fields.Item("U_PINumUDt").Value = DateTime.Now;
                }
                if ((!string.IsNullOrEmpty(this.Grid0.DataTable.GetValue("ORDER NUMBER", row).ToString())) && this.Grid0.DataTable.GetValue("PODNum", row).ToString() != this.Grid0.DataTable.GetValue("ORDER NUMBER", row).ToString())
                {
                    userTable.UserFields.Fields.Item("U_ODNum").Value = this.Grid0.DataTable.GetValue("ORDER NUMBER", row).ToString();
                    userTable.UserFields.Fields.Item("U_ODNumUDt").Value = DateTime.Now;
                }
                if (!string.IsNullOrEmpty(Convert.ToString(this.Grid0.DataTable.GetValue("PODATE", row))))
                {
                    userTable.UserFields.Fields.Item("U_PODATE").Value = Convert.ToDateTime(this.Grid0.DataTable.GetValue("PODATE", row).ToString().Trim());                    
                }
                
                userTable.UserFields.Fields.Item("U_Serial").Value = this.Grid0.DataTable.GetValue("SERIAL", row).ToString();
                userTable.UserFields.Fields.Item("U_Off_OnStatus").Value =this.Grid0.DataTable.GetValue("OFF_ON", row).ToString();


                userTable.UserFields.Fields.Item("U_Createdt").Value = DateTime.Now;
                userTable.UserFields.Fields.Item("U_Createtime").Value = DateTime.Now;
                userTable.UserFields.Fields.Item("U_type").Value = this.Grid0.DataTable.GetValue("Sales", row).ToString();
                userTable.UserFields.Fields.Item("U_LineNum").Value = this.Grid0.DataTable.GetValue("LineNum", row).ToString();
                userTable.Add();
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objglobalmethods.WriteErrorLog(ex.ToString());
                clsModule.objaddon.objapplication.StatusBar.SetText("SaveLogs: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Grid0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.CharPressed == 9)
            {
                this.Grid0.DataTable.SetValue("Edit", pVal.Row, "1");
            }
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            string strSQL;
            int find;
            if (string.IsNullOrEmpty(curcol)) return;
            this.Grid0.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
          
                for (find = 0; find <= this.Grid0.Rows.Count - 1; find++)
                {
                    strSQL = this.Grid0.DataTable.GetValue(curcol, this.Grid0.GetDataTableRowIndex(find)).ToString().ToUpper();
                    if (strSQL.Contains(EditText0.Value.ToString().ToUpper()))
                    {
                        this.Grid0.Rows.SelectedRows.Clear();
                        this.Grid0.Rows.SelectedRows.Add(find);
                        break;
                    }
                }
          
        }
        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            curcol = pVal.ColUID;

        }

        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
    }
}
