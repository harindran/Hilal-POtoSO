using System;
using System.Collections.Generic;
using System.Xml;
using Common.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace SBOAddonProject1
{
    [FormAttribute("Formchk", "Choose.b1f")]
    class Form1 : UserFormBase
    {

        public string ItemCode;
        public int rowNo;
        private SAPbouiCOM.Form Form;
        private int FormCount = 0;
        public SAPbouiCOM.ItemEvent PValtype;
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid3 = ((SAPbouiCOM.Grid)(this.GetItem("Item_8").Specific));
            this.Grid3.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid3_DoubleClickAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
            this.VisibleAfter += new VisibleAfterHandler(this.Form1_VisibleAfter);



        }

        private void Form1_VisibleAfter(SBOItemEventArg pVal)
        {
            Loaddata();
        }

        private void editable()
        {
            for (int i = 0; i < this.Grid3.Columns.Count; i++)
            {
                SAPbouiCOM.GridColumn column = this.Grid3.Columns.Item(i);
                column.Editable = false;
            }
        }
        private void Loaddata()
        {
            if (string.IsNullOrWhiteSpace(ItemCode)) { return; }
            //ItemCode = "14-140-HB002-00112";
            string query2 = "";
            Recordset recordset = (Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            switch (PValtype.FormType)
            {
                case 142:
                    query2 = " SELECT  o.\"DocEntry\",o.\"DocNum\" ,o.\"DocDate\",o.\"CardCode\",o.\"CardName\" ,sum(o.\"DocTotal\") AS \"DocTotal\", " +
                                   " CAST (r.\"U_INNum\" AS VARCHAR(50)) AS \"Internet Number\",CAST (r.\"U_PONum\" AS VARCHAR(50)) AS \"Purchase Number\", " +
                                   " CAST(r.\"U_ODNum\" AS VARCHAR(50)) AS \"Order Number\",CAST (r.\"U_PINum\" AS VARCHAR(50)) AS \"PI Number\" " +
                                   " FROM " + clsModule.objaddon.objglobalmethods.SODB + ".ORDR o " +
                                   " LEFT JOIN " + clsModule.objaddon.objglobalmethods.SODB + ".RDR1 r  ON o.\"DocEntry\" = r.\"DocEntry\" " +
                                   " LEFT JOIN " + clsModule.objaddon.objglobalmethods.SODB + ".OITM o2 ON o2.\"ItemCode\" = r.\"ItemCode\" " +
                                   " WHERE o2.\"ItemCode\" = '" + ItemCode + "' AND o.\"DocStatus\" = 'O' " +
                                   " AND o.\"DocEntry\" NOT IN ( SELECT   CAST(\"U_SODocEntry\" AS Integer) FROM POR1 p WHERE CAST(\"U_SODocEntry\" AS VARCHAR(50)) <>'' and \"ItemCode\"='" + ItemCode +"')" +
                                   " GROUP BY o.\"DocEntry\",o.\"DocNum\" ,o.\"DocDate\",o.\"CardCode\",o.\"CardName\", " +
                                   " CAST(r.\"U_INNum\" AS VARCHAR(50)) ,CAST (r.\"U_PONum\" AS VARCHAR(50)) , " +
                                   " CAST(r.\"U_ODNum\" AS VARCHAR(50)),CAST (r.\"U_PINum\" AS VARCHAR(50)) ";
                    break;

                case 139:
                    query2 = " SELECT  o.\"DocEntry\",o.\"DocNum\" ,o.\"DocDate\",o.\"CardCode\",o.\"CardName\" ,sum(o.\"DocTotal\") AS \"DocTotal\", " +
                                   " CAST (r.\"U_INNum\" AS VARCHAR(50)) AS \"Internet Number\",CAST (r.\"U_PONum\" AS VARCHAR(50)) AS \"Purchase Number\", " +
                                   " CAST(r.\"U_ODNum\" AS VARCHAR(50)) AS \"Order Number\",CAST (r.\"U_PINum\" AS VARCHAR(50)) AS \"PI Number\" " +
                                   " FROM " + clsModule.objaddon.objglobalmethods.PODB + ".OPOR o " +
                                   " LEFT JOIN " + clsModule.objaddon.objglobalmethods.PODB + ".POR1 r  ON o.\"DocEntry\" = r.\"DocEntry\" " +
                                   " LEFT JOIN " + clsModule.objaddon.objglobalmethods.PODB + ".OITM o2 ON o2.\"ItemCode\" = r.\"ItemCode\" " +
                                   " WHERE o2.\"ItemCode\" = '" + ItemCode + "' AND o.\"DocStatus\" = 'O' " +
                                   " AND o.\"DocEntry\" NOT IN ( SELECT   CAST(\"U_SODocEntry\" AS Integer) FROM RDR1 p WHERE CAST(\"U_SODocEntry\" AS VARCHAR(50)) <>'' and \"ItemCode\"='" + ItemCode + "')" +
                                   " GROUP BY o.\"DocEntry\",o.\"DocNum\" ,o.\"DocDate\",o.\"CardCode\",o.\"CardName\", " +
                                   " CAST(r.\"U_INNum\" AS VARCHAR(50)) ,CAST (r.\"U_PONum\" AS VARCHAR(50)) , " +
                                   " CAST(r.\"U_ODNum\" AS VARCHAR(50)),CAST (r.\"U_PINum\" AS VARCHAR(50)) ";
                    break;
            }
            recordset.DoQuery(query2);
            this.Grid3.DataTable.ExecuteQuery(query2);
            this.Grid3.AutoResizeColumns();
            editable();



            return;
            int i = 0;
            if (recordset.RecordCount > 0)
            {
                this.Grid3.DataTable.Rows.Add(recordset.RecordCount);

                while (!recordset.EoF)
                {
                    string d2d = recordset.Fields.Item("DocEntry").Value.ToString();
                    this.Grid3.DataTable.SetValue("DocNum", i, recordset.Fields.Item("DocNum").Value.ToString());
                    this.Grid3.DataTable.SetValue("DocEntry", i, recordset.Fields.Item("DocEntry").Value.ToString());
                    this.Grid3.DataTable.SetValue("DocDate", i, recordset.Fields.Item("DocDate").Value.ToString());
                    this.Grid3.DataTable.SetValue("CardCode", i, recordset.Fields.Item("CardCode").Value.ToString());
                    this.Grid3.DataTable.SetValue("DocTotal", i, recordset.Fields.Item("DocTotal").Value.ToString());
                    recordset.MoveNext();
                    i++;

                }
                this.Grid3.AutoResizeColumns();
                editable();

            }

        }
        private void OnCustomInitialize()
        {
            Form = clsModule.objaddon.objapplication.Forms.GetForm("Formchk", FormCount);
        }

        private void Form_LoadAfter(SBOItemEventArg pVal)
        {

        }

        private Grid Grid3;

        private void Grid3_DoubleClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            SAPbouiCOM.Matrix Matrix3 = (SAPbouiCOM.Matrix)clsModule.objaddon.objglobalmethods.ActualForm.Items.Item("38").Specific;
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_SONum").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("DocNum", pVal.Row).ToString();
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_SODocEntry").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("DocEntry", pVal.Row).ToString();
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_PINum").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("PI Number", pVal.Row).ToString();
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_INNum").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("Internet Number", pVal.Row).ToString();
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_PONum").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("Purchase Number", pVal.Row).ToString();
            ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_ODNum").Cells.Item(rowNo).Specific).Value = this.Grid3.DataTable.GetValue("Order Number", pVal.Row).ToString();
            Form.Close();

        }
    }
}