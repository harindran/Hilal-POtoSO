
using HillalPOtoSO;
using SAPbouiCOM.Framework;
using SBOAddonProject1;
using System;
using System.Data;

namespace Common.Common
{
    class clsAddon
    {
        public clsMenuEvent objmenuevent;
        public SAPbouiCOM.Application objapplication;
        public SAPbobsCOM.Company objcompany;
        private SAPbobsCOM.Company objAnothercompany;
        public clsRightClickEvent objrightclickevent;
        public clsGlobalMethods objglobalmethods;                  
        public string[] HWKEY = { "L1653539483", "X1211807750","K1600107675", "F0637636550" };
        #region Constructor
        public clsAddon()
        {

        }
        #endregion

        public void Intialize(string[] args)
        {
            try
            {
                Application oapplication;
                if ((args.Length < 1))
                    oapplication = new Application();
                else
                    oapplication = new Application(args[0]);
                objapplication = Application.SBO_Application;

              

              
                if (isValidLicense())
                {
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objcompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                    Create_DatabaseFields(); // UDF & UDO Creation Part    
                    Menu(); // Menu Creation Part
                    Create_Objects(); // Object Creation Part
                   //AnotherCompany();
                    objapplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(objapplication_AppEvent);
                    objapplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(objapplication_MenuEvent);
                    objapplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objapplication_ItemEvent);               
                    objapplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(ref FormDataEvent);
                    objapplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(objapplication_RightClickEvent);

                  

                    objapplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    oapplication.Run();
                }
                else
                {
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
            }          
            catch (Exception ex)
            {
                objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public bool isValidLicense()
        {
            try
            {
                if (clsModule.HANA)
                {
                    try
                    {
                        if (objapplication.Forms.ActiveForm.TypeCount > 0)
                        {
                            for (int i = 0; i <= objapplication.Forms.ActiveForm.TypeCount - 1; i++)
                                objapplication.Forms.ActiveForm.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }

                objapplication.Menus.Item("257").Activate();
                SAPbouiCOM.EditText objedit = (SAPbouiCOM.EditText)objapplication.Forms.ActiveForm.Items.Item("79").Specific;

                string CrrHWKEY = objedit.Value.ToString();
                objapplication.Forms.ActiveForm.Close();

                for (int i = 0; i <= HWKEY.Length - 1; i++)
                {
                    if (HWKEY[i] == CrrHWKEY)
                    {
                        return true;
                    }

                }

                System.Windows.Forms.MessageBox.Show("Installing Add-On failed due to License mismatch");
                return false;
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return true;
        }

        public void Create_Objects()
        {
            objmenuevent = new clsMenuEvent();
            objrightclickevent = new clsRightClickEvent();
            objglobalmethods = new clsGlobalMethods();                      
        }

        private void Create_DatabaseFields()
        {
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            var objtable = new clsTable();
            objtable.FieldCreation();
            objapplication.StatusBar.SetText(" Database Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        #region Menu Creation Details

        private void Menu()
        {
            int Menucount = 0;
            if (objapplication.Menus.Item("43545").SubMenus.Exists("POReport"))
                return;
            Menucount = objapplication.Menus.Item("43545").SubMenus.Count;
            Menucount += 1;
            CreateMenu("", Menucount, "POReport", SAPbouiCOM.BoMenuType.mt_STRING, "POReport", "43545");
      

        }

        private void CreateMenu(string ImagePath, int Position, string DisplayName, SAPbouiCOM.BoMenuType MenuType, string UniqueID, string ParentMenuID)
        {
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuPackage;
                SAPbouiCOM.MenuItem parentmenu;
                parentmenu = objapplication.Menus.Item(ParentMenuID);
                if (parentmenu.SubMenus.Exists(UniqueID.ToString()))
                    return;
                oMenuPackage = (SAPbouiCOM.MenuCreationParams)objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuPackage.Image = ImagePath;
                oMenuPackage.Position = Position;
                oMenuPackage.Type = MenuType;
                oMenuPackage.UniqueID = UniqueID;
                oMenuPackage.String = DisplayName;
                parentmenu.SubMenus.AddEx(oMenuPackage);
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
        }

        #endregion

        public bool FormExist(string FormID)
        {
            bool FormExistRet = false;
            try
            {
                FormExistRet = false;
                foreach (SAPbouiCOM.Form uid in clsModule.objaddon.objapplication.Forms)
                {
                    if (uid.TypeEx == FormID)
                    {
                        FormExistRet = true;
                        break;
                    }
                }
                if (FormExistRet)
                {
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Visible = true;
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Select();
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return FormExistRet;

        }

        private void AnotherCompany()
        {
            objAnothercompany= new SAPbobsCOM.Company();
          
          // objAnothercompany.Server = clsModule.objaddon.objglobalmethods.GetConnectionString("Server");

            objAnothercompany.Server = "GRA@Graham.tmicloud.net:30013";
            objAnothercompany.LicenseServer = "https://stefan.tmicloud.net:40000";
            objAnothercompany.SLDServer = "https://newton.tmicloud.net:40000";
            objAnothercompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;
            objAnothercompany.DbUserName = "OECDBBR";
            objAnothercompany.CompanyDB = "OEC_TEST";
            objAnothercompany.DbPassword = "India@1947";
            objAnothercompany.UserName = "TMICLOUD\\Chitra";
            objAnothercompany.Password = "N%wt$n@19%6Nqw";
            objAnothercompany.UseTrusted = false;


            int result = objAnothercompany.Connect();
            string error;
            if (result != 0)
            {
                objAnothercompany.GetLastError(out result, out error);
                Console.WriteLine("Failed to connect to SAP Business One.");
                return;
            }
        }

        #region VIRTUAL FUNCTIONS
        public virtual void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        { }

        public virtual void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        { }

        public virtual void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool BubbleEvent)
        { }

        public virtual void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        { }


        #endregion

        #region ItemEvent

        private void objapplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
              
                if (pVal.BeforeAction)
                {
                  
                    {
                        switch (pVal.EventType)
                        {
                            case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                                {
                                    SAPbouiCOM.BoEventTypes EventEnum;
                                    EventEnum = pVal.EventType;
                                    break;
                                }
                            case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                                {                                   
                                    break;
                                }
                            case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                                {
                                    break;
                                }
                            case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                                {
                                    switch (pVal.FormType)
                                    {
                                        case 142: ////Purchase order
                                            if (pVal.ItemUID == "38" && pVal.ColUID== "U_SONum")
                                            {
                                                                                             
                                                clsModule.objaddon.objglobalmethods.ActualForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                                                SAPbouiCOM.Matrix Matrix3 = (SAPbouiCOM.Matrix)clsModule.objaddon.objglobalmethods.ActualForm.Items.Item("38").Specific;
                                                int selectrow = pVal.Row;
                                                
                                                    string valueSONUM=  ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_SONum").Cells.Item(selectrow).Specific).Value.ToString();
                                                if (string.IsNullOrWhiteSpace(valueSONUM))
                                                {
                                                    Form1 form1 = new Form1();
                                                    form1.ItemCode = ((SAPbouiCOM.EditText)Matrix3.Columns.Item("1").Cells.Item(selectrow).Specific).Value.ToString();
                                                    form1.rowNo = selectrow;
                                                    form1.PValtype = pVal;
                                                    form1.Show();
                                                }
                                            }
                                            break;
                                        case 139: //sales order
                                            if (pVal.ItemUID == "38" && pVal.ColUID == "U_SONum")
                                            {

                                                clsModule.objaddon.objglobalmethods.ActualForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                                                SAPbouiCOM.Matrix Matrix3 = (SAPbouiCOM.Matrix)clsModule.objaddon.objglobalmethods.ActualForm.Items.Item("38").Specific;
                                                int selectrow = pVal.Row;
                                                string valueSONUM = ((SAPbouiCOM.EditText)Matrix3.Columns.Item("U_SONum").Cells.Item(selectrow).Specific).Value.ToString();
                                                if (string.IsNullOrWhiteSpace(valueSONUM))
                                                {
                                                    Form1 form1 = new Form1();
                                                    form1.ItemCode = ((SAPbouiCOM.EditText)Matrix3.Columns.Item("1").Cells.Item(selectrow).Specific).Value.ToString();
                                                    form1.rowNo = selectrow;
                                                    form1.PValtype = pVal;
                                                    form1.Show();
                                                }
                                            }
                                            break;
                                    }
                                    break;
                                }
                            
                        }
                    }

                }
                else
                {
                   
                    switch (pVal.EventType)
                    {
                       
                        case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                            {
                                break;
                            }
                        case SAPbouiCOM.BoEventTypes.et_CLICK:
                            {
                              
                                break;
                            }
                    }
                }

            }
            catch (Exception ex)
            {
                return;
            }
        }

        #endregion

        #region FormDataEvent

        private void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {               
            }
            catch (Exception ex)
            {
               //throw ex;
            }
        }

        #endregion

        #region MenuEvent
        private void objapplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.BeforeAction == false)
            {
                switch (pVal.MenuUID)
                {
                    case "POReport":
                        POReport Form1 = new POReport();
                        Form1.Show();
                        break;


                }
            }          
        }

        #endregion
        public void Cleartext(SAPbouiCOM.Form oForm)
        {
          

        }
        #region RightClickEvent

        private void objapplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
              

            }
            catch (Exception ex) { }

        }

        #endregion

        #region AppEvent

        private void objapplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    try
                    {
                        System.Windows.Forms.Application.Exit();
                        if (objapplication != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                        if (objcompany != null)
                        {
                            if (objcompany.Connected)
                                objcompany.Disconnect();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                        }
                        GC.Collect();

                    }
                    catch (Exception ex)
                    {
                    }
                    break;

            }
        }

        #endregion

    }


}
