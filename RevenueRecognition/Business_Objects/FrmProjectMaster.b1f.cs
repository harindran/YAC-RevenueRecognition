using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using RevenueRecognition.Common;
using System.Drawing;

namespace RevenueRecognition.Business_Objects
{
    [FormAttribute("REVPRJMSTR", "Business_Objects/FrmProjectMaster.b1f")]
    class FrmProjectMaster : UserFormBase
    {
        public FrmProjectMaster()
        {
        }
        public static SAPbouiCOM.Form objform, TempForm;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsContent, odbdsAttachment,odbdsBoqItem, odbdsBoqLabour;
        private string strSQL, strQuery;
        private SAPbobsCOM.Recordset objRs, Recordset;
        SAPbouiCOM.ISBOChooseFromListEventArg pCFL;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lprjcode").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tprjcode").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.EditText0.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText0_ChooseFromListBefore);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lprjname").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tprjname").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("ldate").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tdate").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkprjcod").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("cprotype").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("lprotype").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("ccomres").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("lcomres").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cstat").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("lstat").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("liprojval").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("tiprojval").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("leprojval").Specific));
            this.EditText12 = ((SAPbouiCOM.EditText)(this.GetItem("teprojval").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("lprerev").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("tprerev").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("lprecost").Specific));
            this.EditText14 = ((SAPbouiCOM.EditText)(this.GetItem("tprecost").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fldrcont").Specific));
            this.Folder0.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder0_PressedAfter);
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxcont").Specific));
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Matrix0.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix0_ClickBefore);
            this.Matrix0.PressedBefore += new SAPbouiCOM._IMatrixEvents_PressedBeforeEventHandler(this.Matrix0_PressedBefore);
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Matrix0.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix0_ChooseFromListAfter);
            this.Matrix0.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix0_ChooseFromListBefore);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxattach").Specific));
            this.Matrix1.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix1_ClickAfter);
            this.Matrix1.DoubleClickAfter += new SAPbouiCOM._IMatrixEvents_DoubleClickAfterEventHandler(this.Matrix1_DoubleClickAfter);
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("btnbrowse").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("btndisp").Specific));
            this.Button3.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button3_ClickAfter);
            this.Button4 = ((SAPbouiCOM.Button)(this.GetItem("btndel").Specific));
            this.Button4.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button4_ClickAfter);
            this.Folder3 = ((SAPbouiCOM.Folder)(this.GetItem("fldrboq").Specific));
            this.Folder3.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder3_PressedAfter);
            this.Folder6 = ((SAPbouiCOM.Folder)(this.GetItem("fboqitem").Specific));
            this.Folder6.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder6_PressedAfter);
            this.Folder7 = ((SAPbouiCOM.Folder)(this.GetItem("fboqlab").Specific));
            this.Folder7.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder7_PressedAfter);
            this.Matrix2 = ((SAPbouiCOM.Matrix)(this.GetItem("mboqitem").Specific));
            this.Matrix2.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix2_KeyDownBefore);
            this.Matrix2.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix2_ClickBefore);
            this.Matrix2.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix2_LinkPressedBefore);
            this.Matrix2.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix2_ValidateAfter);
            this.Matrix2.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix2_ChooseFromListAfter);
            this.Matrix2.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix2_ChooseFromListBefore);
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("mboqlab").Specific));
            this.Matrix3.ComboSelectAfter += new SAPbouiCOM._IMatrixEvents_ComboSelectAfterEventHandler(this.Matrix3_ComboSelectAfter);
            this.Matrix3.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix3_ClickBefore);
            this.Matrix3.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix3_KeyDownBefore);
            this.Matrix3.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix3_LinkPressedBefore);
            this.Matrix3.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix3_ValidateAfter);
            this.Matrix3.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix3_ChooseFromListAfter);
            this.Matrix3.ChooseFromListBefore += new SAPbouiCOM._IMatrixEvents_ChooseFromListBeforeEventHandler(this.Matrix3_ChooseFromListBefore);
            this.Folder8 = ((SAPbouiCOM.Folder)(this.GetItem("fldr1").Specific));
            this.Folder9 = ((SAPbouiCOM.Folder)(this.GetItem("fldr2").Specific));
            this.Folder11 = ((SAPbouiCOM.Folder)(this.GetItem("fldrattach").Specific));
            this.Folder11.PressedAfter += new SAPbouiCOM._IFolderEvents_PressedAfterEventHandler(this.Folder11_PressedAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lctdesc").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tctdesc").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("txtentry").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("cpoc").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lpoc").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private void OnCustomInitialize()
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                objform.Freeze(true);
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR");
                odbdsContent = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR1");//Content
                odbdsAttachment = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR2"); //Attachments
                odbdsBoqItem = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR3"); //BOQ Item
                odbdsBoqLabour = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR4"); //BOQ Labour
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tprjname", false, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tiprojval", false, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "teprojval", false, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "txtentry", false, true, false);
                strSQL = "Select * from \"@AT_PROJTYPE\"";
                clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("cprotype").Specific), strSQL, new[] { "-,-" });
                strSQL = "Select * from \"@AT_TYPE\"";
                clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("ccomres").Specific), strSQL, new[] { "-,-" });
                Manage_Fields();//Hiding Matrix Columns
                CostCenter(); //CostCenter
                Matrix0.Item.Enabled = false; Matrix1.Item.Enabled = false; Matrix2.Item.Enabled = false; Matrix3.Item.Enabled = false;
                //if (clsModule.objaddon.HANA == true)
                //    strSQL = "Select \"USERID\",\"TPLId\" from OUSR Where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'";
                //else
                //    strSQL = "Select USERID,TPLId from OUSR Where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'";
                //objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //objRs.DoQuery(strSQL);

                ////clsModule.objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform, "-REVPRJMSTR", Convert.ToInt32(objRs.Fields.Item("USERID").Value), Convert.ToInt32(objRs.Fields.Item("TPLId").Value)); //REVPRJMSTR

                ////clsModule.objaddon.objglobalmethods.Create_User_Queries("Select * from \"@AT_PROJMSTR\"", "Project Master Query");
                //objform.Settings.Enabled = true;
                //objform.Settings.MatrixUID = "mtxcont";
                //Folder11.Item.Left = Folder3.Item.Left + Folder3.Item.Width + 5;
                //objform.ActiveItem = "tprjcode";
                //Folder0.Item.Click();

                ////********************** Dynamic UDF Creation in Line Level of Matrix **************************************
                //if (clsModule.objaddon.HANA == true)
                //{
                //    strSQL = "Select \"USERID\",\"TPLId\" from OUSR Where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'";
                //    strQuery = "Select '@' || \"SonName\" \"TableName\" from UDO1 Where \"Code\" = '" + objform.BusinessObject.Type + "'";
                //}
                //else
                //{
                //    strSQL = "Select USERID,TPLId from OUSR Where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'";
                //    strQuery = "Select '@' + SonName TableName from UDO1 Where Code = '" + objform.BusinessObject.Type + "'";
                //}               

                //Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //Recordset.DoQuery(strQuery);

                //Dictionary<string, string> Table_Matrix = new Dictionary<string, string>(); // Adding Matrix UID of each Line Table
                //List<string> MatrixIDs = new List<string>();
                //string MatID;
                //Table_Matrix.Add("mtxcont", "@AT_PROJMSTR1");
                //Table_Matrix.Add("mtxattach", "@AT_PROJMSTR2");
                //Table_Matrix.Add("mboqitem", "@AT_PROJMSTR3");
                //Table_Matrix.Add("mboqlab", "@AT_PROJMSTR4");

                //if (Recordset.RecordCount > 0)
                //{
                //    for (int i = 0; i < Recordset.RecordCount; i++)
                //    {
                //        if (!Table_Matrix.ContainsValue(Convert.ToString(Recordset.Fields.Item("TableName").Value))) continue;

                //        foreach (var pair in Table_Matrix)
                //        {
                //            if (pair.Value == Convert.ToString(Recordset.Fields.Item("TableName").Value))
                //            {
                //                MatrixIDs.Add(pair.Key);
                //            }
                //        }
                //        MatID = String.Format("'{0}'", String.Join("','", MatrixIDs));
                //        //strSQL = Table_Matrix[Convert.ToString(Recordset.Fields.Item("TableName").Value)];

                //        clsModule.objaddon.objglobalmethods.Create_Dynamic_LineTable_UDF(objform, Convert.ToString(Recordset.Fields.Item("TableName").Value), objform.TypeEx, String.Format("'{0}'", String.Join("','", MatrixIDs)));
                //        Recordset.MoveNext();
                //    }
                //}
                ////********************** Dynamic UDF END **************************************

                //objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //objRs.DoQuery(strSQL);
                //clsModule.objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform, "-" + objform.TypeEx, Convert.ToInt32(objRs.Fields.Item("USERID").Value), Convert.ToInt32(objRs.Fields.Item("TPLId").Value)); //REVPRJMSTR


                objform.Freeze(false);

            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #region Fields
        
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText12;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.EditText EditText14;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button4;
        private SAPbouiCOM.Folder Folder3;
        private SAPbouiCOM.Folder Folder6;
        private SAPbouiCOM.Folder Folder7;
        private SAPbouiCOM.Matrix Matrix2;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.Folder Folder8;
        private SAPbouiCOM.Folder Folder9;
        private SAPbouiCOM.Folder Folder11;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.StaticText StaticText3;

        #endregion

        #region Header Items & Form Events

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("REVPRJMSTR", pVal.FormTypeCount);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Exception: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                bool originFlag = false;
                if (EditText0.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Project Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText0.Item.Click();
                    return;
                }
                //if (EditText2.Value == "")
                //{
                //    clsModule.objaddon.objapplication.StatusBar.SetText("Sales Order is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    BubbleEvent = false; EditText2.Item.Click();
                //    return;
                //}
                if (EditText6.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Date is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText6.Item.Click();
                    return;
                }
                if (EditText12.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Estimated Project is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; EditText12.Item.Click();
                    return;
                }
                
                Matrix0.FlushToDataSource();
                for (int ContentRow = 0; ContentRow <= odbdsContent.Size - 1; ContentRow++)
                {
                    if (odbdsContent.GetValue("U_SOEntry", ContentRow) != "" && odbdsContent.GetValue("U_Origin", ContentRow) == "Y")
                    {
                        originFlag = true;
                    }

                }
                if(originFlag==false) { clsModule.objaddon.objapplication.StatusBar.SetText("Please select the origin sales order...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); BubbleEvent = false;return; }
                //Matrix1.FlushToDataSource();

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Add Button

        private void EditText0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.ActionSuccess == true)
                    return;
                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item("cfl_prjcode");
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond;
                var oEmptyConds = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                oCond = oConds.Add();
                oCond.Alias = "Active";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";
                oCFL.SetConditions(oConds);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Project Code 

        private void EditText0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false)
                    return;
                pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pCFL.SelectedObjects != null)
                {
                    try
                    {
                        odbdsHeader.SetValue("Code", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrjCode").Cells.Item(0).Value));
                        odbdsHeader.SetValue("Name", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrjName").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }

                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Project Code 

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                if (pVal.ActionSuccess == true && objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                    clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "soentry", "#");
                }
            }
            catch (Exception ex)
            {
            }

        } //Add Button

        ////Sales Order
        //private void EditText2_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;
        //    if (EditText0.Value == "")
        //    {
        //        clsModule.objaddon.objapplication.StatusBar.SetText("Project Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //        BubbleEvent = false;
        //        return;
        //    }
        //    try
        //    {
        //        if (pVal.ActionSuccess == true) return;
        //        ChooseFromList_Condition("cfl_so", "Select distinct T0.\"DocEntry\" from ORDR T0 join RDR1 T1 on T0.\"DocEntry\"=T1.\"DocEntry\" where T1.\"Project\" ='"+ EditText0.Value + "'");

        //    }
        //    catch (Exception ex)
        //    {
        //        clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }

        //}

        //private void EditText2_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    try
        //    {
        //        if (pVal.ActionSuccess == false)
        //            return;
        //        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
        //        if (pCFL.SelectedObjects != null)
        //        {
        //            try
        //            {
        //                odbdsHeader.SetValue("U_SOEntry", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value));
        //                odbdsHeader.SetValue("U_SONo", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value));
        //                odbdsHeader.SetValue("U_CardCode", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value));
        //                odbdsHeader.SetValue("U_CardName", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value));
        //                odbdsHeader.SetValue("U_ProjValue", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocTotal").Cells.Item(0).Value));
        //            }
        //            catch (Exception ex)
        //            {
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }

        //}

        ////Engineer Name
        //private void EditText9_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;
        //    try
        //    {
        //        if (pVal.ActionSuccess == true)
        //            return;
        //        SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item("cfl_engcod");
        //        SAPbouiCOM.Conditions oConds;
        //        SAPbouiCOM.Condition oCond;
        //        var oEmptyConds = new SAPbouiCOM.Conditions();
        //        oCFL.SetConditions(oEmptyConds);
        //        oConds = oCFL.GetConditions();

        //        oCond = oConds.Add();
        //        oCond.Alias = "Active";
        //        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
        //        oCond.CondVal = "Y";
        //        oCFL.SetConditions(oConds);

        //    }
        //    catch (Exception ex)
        //    {
        //        clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }

        //}

        //private void EditText9_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        //{
        //    try
        //    {
        //        if (pVal.ActionSuccess == false)
        //            return;
        //        pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
        //        if (pCFL.SelectedObjects != null)
        //        {
        //            try
        //            {
        //                odbdsHeader.SetValue("U_EmpName", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("lastName").Cells.Item(0).Value)+","+ Convert.ToString(pCFL.SelectedObjects.Columns.Item("firstName").Cells.Item(0).Value));
        //                odbdsHeader.SetValue("U_EmpID", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("empID").Cells.Item(0).Value));
        //            }
        //            catch (Exception ex)
        //            {
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //    }

        //}

        #endregion

        #region Content Tab Events

        private void Folder0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {             
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mtxcont";
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "sonum", "#");
                Matrix0.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //Content Tab

        private void Matrix0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (EditText0.Value == "") { BubbleEvent = false; clsModule.objaddon.objapplication.StatusBar.SetText("Select Project Code!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); EditText0.Item.Click(); return; }

                if (pVal.ActionSuccess == true) return;

                if (pVal.ColUID == "sonum")
                    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("sonum").Cells.Item(pVal.Row).Specific).String == "") Matrix0.ClearRowData(pVal.Row);
                {
                    //ChooseFromList_QueryCondition("cfl_lso", "Select distinct T0.\"DocEntry\" from ORDR T0 join RDR1 T1 on T0.\"DocEntry\"=T1.\"DocEntry\" where T1.\"Project\" ='" + EditText0.Value + "' ");//and T1.\"DocEntry\" <>'" + EditText2.Value + "'
                    ChooseFromList_Condition("cfl_lso", "Project", EditText0.Value, "", "", "",Matrix0,"DocEntry","U_SOEntry","N",odbdsContent);
                    
                }


            }
            catch (Exception)
            {               
            }

        } //Content Matrix

        private void Matrix0_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;

                if (pVal.ColUID == "sonum")
                {
                    Matrix0.FlushToDataSource();
                    pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                    if (pCFL.SelectedObjects != null)
                    {
                        try
                        {
                            odbdsContent.SetValue("U_SONo", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value));
                            odbdsContent.SetValue("U_SOEntry", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value));
                            if (objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.Row == 1) odbdsContent.SetValue("U_Origin", pVal.Row - 1, "Y");
                            odbdsContent.SetValue("U_CardCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value));
                            odbdsContent.SetValue("U_CardName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value));                            
                            if (Convert.ToString(pCFL.SelectedObjects.Columns.Item("SlpCode").Cells.Item(0).Value) != "-1")
                            {
                                odbdsContent.SetValue("U_SlpCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("SlpCode").Cells.Item(0).Value));
                                strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"SlpName\" from OSLP Where \"SlpCode\"=" + pCFL.SelectedObjects.Columns.Item("SlpCode").Cells.Item(0).Value + "");
                                odbdsContent.SetValue("U_SlpName", pVal.Row - 1, strQuery);
                            }                           
                            if(Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_EngCode").Cells.Item(0).Value)!="")
                            {
                                odbdsContent.SetValue("U_EngCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("U_EngCode").Cells.Item(0).Value));
                                strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"lastName\" ||', '|| \"firstName\" \"Eng Name\" from OHEM Where \"empID\"=" + pCFL.SelectedObjects.Columns.Item("U_EngCode").Cells.Item(0).Value + "");
                                odbdsContent.SetValue("U_EngName", pVal.Row - 1, strQuery);
                            }                            
                            odbdsContent.SetValue("U_IProjValue", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocTotal").Cells.Item(0).Value));
                            strQuery = "Select sum(T0.\"CashSum\"+ T0.\"CreditSum\" + T0.\"CheckSum\" + T0.\"TrsfrSum\") \"Advance Amount\" from ORCT T0 join RCT2 T1 On T0.\"DocEntry\"=T1.\"DocNum\" Where T0.\"CardCode\"='" + Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value) + "' and T0.\"Canceled\"='N'";
                            strQuery += "\n and T1.\"DocEntry\" in (Select \"DocEntry\" from DPI1 Where \"BaseType\"='17' and \"BaseEntry\"='"+ Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value) + "')";
                            strQuery = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);
                            odbdsContent.SetValue("U_Advance", pVal.Row - 1, strQuery);                            
                            if (Convert.ToDouble(odbdsContent.GetValue("U_NetAdd", pVal.Row - 1))==0) odbdsContent.SetValue("U_NetAdd", pVal.Row - 1,Convert.ToString(0));
                            if (Convert.ToDouble(odbdsContent.GetValue("U_NetDed", pVal.Row - 1)) == 0) odbdsContent.SetValue("U_NetDed", pVal.Row - 1, Convert.ToString(0));
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    Matrix0.LoadFromDataSource();
                    Calculate_ProjectValue();

                }

                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
            }
            catch (Exception)
            {

                //throw;
            }

        } //Content Matrix

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "sonum":
                    case "netadd":
                    case "netded":
                    case "origin":
                        Calculate_ProjectValue();
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "sonum", "#");                       
                        if(pVal.ColUID!="origin")Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                }
                objform.Freeze(true);
                Matrix0.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        } //Content Matrix        

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {                
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                switch (pVal.ColUID)
                {
                    case "sonum":
                        BubbleEvent = false;
                        objRs.DoQuery("select \"DocNum\",To_Varchar(\"DocDate\",'yyyyMMdd') \"DocDate\" from ORDR where \"DocEntry\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("soentry").Cells.Item(pVal.Row).Specific).String + "'");
                        if (objRs.RecordCount > 0)
                        {
                            clsModule.objaddon.objapplication.Menus.Item("2050").Activate();                            
                            TempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            TempForm.Freeze(true);
                            TempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("8").Specific).String = Convert.ToString(objRs.Fields.Item(0).Value);
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("10").Specific).String = Convert.ToString(objRs.Fields.Item(1).Value);
                            TempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            TempForm.Freeze(false);
                        }
                        break;                    
                }                
             

            }
            catch (Exception ex)
            {
                TempForm.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Content Matrix

        private void Matrix0_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (pVal.ColUID)
                {
                    case "origin":
                        SAPbouiCOM.CheckBox checkBox = (SAPbouiCOM.CheckBox)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific;
                        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("sonum").Cells.Item(pVal.Row).Specific).String == "") { checkBox.Checked = false;  Matrix0.Columns.Item("sonum").Cells.Item(pVal.Row).Click(); BubbleEvent = false; return; }
                        for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                        {
                            checkBox = (SAPbouiCOM.CheckBox)Matrix0.Columns.Item(pVal.ColUID).Cells.Item(i).Specific;
                            if (i != pVal.Row)
                            {
                                if (checkBox.Checked == true) checkBox.Checked = false;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {

            }


        } //Content Matrix

        private void Matrix0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.InnerEvent == true) return;
                if (pVal.Row == 0)
                {
                    BubbleEvent = false; Matrix0.Item.Click();
                }

            }
            catch (Exception ex)
            {
            }

        } //Content Matrix

        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.CharPressed == 9 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None && pVal.ColUID!="origin") Matrix0.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                int colID = Matrix0.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) //UP
                {
                    Matrix0.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) //Down 
                {
                    Matrix0.SetCellFocus(pVal.Row + 1, colID);
                }


            }
            catch (Exception ex)
            {
            }

        } //Content Matrix

        #endregion

        #region Attachment Tab Events
        //Attachment 

        private void Folder11_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {            
            try
            {
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mtxattach";
                Matrix1.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }  //Attachment Tab

        private void Matrix1_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix1, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {

            }

        } //Attachment Matrix

        private void Matrix1_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Matrix1.SelectRow(pVal.Row, true, false);
                if (Matrix1.IsRowSelected(pVal.Row) == true)
                {
                    objform.Items.Item("btndisp").Enabled = true;
                    objform.Items.Item("btndel").Enabled = true;
                }
            }
            catch (Exception ex)
            {

            }

        } //Attachment Matrix
        
        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.SetAttachMentFile(objform, odbdsHeader, Matrix1, odbdsAttachment);
                if (Matrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder) == -1)
                {
                    objform.Items.Item("btndisp").Enabled = false;
                    objform.Items.Item("btndel").Enabled = false;
                }
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }  //Browse Attachment
        
        private void Button3_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                //if (pVal.ActionSuccess == false) return;
                clsModule.objaddon.objglobalmethods.OpenAttachment(Matrix1, odbdsAttachment, pVal.Row);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }  //Display Attachment
        
        private void Button4_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                clsModule.objaddon.objglobalmethods.DeleteRowAttachment(objform, Matrix1, odbdsAttachment, Matrix1.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }


        } //Delete Attachment

        #endregion

        #region BOQ Tab Events

        private void Folder3_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Folder6.Item.Click();
            }
            catch (Exception ex)
            {
            }
        } //BOQ Tab

        private void Folder6_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "itemcode", "#");
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mboqitem";
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "sonum", "#");
                Matrix2.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        } //BOQ Item Tab

        private void Matrix2_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (EditText0.Value == "") { BubbleEvent = false; clsModule.objaddon.objapplication.StatusBar.SetText("Select Project Code!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); EditText0.Item.Click(); return; }

                if (pVal.ActionSuccess == true) return;

                if (pVal.ColUID == "itemcode")
                {
                    ChooseFromList_Condition("cflitem", "validFor", "Y", "", "", "");
                }
                else if (pVal.ColUID == "sonum")
                {
                    //if (((SAPbouiCOM.EditText)Matrix2.Columns.Item("itemcode").Cells.Item(pVal.Row).Specific).String == "") { BubbleEvent = false; return; }
                    //ChooseFromList_QueryCondition("cfl_lso1", "Select distinct T0.\"DocEntry\" from ORDR T0 join RDR1 T1 on T0.\"DocEntry\"=T1.\"DocEntry\" where T1.\"Project\" ='" + EditText0.Value + "' ");
                    ChooseFromList_Condition("cfl_lso1", "Project", EditText0.Value, "", "", "",Matrix0,"DocEntry","U_SOEntry","Y", odbdsContent);
                }
                else if (pVal.ColUID == "cc1")
                {
                    ChooseFromList_Condition ("itemcc1", "DimCode", "1", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc2")
                {
                    ChooseFromList_Condition("itemcc2", "DimCode", "2", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc3")
                {
                    ChooseFromList_Condition("itemcc3", "DimCode", "3", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc4")
                {
                    ChooseFromList_Condition("itemcc4", "DimCode", "4", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc5")
                {
                    ChooseFromList_Condition("itemcc5", "DimCode", "5", "", "Locked", "N");
                }


            }
            catch (Exception)
            {
            }

        } //BOQ Item Matrix

        private void Matrix2_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pCFL.SelectedObjects == null) return;
                Matrix2.FlushToDataSource();
                if (pVal.ColUID == "itemcode")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_ItemCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("ItemCode").Cells.Item(0).Value));
                        odbdsBoqItem.SetValue("U_ItemName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("ItemName").Cells.Item(0).Value));
                        odbdsBoqItem.SetValue("U_Quantity", pVal.Row - 1, "1");
                        odbdsBoqItem.SetValue("U_Project", pVal.Row - 1, EditText0.Value);
                        strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"UomCode\" from OUOM Where \"UomEntry\"=" + pCFL.SelectedObjects.Columns.Item("UgpEntry").Cells.Item(0).Value + "");
                        odbdsBoqItem.SetValue("U_Uom", pVal.Row - 1, strQuery);
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "project")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_Project", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrjCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "sonum")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_SONo", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value));
                        odbdsBoqItem.SetValue("U_SOEntry", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value));
                        strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"OcrCode\" from RDR1 Where \"DocEntry\"=" + pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value + "");
                        odbdsBoqItem.SetValue("U_OcrCode", pVal.Row - 1, strQuery);
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc1")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_OcrCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc2")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_OcrCode2", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc3")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_OcrCode3", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc4")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_OcrCode4", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc5")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_OcrCode5", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "uom")
                {
                    try
                    {
                        odbdsBoqItem.SetValue("U_Uom", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("UomCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                Matrix2.LoadFromDataSource();
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
            }
            catch (Exception)
            {

            }

        } //BOQ Item Matrix

        private void Matrix2_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "sonum":                        
                    case "qty":
                    case "price":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix2, "sonum", "#");
                        objform.Freeze(true);
                        Matrix2.FlushToDataSource();
                        double Total = Convert.ToDouble(odbdsBoqItem.GetValue("U_Quantity", pVal.Row - 1).ToString()) * Convert.ToDouble(odbdsBoqItem.GetValue("U_UnitPrice", pVal.Row - 1).ToString());
                        odbdsBoqItem.SetValue("U_Total", pVal.Row - 1,Convert.ToString(Total));
                        Matrix2.LoadFromDataSource();
                        Calculate_ProjectValue();
                        Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        objform.Freeze(false);
                        if (((SAPbouiCOM.EditText)Matrix2.Columns.Item("sonum").Cells.Item(pVal.Row).Specific).String=="") return;
                        if (pVal.ColUID=="qty" || pVal.ColUID == "price")
                        {
                            if (Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String) <= 0)
                            {
                                ((SAPbouiCOM.EditText)Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String = "1";
                                Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                                clsModule.objaddon.objapplication.StatusBar.SetText("In \"" + Matrix2.Columns.Item(pVal.ColUID).UniqueID + "\" (" + Matrix2.Columns.Item(pVal.ColUID).TitleObject.Caption + ") column, enter value greater than 0 ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }
                        }
                        break;
                }
                objform.Freeze(true);
                Matrix2.AutoResizeColumns();
                objform.Freeze(false);

            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        } //BOQ Item Matrix

        private void Matrix2_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                switch (pVal.ColUID)
                {
                    case "sonum":
                        BubbleEvent = false;
                        objRs.DoQuery("select \"DocNum\",To_Varchar(\"DocDate\",'yyyyMMdd') \"DocDate\" from ORDR where \"DocEntry\"='" + ((SAPbouiCOM.EditText)Matrix2.Columns.Item("soentry").Cells.Item(pVal.Row).Specific).String + "'");
                        if (objRs.RecordCount > 0)
                        {
                            clsModule.objaddon.objapplication.Menus.Item("2050").Activate();
                            TempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            TempForm.Freeze(true);
                            TempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("8").Specific).String = Convert.ToString(objRs.Fields.Item(0).Value);
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("10").Specific).String = Convert.ToString(objRs.Fields.Item(1).Value);
                            TempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            TempForm.Freeze(false);
                        }
                        break;
                }


            }
            catch (Exception ex)
            {
                TempForm.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //BOQ Item Matrix

        private void Matrix2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.InnerEvent == true) return;
                if (pVal.Row == 0)
                {
                    BubbleEvent = false; Matrix2.Item.Click();
                }

            }
            catch (Exception ex)
            {
            }

        } //BOQ Item Matrix 

        private void Matrix2_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.CharPressed == 9 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) Matrix2.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                int colID = Matrix2.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Up
                {
                    Matrix2.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None)//Down
                {
                    Matrix2.SetCellFocus(pVal.Row + 1, colID);
                }

            }
            catch (Exception ex)
            {
            }

        }

        private void Folder7_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE) clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "code", "#");
                objform.Freeze(true);
                objform.Settings.MatrixUID = "mboqlab";
                clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "sonum", "#");
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }

        }  //BOQ Labour Tab          
               
        private void Matrix3_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                switch (pVal.ColUID)
                {
                    case "sonum":
                        BubbleEvent = false;
                        objRs.DoQuery("select \"DocNum\",To_Varchar(\"DocDate\",'yyyyMMdd') \"DocDate\" from ORDR where \"DocEntry\"='" + ((SAPbouiCOM.EditText)Matrix3.Columns.Item("soentry").Cells.Item(pVal.Row).Specific).String + "'");
                        if (objRs.RecordCount > 0)
                        {
                            clsModule.objaddon.objapplication.Menus.Item("2050").Activate();
                            TempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            TempForm.Freeze(true);
                            TempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("8").Specific).String = Convert.ToString(objRs.Fields.Item(0).Value);
                            ((SAPbouiCOM.EditText)TempForm.Items.Item("10").Specific).String = Convert.ToString(objRs.Fields.Item(1).Value);
                            TempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            TempForm.Freeze(false);
                        }
                        break;
                    case "code":
                        SAPbouiCOM.Column column = Matrix3.Columns.Item("code");
                        SAPbouiCOM.LinkedButton linkedButton = (SAPbouiCOM.LinkedButton)column.ExtendedObject;
                        linkedButton.LinkedObjectType = ((SAPbouiCOM.EditText)Matrix3.Columns.Item("objtype").Cells.Item(pVal.Row).Specific).String;
                        break;
                }


            }
            catch (Exception ex)
            {
                TempForm.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //BOQ Labour Matrix    
             
        private void Matrix3_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.CharPressed == 9 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                int colID = Matrix3.GetCellFocus().ColumnIndex;
                if (pVal.CharPressed == 38 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) //Up
                {
                    Matrix3.SetCellFocus(pVal.Row - 1, colID);
                }
                else if (pVal.CharPressed == 40 && pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_None) //Down
                {
                    Matrix3.SetCellFocus(pVal.Row + 1, colID);
                }
                if (pVal.ColUID == "code")
                {
                    SAPbouiCOM.Column column = Matrix3.Columns.Item("code");
                    if (pVal.Modifiers == SAPbouiCOM.BoModifiersEnum.mt_CTRL && pVal.CharPressed == 9)
                    {
                        column.ChooseFromListUID.Remove(pVal.Row);
                        column.ChooseFromListUID = "cfllabbp";
                        ChooseFromList_Condition("cfllabbp", "validFor", "Y", "", "", "");
                    }
                    else
                    {
                        column.ChooseFromListUID.Remove(pVal.Row);
                        column.ChooseFromListUID = "cfl_gl";
                        ChooseFromList_Condition("cfl_gl", "Postable", "Y", "", "LocManTran", "N");
                    }
                }              
                
            }
            catch (Exception ex)
            {

            }


        } //BOQ Labour Matrix    

        private void Matrix3_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                switch (pVal.ColUID)
                {
                    case "sonum":
                    case "total":
                        clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix3, "sonum", "#");
                        Calculate_ProjectValue();
                        Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
                        break;
                }
                objform.Freeze(true);
                Matrix3.AutoResizeColumns();
                objform.Freeze(false);

            }
            catch (Exception)
            {
                objform.Freeze(false);
            }

        } //BOQ Labour Matrix

        private void Matrix3_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.InnerEvent == true) return;
                if (pVal.Row == 0)
                {
                    BubbleEvent = false; Matrix3.Item.Click();
                }

            }
            catch (Exception ex)
            {
            }

        } //BOQ Labour Matrix 

        private void Matrix3_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (EditText0.Value == "") { BubbleEvent = false; clsModule.objaddon.objapplication.StatusBar.SetText("Select Project Code!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); EditText0.Item.Click(); return; }

                if (pVal.ActionSuccess == true) return;

                if (pVal.ColUID == "sonum")
                {
                    //if (((SAPbouiCOM.EditText)Matrix3.Columns.Item("code").Cells.Item(pVal.Row).Specific).String == "") { BubbleEvent = false; return; }
                    //ChooseFromList_QueryCondition("cfl_lso2", "Select distinct T0.\"DocEntry\" from ORDR T0 join RDR1 T1 on T0.\"DocEntry\"=T1.\"DocEntry\" where T1.\"Project\" ='" + EditText0.Value + "' ");
                    ChooseFromList_Condition("cfl_lso2", "Project", EditText0.Value, "", "", "", Matrix0, "DocEntry", "U_SOEntry", "Y", odbdsContent);
                }
                else if(pVal.ColUID== "cosglc")
                {
                    if (((SAPbouiCOM.ComboBox)Matrix3.Columns.Item("labtype").Cells.Item(pVal.Row).Specific).Selected.Value=="-") { BubbleEvent = false;  clsModule.objaddon.objapplication.StatusBar.SetText("Select Labour Type!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); Matrix3.Columns.Item("labtype").Cells.Item(pVal.Row).Click(); ; return; }
                    ChooseFromList_Condition("cfl_cosgl", "Postable", "Y", "", "LocManTran", "N");
                }
                else if (pVal.ColUID == "cc1")
                {
                    ChooseFromList_Condition("labcc1", "DimCode", "1", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc2")
                {
                    ChooseFromList_Condition("labcc2", "DimCode", "2", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc3")
                {
                    ChooseFromList_Condition("labcc3", "DimCode", "3", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc4")
                {
                    ChooseFromList_Condition("labcc4", "DimCode", "4", "", "Locked", "N");
                }
                else if (pVal.ColUID == "cc5")
                {
                    ChooseFromList_Condition("labcc5", "DimCode", "5", "", "Locked", "N");
                }

            }
            catch (Exception)
            {
            }

        } //BOQ Labour Matrix          

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                if (pVal.CharPressed == 9 & EditText1.Value != "")
                {
                    if (clsModule.objaddon.HANA == true)
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 \"Status\" from \"@AT_PROJMSTR\" where \"Code\"='" + EditText0.Value + "' ");
                    else
                        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("select 1 Status from [@AT_PROJMSTR] where Code='" + EditText0.Value + "' ");

                    if (strSQL == "1")
                    {
                        strSQL = EditText0.Value;
                        objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                        EditText0.Value = strSQL;
                        objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        Folder0.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        return;
                    }
                }

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("E_Key_Down_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                bool flag = false;
                Matrix0.FlushToDataSource();
                for (int ContentRow = 0; ContentRow <= odbdsContent.Size - 1; ContentRow++)
                {
                    if (odbdsContent.GetValue("U_SOEntry", ContentRow) != "" && odbdsContent.GetValue("U_Origin", ContentRow) == "Y")
                    {
                        odbdsContent.GetValue("U_Advance", ContentRow);
                        strQuery = "Select SUM(T0.\"CashSum\"+ T0.\"CreditSum\" + T0.\"CheckSum\" + T0.\"TrsfrSum\") \"Advance Amount\" from ORCT T0 join RCT2 T1 On T0.\"DocEntry\"=T1.\"DocNum\" Where T0.\"CardCode\"='" + Convert.ToString(odbdsContent.GetValue("U_CardCode", ContentRow)) + "' and T0.\"Canceled\"='N'";
                        strQuery += "\n and T1.\"DocEntry\" in (Select \"DocEntry\" from DPI1 Where \"BaseType\"='17' and \"BaseEntry\"='" + Convert.ToString(odbdsContent.GetValue("U_SOEntry", ContentRow)) + "')";
                        strQuery = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);
                        if (Convert.ToDouble( odbdsContent.GetValue("U_Advance", ContentRow)) != Convert.ToDouble(strQuery)) flag = true;
                        odbdsContent.SetValue("U_Advance", ContentRow, strQuery);
                        break;
                    }
                }
                Matrix0.LoadFromDataSource();
                if (flag == true)
                {
                    if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    objform.Items.Item("1").Click();
                }
              

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DataLoad_After: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix3_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ColUID == "labtype")
                {
                    if (((SAPbouiCOM.ComboBox)Matrix3.Columns.Item("labtype").Cells.Item(pVal.Row).Specific).Selected.Value != "-") return;
                    Matrix3.FlushToDataSource();
                    odbdsBoqLabour.SetValue("U_Cosglc", pVal.Row - 1, "");
                    odbdsBoqLabour.SetValue("U_Cosgln", pVal.Row - 1, "");
                    Matrix3.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {

            }

        }

        private void Matrix3_ChooseFromListAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == false) return;
                pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (pCFL.SelectedObjects == null) return;
                Matrix3.FlushToDataSource();
                if (pVal.ColUID == "code")
                {
                    try
                    {
                        if (Convert.ToString(pCFL.SelectedObjects.Columns.Item("ObjType").Cells.Item(0).Value) == "1")
                        {
                            odbdsBoqLabour.SetValue("U_AcctCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value));
                            odbdsBoqLabour.SetValue("U_AcctName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value));
                        }
                        else
                        {
                            odbdsBoqLabour.SetValue("U_AcctCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardCode").Cells.Item(0).Value));
                            odbdsBoqLabour.SetValue("U_AcctName", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value));
                        }

                        odbdsBoqLabour.SetValue("U_Project", pVal.Row - 1, EditText0.Value);
                        odbdsBoqLabour.SetValue("U_ObjType", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("ObjType").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "sonum")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_SONo", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocNum").Cells.Item(0).Value));
                        odbdsBoqLabour.SetValue("U_SOEntry", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value));
                        strQuery = clsModule.objaddon.objglobalmethods.getSingleValue("Select \"OcrCode\" from RDR1 Where \"DocEntry\"=" + pCFL.SelectedObjects.Columns.Item("DocEntry").Cells.Item(0).Value + "");
                        odbdsBoqLabour.SetValue("U_OcrCode", pVal.Row - 1, strQuery);
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cosglc")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_Cosglc", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value));
                        odbdsBoqLabour.SetValue("U_Cosgln", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc1")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_OcrCode", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc2")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_OcrCode2", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc3")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_OcrCode3", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc4")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_OcrCode4", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                else if (pVal.ColUID == "cc5")
                {
                    try
                    {
                        odbdsBoqLabour.SetValue("U_OcrCode5", pVal.Row - 1, Convert.ToString(pCFL.SelectedObjects.Columns.Item("PrcCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                }
                Matrix3.LoadFromDataSource();
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                Matrix3.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click();
            }
            catch (Exception)
            {

            }

        } //BOQ Labour Matrix  

        #endregion      
        
        #region Functions

        private void ChooseFromList_QueryCondition(string CFLID, string Query)
        {
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item(CFLID);
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond = null;
                var oEmptyConds = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                if (Query != "")
                {
                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    objRs.DoQuery(Query);
                    if (objRs.RecordCount > 0)
                    {
                        for (int i = 0; i < objRs.RecordCount; i++)
                        {
                            if (i == 0)
                            {
                                oCond = oConds.Add();
                                oCond.Alias = objRs.Fields.Item(0).Name;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = objRs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                oCond = oConds.Add();
                                oCond.Alias = objRs.Fields.Item(0).Name;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = objRs.Fields.Item(0).Value.ToString();
                            }
                            objRs.MoveNext();
                        }
                    }

                }

                oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void ChooseFromList_Condition(string CFLID, string Alias, string CondVal, string Query, string Alias1, string CondVal1,SAPbouiCOM.Matrix matrix=null,string matrixAlias="",string MatrixCol="", string MatConEqual = "Y",SAPbouiCOM.DBDataSource dBDataSource=null)
        {
            try
            {
                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item(CFLID);
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond;
                var oEmptyConds = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                oCond = oConds.Add();
                oCond.Alias = Alias;// "Postable";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CondVal;// "Y";
                if (Alias1 != "" && CondVal1 != "")
                {
                    oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                    oCond = oConds.Add();
                    oCond.Alias = Alias1;
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal = CondVal1;
                }
                if (Query != "")
                {
                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    objRs.DoQuery(Query);
                    if (objRs.RecordCount > 0)
                    {
                        for (int i = 0; i < objRs.RecordCount; i++)
                        {
                            if (i == 0)
                            {
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                                oCond = oConds.Add();
                                oCond.Alias = objRs.Fields.Item(0).Name;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = objRs.Fields.Item(0).Value.ToString();
                            }
                            else
                            {
                                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR;
                                oCond = oConds.Add();
                                oCond.Alias = objRs.Fields.Item(0).Name;
                                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCond.CondVal = objRs.Fields.Item(0).Value.ToString();
                            }
                            objRs.MoveNext();
                        }
                    }
                    else
                    {
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = Alias;
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE;
                        oCond.CondVal = CondVal;
                    }
                }

                if (matrix != null)
                {
                    matrix.FlushToDataSource();
                    for (int Row = 0; Row <= dBDataSource.Size - 1; Row++)
                    {
                        if (Row == dBDataSource.Size - 1) ;// oCond.BracketCloseNum = 1;
                        if (dBDataSource.GetValue(MatrixCol, Row) == "") continue;
                        if (oConds.Count > 0 && Row == 0) { oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND; }//oCond.BracketOpenNum = 1;
                        if (MatConEqual == "Y") oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR; 
                        else oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                        oCond = oConds.Add();
                        oCond.Alias = matrixAlias;
                        if (MatConEqual=="Y")
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        else
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                        oCond.CondVal = dBDataSource.GetValue(MatrixCol, Row);
                        
                    }
                   
                }

                oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void CostCenter()
        {
            try
            {
                //***********BOQ Item***************
                Matrix2.Columns.Item("cc1").Visible = false;
                Matrix2.Columns.Item("cc2").Visible = false;
                Matrix2.Columns.Item("cc3").Visible = false;
                Matrix2.Columns.Item("cc4").Visible = false;
                Matrix2.Columns.Item("cc5").Visible = false;

                //***********BOQ Labour***********

                Matrix3.Columns.Item("cc1").Visible = false;
                Matrix3.Columns.Item("cc2").Visible = false;
                Matrix3.Columns.Item("cc3").Visible = false;
                Matrix3.Columns.Item("cc4").Visible = false;
                Matrix3.Columns.Item("cc5").Visible = false;

                objRs = (SAPbobsCOM.Recordset) clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (clsModule.objaddon.HANA)
                    objRs.DoQuery("select 'cc'||\"DimCode\" as \"Code\",* from ODIM where \"DimActive\"='Y'");
                else
                    objRs.DoQuery("select CONCAT('cc',DimCode) as Code,* from ODIM where DimActive='Y'");
 
                if (objRs.RecordCount>0)
                {
                    for (int i = 0; i < objRs.RecordCount ; i++)
                    {
                        Matrix2.Columns.Item(objRs.Fields.Item("Code").Value).Visible = true;
                        Matrix2.Columns.Item(objRs.Fields.Item("Code").Value).TitleObject.Caption =(string)objRs.Fields.Item("DimDesc").Value;
                        Matrix3.Columns.Item(objRs.Fields.Item("Code").Value).Visible = true;
                        Matrix3.Columns.Item(objRs.Fields.Item("Code").Value).TitleObject.Caption = (string)objRs.Fields.Item("DimDesc").Value;
                        objRs.MoveNext();
                    }
                }
                 
            }
            catch (Exception)
            {

            }
        }

        public void Calculate_ProjectValue()
        {
            try
            {
                decimal InitProjValue = 0, EstiCost=0,EstiProjValue=0;
                objform.Freeze(true);
                Matrix0.FlushToDataSource();
                for (int ContentRow = 0; ContentRow <= odbdsContent.Size - 1; ContentRow++)
                {
                    if (odbdsContent.GetValue("U_SOEntry", ContentRow) != "" && odbdsContent.GetValue("U_Origin", ContentRow) == "Y")
                    {
                       InitProjValue = ((Convert.ToDecimal(odbdsContent.GetValue("U_IProjValue", ContentRow))+ Convert.ToDecimal(odbdsContent.GetValue("U_NetAdd", ContentRow)))- Convert.ToDecimal(odbdsContent.GetValue("U_NetDed", ContentRow)));
                       
                    }
                    Matrix2.FlushToDataSource();
                    for (int BOQItemRow = 0; BOQItemRow <= odbdsBoqItem.Size - 1; BOQItemRow++) //BOQ Item 
                    {
                        if (odbdsBoqItem.GetValue("U_SOEntry", BOQItemRow)!="" && odbdsBoqItem.GetValue("U_SOEntry", BOQItemRow) == odbdsContent.GetValue("U_SOEntry", ContentRow))
                        {
                            EstiCost = EstiCost+ Convert.ToDecimal(odbdsBoqItem.GetValue("U_Total", BOQItemRow));
                        }
                    }
                    Matrix3.FlushToDataSource();
                    for (int BOQLabourRow = 0; BOQLabourRow <= odbdsBoqLabour.Size - 1; BOQLabourRow++) //BOQ Labour
                    {
                        if (odbdsBoqLabour.GetValue("U_SOEntry", BOQLabourRow)!="" && odbdsBoqLabour.GetValue("U_SOEntry", BOQLabourRow) == odbdsContent.GetValue("U_SOEntry", ContentRow))
                        {
                            EstiCost = EstiCost + Convert.ToDecimal(odbdsBoqLabour.GetValue("U_Total", BOQLabourRow));
                        }
                    }
                    odbdsContent.SetValue("U_EstValue", ContentRow, Convert.ToString(EstiCost)); //Estimated Cost
                    EstiCost = 0;
                    EstiProjValue = EstiProjValue + Convert.ToDecimal(odbdsContent.GetValue("U_EstValue", ContentRow));
                }
                Matrix0.LoadFromDataSource();
                odbdsHeader.SetValue("U_ProjValue", 0, Convert.ToString(InitProjValue)); //Initial Project Cost
                //string cost = Matrix0.Columns.Item("estval").ColumnSetting.SumValue;
                odbdsHeader.SetValue("U_EstProjValue", 0, Convert.ToString(EstiProjValue)); //Estimated Project Cost
                objform.Freeze(false);

            }
            catch (Exception ex)
            {
                objform.Freeze(false);
            }
        }

        private void Manage_Fields()
        {
            try
            {
                ComboBox1.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly; //com/Res Field
                ComboBox2.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly; //Status Field
                ComboBox0.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly; //Project Type Field
                ComboBox3.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly; //POC Field
                Matrix3.Columns.Item("labtype").ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly;
                Matrix0.Columns.Item("soentry").Visible = false;
                Matrix0.Columns.Item("slpcode").Visible = false;
                Matrix0.Columns.Item("engcode").Visible = false;
                Matrix1.Columns.Item("srcpath").Visible = false;
                Matrix1.Columns.Item("fileext").Visible = false;
                Matrix1.Columns.Item("trgtpath").Editable = false;
                Matrix1.Columns.Item("filename").Editable = false;
                Matrix1.Columns.Item("date").Editable = false;
                Matrix2.Columns.Item("soentry").Visible = false;
                Matrix3.Columns.Item("soentry").Visible = false;
                Matrix3.Columns.Item("objtype").Visible = false;
                Matrix0.Columns.Item("estval").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix2.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                Matrix3.Columns.Item("total").ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            }
            catch (Exception ex)
            {
            }
        }



        #endregion

        
    }
}
