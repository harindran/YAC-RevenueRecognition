using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using RevenueRecognition.Common;
using System.Drawing;
using SAPbobsCOM;
using System.Globalization;
using Microsoft.VisualBasic;

namespace RevenueRecognition.Business_Objects
{
    [FormAttribute("REVREC", "Business_Objects/FrmRevenueRecognition.b1f")]
    class FrmRevenueRecognition : UserFormBase
    {
        public FrmRevenueRecognition()
        {
        }
        public static SAPbouiCOM.Form objform;
        private string strSQL, strQuery, Localization,JournalID;
        private bool tranflag = false;
        private SAPbobsCOM.Recordset objRs,Recordset;
        SAPbouiCOM.ISBOChooseFromListEventArg pCFL;
        public SAPbouiCOM.DBDataSource odbdsHeader, odbdsDetails;
        int errorCode;
        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lglc").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tglc").Specific));
            this.EditText0.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText0_ChooseFromListAfter);
            this.EditText0.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText0_ChooseFromListBefore);
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lno").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("tdocnum").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("series").Specific));
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lgln").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("tgln").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("fldrcont").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lrem").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("trem").Specific));
            this.Folder1 = ((SAPbouiCOM.Folder)(this.GetItem("fldrnew").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lposdate").Specific));
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("tposdate").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkglc").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lvochid").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("tvochid").Specific));
            this.LinkedButton1 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lktran").Specific));
            this.LinkedButton1.PressedBefore += new SAPbouiCOM._ILinkedButtonEvents_PressedBeforeEventHandler(this.LinkedButton1_PressedBefore);
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("lstatus").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("cstatus").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("bgetdata").Specific));
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("cmonth").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("lyear").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("cyear").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("lprojtype").Specific));
            this.ComboBox4 = ((SAPbouiCOM.ComboBox)(this.GetItem("cprojtype").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("mtxcont").Specific));
            this.Matrix0.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix0_ClickAfter);
            this.Matrix0.LinkPressedBefore += new SAPbouiCOM._IMatrixEvents_LinkPressedBeforeEventHandler(this.Matrix0_LinkPressedBefore);
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("ljeid").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("tjeid").Specific));
            this.LinkedButton2 = ((SAPbouiCOM.LinkedButton)(this.GetItem("lkjeid").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.ResizeAfter += new SAPbouiCOM.Framework.FormBase.ResizeAfterHandler(this.Form_ResizeAfter);
            this.DataAddBefore += new SAPbouiCOM.Framework.FormBase.DataAddBeforeHandler(this.Form_DataAddBefore);
            this.DataLoadAfter += new SAPbouiCOM.Framework.FormBase.DataLoadAfterHandler(this.Form_DataLoadAfter);
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }


        private void OnCustomInitialize()
        {
            try
            {                
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_REV_RECO");
                odbdsDetails = objform.DataSources.DBDataSources.Item("@AT_REV_RECO1");
                clsModule.objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "AT_REVREC");
                ((SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "invnum", "#");
                strSQL = "Select * from \"@AT_PROJTYPE\"";
                
                clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("cprojtype").Specific), strSQL, new[] { "-,-" });
                Load_Data(); // Loading Year & Month    
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "bgetdata", true, false, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "series", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tglc", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "tposdate", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cstatus", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cprojtype", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cyear", true, true, false);
                clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Editable(objform, "cmonth", true, true, false);
                //clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Visible(objform, "tjeid", false,true,true);
                //clsModule.objaddon.objglobalmethods.SetAutomanagedattribute_Visible(objform, "ljeid", false, true, true);
                Matrix0.CommonSetting.FixedColumnsCount = 9;
                Matrix0.Columns.Item("ocrcode2").Visible = false;
                Matrix0.Columns.Item("ocrcode3").Visible = false;
                Matrix0.Columns.Item("ocrcode4").Visible = false;
                Matrix0.Columns.Item("ocrcode5").Visible = false;
                //strSQL = Matrix0.Columns.Item(2).DataBind.Alias;
                EditText6.Item.Visible = false;
                StaticText6.Item.Visible = false;
                LinkedButton2.Item.Visible = false;
                Matrix0.AutoResizeColumns();
                Folder0.Item.Click();
                objform.ActiveItem = "tposdate";
                objform.Settings.Enabled = true;

                //********************** Dynamic UDF Creation in Line Level of Matrix **************************************
                if (clsModule.objaddon.HANA == true)
                {
                    strSQL = "Select \"USERID\",\"TPLId\" from OUSR Where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'";
                    strQuery = "Select '@' || \"SonName\" \"TableName\" from UDO1 Where \"Code\" = '" + objform.BusinessObject.Type + "'";
                }                    
                else
                {
                    strSQL = "Select USERID,TPLId from OUSR Where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'";
                    strQuery = "Select '@' + SonName TableName from UDO1 Where Code = '" + objform.BusinessObject.Type + "'";
                }          
               
                Recordset = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                Recordset.DoQuery(strQuery);

                Dictionary<string, string> Table_Matrix = new Dictionary<string, string>(); // Adding Matrix UID of each Line Table
                List<string> MatrixIDs = new List<string>();
                string MatID;
                Table_Matrix.Add("mtxcont", "@AT_REV_RECO1");
                

                if (Recordset.RecordCount > 0)
                {
                    for (int i = 0; i < Recordset.RecordCount; i++)
                    {                        
                        if (!Table_Matrix.ContainsValue(Convert.ToString(Recordset.Fields.Item("TableName").Value))) continue;

                        foreach (var pair in Table_Matrix)
                        {
                            if (pair.Value == Convert.ToString(Recordset.Fields.Item("TableName").Value))
                            {
                                MatrixIDs.Add(pair.Key);
                            }
                        }
                        MatID = String.Format("'{0}'", String.Join("','", MatrixIDs));
                        //strSQL = Table_Matrix[Convert.ToString(Recordset.Fields.Item("TableName").Value)];

                        clsModule.objaddon.objglobalmethods.Create_Dynamic_LineTable_UDF(objform, Convert.ToString(Recordset.Fields.Item("TableName").Value), objform.TypeEx, String.Format("'{0}'", String.Join("','", MatrixIDs)));
                        Recordset.MoveNext();
                    }                    
                }
                //********************** Dynamic UDF END **************************************

                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                objRs.DoQuery(strSQL);

                clsModule.objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform,"-" + objform.TypeEx, Convert.ToInt32(objRs.Fields.Item("USERID").Value), Convert.ToInt32(objRs.Fields.Item("TPLId").Value)); //REVREC

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #region Fields
        
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.Folder Folder0;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.EditText EditText3;
        private SAPbouiCOM.Folder Folder1;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.EditText EditText4;
        private SAPbouiCOM.LinkedButton LinkedButton0;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.LinkedButton LinkedButton1;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.ComboBox ComboBox4;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.LinkedButton LinkedButton2;

        #endregion

        #region Choose From List Events

        private void EditText0_ChooseFromListBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //if (EditText0.Value != "") EditText0.Value = "";
                if (EditText2.Value != ""& EditText0.Value == "") EditText2.Value = "";
                if (pVal.ActionSuccess == true)
                    return;
                SAPbouiCOM.ChooseFromList oCFL = objform.ChooseFromLists.Item("cfl_glc");
                SAPbouiCOM.Conditions oConds;
                SAPbouiCOM.Condition oCond;
                var oEmptyConds = new SAPbouiCOM.Conditions();
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                oCond = oConds.Add();
                oCond.Alias = "Postable";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "Y";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = "LocManTran";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = "N";               
                oCFL.SetConditions(oConds);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

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
                        odbdsHeader.SetValue("U_GLCode", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctCode").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                    try
                    {
                        odbdsHeader.SetValue("U_GLName", 0, Convert.ToString(pCFL.SelectedObjects.Columns.Item("AcctName").Cells.Item(0).Value));
                    }
                    catch (Exception ex)
                    {
                    }
                    Size Fieldsize = System.Windows.Forms.TextRenderer.MeasureText(EditText2.Value, new Font("Arial", 12.0f));
                    if (Fieldsize.Width <= 155) EditText2.Item.Width = 155;
                    else EditText2.Item.Width = Fieldsize.Width;
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }        

        private void ChooseFromList_Condition(string CFLID,string Alias,string CondVal,string Query, string Alias1, string CondVal1)
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
                if (Alias1 != "" && CondVal1!="")
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
                
                    
                    oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {
                //clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void ChooseFromList_TwoAlias_Condition(string CFLID, string Alias1, string CondVal1, string Alias2, string CondVal2)
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
                oCond.Alias = Alias1;// "Postable";
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CondVal1;// "Y";
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;
                oCond = oConds.Add();
                oCond.Alias = Alias2;
                oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCond.CondVal = CondVal2;

                oCFL.SetConditions(oConds);

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Item Events

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                if (Matrix0.VisualRowCount == 0)
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Line details Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; return;
                }
                if (EditText0.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("G/L Account Code is Missing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    BubbleEvent = false; return;
                }
                
                //for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                //{
                //    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("project").Cells.Item(i).Specific).String != "")
                //    {
                //        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("glaccc").Cells.Item(i).Specific).String == "")
                //        {
                //            clsModule.objaddon.objapplication.StatusBar.SetText("To G/L is Missing...On Line: " + i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //            BubbleEvent = false; return;
                //        }
                //        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("revamt").Cells.Item(i).Specific).String == "")
                //        {
                //            clsModule.objaddon.objapplication.StatusBar.SetText("Revenue Amount is Missing...On Line: " + i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //            BubbleEvent = false; return;
                //        }
                //        if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("invnum").Cells.Item(i).Specific).String == "")
                //        {
                //            clsModule.objaddon.objapplication.StatusBar.SetText("Invoice Num is Missing...On Line: " + i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //            BubbleEvent = false; return;
                //        }
                //    }
                //}

                RemoveLastrow(Matrix0, "project");

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_PressedAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                if (pVal.ActionSuccess == true & objform.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    clsModule.objaddon.objglobalmethods.LoadSeries(objform, odbdsHeader, "REVREC");
                    ((SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                    ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                    //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "invnum", "#");
                    Folder0.Item.Click();
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button2_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {                
                //strSQL = "Select TO_VARCHAR(LAST_DAY(TO_VARCHAR('" + ComboBox3.Selected.Value + "'||'" + ComboBox2.Selected.Value + "'||'01','yyyyMMdd')),'yyyyMMdd') from dummy";
                //strSQL = clsModule.objaddon.objglobalmethods.getSingleValue(strSQL);
                //strSQL = "Call \"ATPL_GetRevenueDetails\" ('"+ strSQL + "','" + ComboBox4.Selected.Value + "')";
                strSQL = "Call \"ATPL_GetRevenueDetails\" ((Select LAST_DAY(TO_VARCHAR('" + ComboBox3.Selected.Value + "'||'" + ComboBox2.Selected.Value + "'||'01','yyyyMMdd')) from dummy),'" + ComboBox4.Selected.Value.ToUpper() + "')";
                if (clsModule.objaddon.FormExist("REVDSEL")) return;
                FrmRevDataSelect activeform = new FrmRevDataSelect(objform, strSQL);
                activeform.Show();
            }
            catch (Exception ex)
            {

            }

        } //Get Data        

        private void Button2_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) { BubbleEvent = false; return; }
                if (Matrix0.VisualRowCount>0) if (clsModule.objaddon.objapplication.MessageBox("Do you want to re-load the details. Continue?", 2, "Yes", "No") != 1) { BubbleEvent = false; return; }
                if (ComboBox4.Value == "-")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Select Project Type...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    ComboBox4.Item.Click();
                    BubbleEvent = false;
                    return;
                }
                if (ComboBox3.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Select Year...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    ComboBox3.Item.Click();
                    BubbleEvent = false;
                    return;
                }
                if (ComboBox2.Value == "")
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Select Month...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    ComboBox2.Item.Click();
                    BubbleEvent = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #endregion

        #region Form Events

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("REVRECO", pVal.FormTypeCount);
            }
            catch (Exception)
            {
                //throw;
            }

        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Matrix0.AutoResizeColumns();
            }
            catch (Exception)
            {
                //throw;
            }

        }

        private void Form_DataAddBefore(ref SAPbouiCOM.BusinessObjectInfo pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (clsModule.objaddon.objapplication.MessageBox("You cannot change this document after you have added it. Continue?", 2, "Yes", "No") != 1) { BubbleEvent = false; return; }


                if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();
                
                strQuery = "Select \"BPLId\" from OBPL where \"Disabled\"='N' and \"MainBPL\"='Y'";
                strQuery = clsModule.objaddon.objglobalmethods.getSingleValue(strQuery);
                JournalID = JournalVoucher(objform.UniqueID, strQuery,out JournalID);
                if (JournalID != "") tranflag = true;
                else tranflag = false;
                if (tranflag == true)
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                else
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    clsModule.objaddon.objcompany.GetLastError(out errorCode, out strQuery);
                    clsModule.objaddon.objapplication.MessageBox("Journal Voucher: " + strQuery + "-" + errorCode, 0, "OK");
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Rolled back Journal Voucher..." + strQuery, SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    BubbleEvent = false;
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                BubbleEvent = false;
            }

        }

        private void Form_DataLoadAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                Matrix0.AutoResizeColumns();
                Size Fieldsize = System.Windows.Forms.TextRenderer.MeasureText(EditText2.Value, new Font("Arial", 12.0f));
                if (Fieldsize.Width <= 155) EditText2.Item.Width = 155;
                else EditText2.Item.Width = Fieldsize.Width;
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_REV_RECO");
                strSQL =clsModule.objaddon.objglobalmethods.getSingleValue( "Select T1.\"TransId\" from OBTF T0 Left join OJDT T1 on T0.\"BatchNum\"=T1.\"BatchNum\" where T0.\"BatchNum\"='"+ EditText5.Value + "' and ifnull(T0.\"BtfStatus\",'')='C'");
                if (strSQL != "") odbdsHeader.SetValue("U_TransId", 0, strSQL);
                if (objform.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                objform.Items.Item("1").Click();
                if (odbdsHeader.GetValue("U_TransId", 0) != "") { EditText6.Item.Visible = true; StaticText6.Item.Visible = true; LinkedButton2.Item.Visible = true; }
                else { EditText6.Item.Visible = false; StaticText6.Item.Visible = false; LinkedButton2.Item.Visible = false; }
                ////if (ComboBox1.Selected.Value == "O") { EditText6.Item.Visible = false; StaticText6.Item.Visible = false; LinkedButton2.Item.Visible = false; }
                ////else { EditText6.Item.Visible = true; StaticText6.Item.Visible = true; LinkedButton2.Item.Visible = true; }
                objform.EnableMenu("1282", true);
                //Matrix0.Item.Enabled = false;
                ////objform.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Form_DataAddAfter(ref SAPbouiCOM.BusinessObjectInfo pVal)
        {
            try
            {
                if (EditText5.Value == "" | pVal.ActionSuccess == false) return;
                int DocNum;
                odbdsHeader = objform.DataSources.DBDataSources.Item("@AT_REV_RECO");
                DocNum = objform.BusinessObject.GetNextSerialNumber(((SAPbouiCOM.ComboBox)objform.Items.Item("series").Specific).Selected.Value, "AT_REVREC");
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string[] Value = EditText5.Value.Split('\t');

                strSQL = "Update OBTF Set \"U_RevRecDE\"='" + odbdsHeader.GetValue("DocEntry", 0) + "',\"U_RevRecDN\"='" + DocNum + "' where \"BatchNum\"='" + Value[0] + "'";
                objRs.DoQuery(strSQL);
                strSQL = "Update \"@AT_REV_RECO\" Set \"Status\"='C' where \"DocEntry\"='" + odbdsHeader.GetValue("DocEntry", 0) + "'";
                objRs.DoQuery(strSQL);
                JournalID = "";
                
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        #endregion

        #region Matrix Events        
       
        private void Matrix0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.Row == 0) return;
                if (Matrix0.IsRowSelected(pVal.Row)==false) Matrix0.SelectRow(pVal.Row, true, false);
                else Matrix0.SelectRow(pVal.Row, false, false);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Matrix0_LinkPressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = false;
            try
            {
                SAPbouiCOM.Form objTempForm = null;
                switch (pVal.ColUID)
                {
                    case "project":
                        try
                        {
                            clsModule.objaddon.objapplication.Menus.Item("REVPRJMSTR").Activate();//Project Master
                            objTempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            objTempForm.Freeze(true);
                            objTempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                            objTempForm.Items.Item("tprjcode").Enabled = true;
                            ((SAPbouiCOM.EditText)objTempForm.Items.Item("tprjcode").Specific).String = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("project").Cells.Item(pVal.Row).Specific).String;
                            objTempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            objTempForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE;
                            objTempForm.Freeze(false);
                        }
                        catch (Exception ex)
                        {
                            objTempForm.Freeze(false);
                            objTempForm = null;
                        }
                        break;
                }
                
                //if (pVal.ColUID == "invnum")
                //{

                //    try
                //    {
                //        clsModule.objaddon.objapplication.Menus.Item("2053").Activate();//AR Invoice
                //        objTempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                //        objTempForm.Freeze(true);
                //        objTempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                //        objTempForm.Items.Item("8").Enabled = true;
                //        ((SAPbouiCOM.EditText)objTempForm.Items.Item("8").Specific).String = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("invnum").Cells.Item(pVal.Row).Specific).String;
                //        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select To_Varchar(\"DocDate\",'yyyyMMdd') \"DocDate\" from OINV where \"DocEntry\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(pVal.Row).Specific).String + "'");
                //        ((SAPbouiCOM.EditText)objTempForm.Items.Item("10").Specific).String = strSQL;
                //        objTempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //        objTempForm.Freeze(false);
                //    }
                //    catch (Exception ex)
                //    {
                //        objTempForm.Freeze(false);
                //        objTempForm = null;
                //    }

                //}
                //else if (pVal.ColUID == "sonum")
                //{
                //    try
                //    {
                //        clsModule.objaddon.objapplication.Menus.Item("2050").Activate();//Sales Order
                //        objTempForm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                //        objTempForm.Freeze(true);
                //        objTempForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;
                //        objTempForm.Items.Item("8").Enabled = true;
                //        ((SAPbouiCOM.EditText)objTempForm.Items.Item("8").Specific).String = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("invnum").Cells.Item(pVal.Row).Specific).String;
                //        strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select To_Varchar(\"DocDate\",'yyyyMMdd') \"DocDate\" from ORDR where \"DocEntry\"='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(pVal.Row).Specific).String + "'");
                //        ((SAPbouiCOM.EditText)objTempForm.Items.Item("10").Specific).String = strSQL;
                //        objTempForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //        objTempForm.Freeze(false);
                //    }
                //    catch (Exception ex)
                //    {
                //        objTempForm.Freeze(false);
                //        objTempForm = null;
                //    }
                //}
            }
            catch (Exception)
            {
                //throw;
            }

        }

        private void ComboBox0_ComboSelectAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (objform.Mode != SAPbouiCOM.BoFormMode.fm_ADD_MODE) return;
                odbdsHeader.SetValue("DocNum", 0, clsModule.objaddon.objglobalmethods.GetDocNum("AT_REVREC", Convert.ToInt32(ComboBox0.Selected.Value)));
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        } //Series

        private void LinkedButton1_PressedBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //try
            //{
            //    string[] Value = EditText5.Value.Split('\t');

            //    strSQL = clsModule.objaddon.objglobalmethods.getSingleValue("Select 1 as \"Status\" from OBTF where \"BatchNum\"='" + Value[0] + "' and ifnull(\"BtfStatus\",'')='O'");
            //    LinkedButton1.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_None;
            //    if (strSQL != "")
            //    {
            //        LinkedButton1.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_JournalVoucher;
            //        StaticText5.Caption = "Journal Voucher";
            //    }
            //    else
            //    {
            //        LinkedButton1.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_JournalPosting;
            //        StaticText5.Caption = "Journal Entry";
            //    }
               
            //}
            //catch (Exception ex)
            //{
            //    clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //}

        }

        #endregion

        #region Functions

        private bool JournalEntry(string FormUID, string Branch)
        {
            try
            {
                string TransId, Series;
                SAPbobsCOM.JournalEntries objjournalentry;
                string JEAmount;
                SAPbouiCOM.EditText oEdit;
                DateTime DocDate;
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                if (((SAPbouiCOM.EditText)objform.Items.Item("ttranid").Specific).String != "") return true;
                
                objjournalentry = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();

                oEdit =(SAPbouiCOM.EditText) objform.Items.Item("tposdate").Specific; // Posting Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                
                objjournalentry.ReferenceDate = DocDate; // Posting Date
                //oEdit = objform.Items.Item("121").Specific; // Due Date
                //DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                //objjournalentry.DueDate = DocDate;   // Due Date
                oEdit = (SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific; // Tax Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                objjournalentry.TaxDate = DocDate;   // Document Date
                objjournalentry.Memo = objform.Title + " - " + ComboBox4.Selected.Description + " - " + ComboBox2.Selected.Description + " - " + ComboBox3.Selected.Description; // Project type + Month + Year
                objjournalentry.UserFields.Fields.Item("U_RevRecDN").Value = EditText1.Value;
                //objjournalentry.UserFields.Fields.Item("U_RevRecDE").Value = "";
                //objjournalentry.Memo = objform.Items.Item("59").Specific.String;
                if (clsModule.objaddon.HANA)
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select \"LawsSet\" from CINF");
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' and \"BPLId\"='" + Branch + "'");
                }
                else
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select LawsSet from CINF");
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' and BPLId='" + Branch + "'");
                }
                if (Localization != "IN")              
                {
                    objjournalentry.AutoVAT = BoYesNoEnum.tNO;
                    objjournalentry.AutomaticWT = BoYesNoEnum.tNO;                    
                }
                if (!string.IsNullOrEmpty(Series)) objjournalentry.Series =Convert.ToInt32(Series);
               
                for (int AccRow = 1; AccRow <= Matrix0.VisualRowCount; AccRow++)
                {
                    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("glaccc").Cells.Item(AccRow).Specific).String != "" )
                    {
                        JEAmount = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("revcurmon").Cells.Item(AccRow).Specific).String;
                        objjournalentry.Lines.AccountCode = EditText0.Value;
                        objjournalentry.Lines.Debit = Convert.ToDouble(JEAmount);
                        if(Branch!="") objjournalentry.Lines.BPLID =Convert.ToInt32(Branch);
                        objjournalentry.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                        objjournalentry.Lines.Add();
                        objjournalentry.Lines.AccountCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("glaccc").Cells.Item(AccRow).Specific).String;
                        objjournalentry.Lines.Credit = Convert.ToDouble(JEAmount);
                        if (Branch != "") objjournalentry.Lines.BPLID = Convert.ToInt32(Branch);
                        objjournalentry.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                        objjournalentry.Lines.Add();
                    }
                }
                if (objjournalentry.Add() != 0)
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    clsModule.objaddon.objapplication.MessageBox("Journal Transaction: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 0, "OK");
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry);
                    return false;
                }
                else
                {
                    if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    TransId = clsModule.objaddon.objcompany.GetNewObjectKey();
                    ((SAPbouiCOM.EditText)objform.Items.Item("ttranid").Specific).String = TransId;
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Entry Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    return true;
                }
            }
            
            catch (Exception ex)
            {                
                if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                clsModule.objaddon.objapplication.SetStatusBarMessage("JE Posting Error: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + " : "+ ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;

            }
            
        }

        private string JournalVoucher(string FormUID, string Branch,out string JVNum)
        {
            JVNum = "";
            try
            {
                string TransId, Series;
                SAPbobsCOM.JournalVouchers journalVouchers  ;
                string JEAmount;
                SAPbouiCOM.EditText oEdit;
                DateTime DocDate;
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                
                //if (((SAPbouiCOM.EditText)objform.Items.Item("ttranid").Specific).String != "") return true;

                journalVouchers = (SAPbobsCOM.JournalVouchers)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);
                clsModule.objaddon.objapplication.StatusBar.SetText("Journal Voucher Creating Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);                

                oEdit = (SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific; // Posting Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);

                journalVouchers.JournalEntries.ReferenceDate = DocDate; // Posting Date
               
                oEdit = (SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific; // Tax Date
                DocDate = DateTime.ParseExact(oEdit.Value, "yyyyMMdd", DateTimeFormatInfo.InvariantInfo);
                journalVouchers.JournalEntries.TaxDate = DocDate;   // Document Date
                journalVouchers.JournalEntries.Memo = "POC Revenue for - " + ComboBox4.Selected.Description + " - " + ComboBox2.Selected.Description + " - " + ComboBox3.Selected.Description; // Project type + Month + Year  //Memo-POC Revenue for May 2023
                journalVouchers.JournalEntries.UserFields.Fields.Item("U_RevRecDN").Value = EditText1.Value;

                if (clsModule.objaddon.HANA)
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select \"LawsSet\" from CINF");
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 \"Series\" FROM NNM1 WHERE \"ObjectCode\"='30' and \"Indicator\"=(Select \"Indicator\" from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between \"F_RefDate\" and \"T_RefDate\") " + " and Ifnull(\"Locked\",'')='N' and \"BPLId\"='" + Branch + "'");
                }
                else
                {
                    Localization = clsModule.objaddon.objglobalmethods.getSingleValue("select LawsSet from CINF");
                    if (Branch == "") Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' ");
                    else Series = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1 Series FROM NNM1 WHERE ObjectCode='30' and Indicator=(Select Indicator from OFPR where '" + DocDate.ToString("yyyyMMdd") + "' Between F_RefDate and T_RefDate) " + " and Isnull(Locked,'')='N' and BPLId='" + Branch + "'");
                }
                if (Localization != "IN")
                {
                    journalVouchers.JournalEntries.AutoVAT = BoYesNoEnum.tNO;
                    journalVouchers.JournalEntries.AutomaticWT = BoYesNoEnum.tNO;
                }
                if (!string.IsNullOrEmpty(Series)) journalVouchers.JournalEntries.Series = Convert.ToInt32(Series);

                Matrix0.FlushToDataSource();
                for (int ContentRow = 0; ContentRow <= odbdsDetails.Size - 1; ContentRow++)
                {
                    if (odbdsDetails.GetValue("U_Project", ContentRow) != "" && Convert.ToDouble(odbdsDetails.GetValue("U_CurRevCost", ContentRow))!=0)
                    {
                        JEAmount = odbdsDetails.GetValue("U_CurRevCost", ContentRow);
                        journalVouchers.JournalEntries.Lines.AccountCode = EditText0.Value;
                        journalVouchers.JournalEntries.Lines.Debit = Convert.ToDouble(JEAmount);
                        if (Branch != "") journalVouchers.JournalEntries.Lines.BPLID = Convert.ToInt32(Branch);
                        journalVouchers.JournalEntries.Lines.ProjectCode = odbdsDetails.GetValue("U_Project", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode1", ContentRow)!="") journalVouchers.JournalEntries.Lines.CostingCode = odbdsDetails.GetValue("U_OcrCode1", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode2", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode2 = odbdsDetails.GetValue("U_OcrCode2", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode3", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode3 = odbdsDetails.GetValue("U_OcrCode3", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode4", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode4 = odbdsDetails.GetValue("U_OcrCode4", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode5", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode5 = odbdsDetails.GetValue("U_OcrCode5", ContentRow);

                        //journalVouchers.JournalEntries.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                        journalVouchers.JournalEntries.Lines.Add();


                        journalVouchers.JournalEntries.Lines.AccountCode = odbdsDetails.GetValue("U_RevAcc", ContentRow);
                        journalVouchers.JournalEntries.Lines.Credit = Convert.ToDouble(JEAmount);
                        if (Branch != "") journalVouchers.JournalEntries.Lines.BPLID = Convert.ToInt32(Branch);
                        journalVouchers.JournalEntries.Lines.ProjectCode = odbdsDetails.GetValue("U_Project", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode1", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode = odbdsDetails.GetValue("U_OcrCode1", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode2", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode2 = odbdsDetails.GetValue("U_OcrCode2", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode3", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode3 = odbdsDetails.GetValue("U_OcrCode3", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode4", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode4 = odbdsDetails.GetValue("U_OcrCode4", ContentRow);
                        if (odbdsDetails.GetValue("U_OcrCode5", ContentRow) != "") journalVouchers.JournalEntries.Lines.CostingCode5 = odbdsDetails.GetValue("U_OcrCode5", ContentRow);

                        //journalVouchers.JournalEntries.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                        journalVouchers.JournalEntries.Lines.Add();
                    }

                }

                
                //for (int AccRow = 1; AccRow <= Matrix0.VisualRowCount; AccRow++)
                //{
                //    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("project").Cells.Item(AccRow).Specific).String != "")
                //    {
                //        JEAmount = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("revcurmon").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.AccountCode = EditText0.Value;
                //        journalVouchers.JournalEntries.Lines.Debit = Convert.ToDouble(JEAmount);
                //        if (Branch != "") journalVouchers.JournalEntries.Lines.BPLID = Convert.ToInt32(Branch);
                //        journalVouchers.JournalEntries.Lines.ProjectCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("project").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.CostingCode = "";

                //        journalVouchers.JournalEntries.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.Add();


                //        journalVouchers.JournalEntries.Lines.AccountCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("glaccc").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.Credit = Convert.ToDouble(JEAmount);
                //        if (Branch != "") journalVouchers.JournalEntries.Lines.BPLID = Convert.ToInt32(Branch);
                //        journalVouchers.JournalEntries.Lines.ProjectCode = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("project").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.CostingCode = "";
                //        journalVouchers.JournalEntries.Lines.UserFields.Fields.Item("U_InvEntry").Value = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("inventry").Cells.Item(AccRow).Specific).String;
                //        journalVouchers.JournalEntries.Lines.Add();
                //    }
                //}

                journalVouchers.JournalEntries.Add();

                if (journalVouchers.Add() != 0)
                {                    
                    clsModule.objaddon.objapplication.MessageBox("Journal Transaction: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), 0, "OK");
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(journalVouchers);
                    return "";
                }
                else
                {                    
                    TransId = clsModule.objaddon.objcompany.GetNewObjectKey();
                    JVNum = TransId;                    
                    ((SAPbouiCOM.EditText)objform.Items.Item("tvochid").Specific).String = TransId.Split('\t')[0];
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Voucher Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                    return JVNum;
                }
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("JE Posting Error: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return "";

            }

        }

        private void RemoveLastrow(SAPbouiCOM.Matrix omatrix, string Columname_check)
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
        
        private void Load_Data()
        {
            try
            {
                for (int i = -5; i <= 1; i++)
                {
                    string date1 = DateTime.Now.AddYears(i).Year.ToString();
                    ComboBox3.ValidValues.Add(date1, date1);
                }
                strSQL = DateTime.Now.Year.ToString();
                //ComboBox3.Select(strSQL, SAPbouiCOM.BoSearchKey.psk_ByValue);
                string[] names = DateTimeFormatInfo.CurrentInfo.MonthNames;
                for (int i = 0; i < names.Length - 1; i++)
                {
                    if (Convert.ToString(i + 1).Length == 1) strSQL = "0" + (i + 1); else strSQL = Convert.ToString(i + 1);
                    ComboBox2.ValidValues.Add(strSQL, names[i]);
                }
                if (DateTime.Now.Month.ToString().Length == 1) strSQL = "0" + DateTime.Now.Month.ToString(); else strSQL = DateTime.Now.Month.ToString();
                //ComboBox2.Select(strSQL, SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        #endregion

        
    }
}
