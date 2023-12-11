using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM.Framework;
using RevenueRecognition.Common;
using System.Drawing;

namespace RevenueRecognition.Business_Objects
{
    [FormAttribute("REVDSEL", "Business_Objects/FrmRevDataSelect.b1f")]
    class FrmRevDataSelect : UserFormBase
    {
        public string Details_Query,strsql,colID;
        public static SAPbouiCOM.Form objform,oRevform;

        public FrmRevDataSelect(SAPbouiCOM.Form form, string Query)
        {
            try
            {
                objform.Freeze(true);
                Details_Query = Query;
                oRevform = form;
                objform.Left =(clsModule.objaddon.objapplication.Desktop.Width- form.MaxWidth)/2;//form.Left + 50;
                objform.Top = (clsModule.objaddon.objapplication.Desktop.Height - form.MaxHeight) / 2;// form.Top + 50;
                strsql = "Percentage Of Completion (" + ((SAPbouiCOM.ComboBox)form.Items.Item("cmonth").Specific).Selected.Description + " - " + ((SAPbouiCOM.ComboBox)form.Items.Item("cyear").Specific).Selected.Description + ") - Selection Criteria";
                objform.Title = strsql;
                Load_Grid(Query);
               clsModule.objaddon.objGlobalVariables.bModal = true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                objform.Freeze(false);
            }

            
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btnok").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("gdata").Specific));
            this.Grid0.DoubleClickBefore += new SAPbouiCOM._IGridEvents_DoubleClickBeforeEventHandler(this.Grid0_DoubleClickBefore);
            this.Grid0.ClickBefore += new SAPbouiCOM._IGridEvents_ClickBeforeEventHandler(this.Grid0_ClickBefore);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lfind").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("tfind").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.ResizeAfter += new ResizeAfterHandler(this.Form_ResizeAfter);

        }

        #region Fields

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Grid Grid0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;

        #endregion


        private void OnCustomInitialize()
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.GetForm("REVDSEL", 1);
                //objform.Settings.Enabled = true;
                objform.EnableMenu("4870", true);
                
            }
            catch (Exception ex)
            {

            }
        }

        private void Form_ResizeAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                Grid0.AutoResizeColumns();
            }
            catch (Exception ex)
            {
            }

        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.Row == -1) return;
                switch (pVal.ColUID)
                {
                    //case "Select":
                    //    if (Convert.ToString(Grid0.DataTable.GetValue("Select", Grid0.GetDataTableRowIndex(pVal.Row))) == "Y")
                    //    {
                    //        Grid0.Rows.SelectedRows.Add(pVal.Row);
                    //    }
                    //    else
                    //    {
                    //        Grid0.Rows.SelectedRows.Remove(pVal.Row);
                    //    }
                    //    break;
                   default :  
                        if (Grid0.Rows.IsSelected(pVal.Row) == true)
                        {
                            //Grid0.DataTable.SetValue("Select", pVal.Row, "N");
                            Grid0.Rows.SelectedRows.Remove(pVal.Row);
                        }
                        else
                        {
                            //Grid0.DataTable.SetValue("Select", pVal.Row, "Y");
                            Grid0.Rows.SelectedRows.Add(pVal.Row);
                        }                        
                        break;

                }
                
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Grid0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                //if (pVal.Row == -1) BubbleEvent = false;
                if (Convert.ToString(Grid0.DataTable.GetValue("Contract Number", Grid0.GetDataTableRowIndex(0)))=="") BubbleEvent = false;
                
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (Grid0.Rows.SelectedRows.Count == 0) { clsModule.objaddon.objapplication.StatusBar.SetText("Select a row...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); BubbleEvent = false;  return; }
                Load_SelectedDetails_ToRevenue("@AT_REV_RECO1");
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (pVal.ActionSuccess == true) objform.Close();

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Grid0_DoubleClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.Row == -1) { colID = pVal.ColUID; Grid0.Item.Click(); }
                else BubbleEvent = false;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            try
            {
                if (EditText0.Value == "") return;
                Grid0.Rows.SelectedRows.Clear();
                for (int i = 0; i < Grid0.DataTable.Rows.Count; i++)
                {
                    strsql = Grid0.DataTable.GetValue(colID, Grid0.GetDataTableRowIndex(i)).ToString().ToUpper();
                    if(Grid0.DataTable.GetValue(colID, Grid0.GetDataTableRowIndex(i)).ToString().ToUpper().Contains(EditText0.Value.ToUpper()) == true)
                    {
                        Grid0.Rows.SelectedRows.Add(Grid0.GetDataTableRowIndex(i));
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }

        private void Load_Grid(string Query)
        {
            try
            {
                objform.DataSources.DataTables.Item("DT_0").ExecuteQuery(Query);
                Grid0.DataTable = objform.DataSources.DataTables.Item("DT_0");
                //Grid0.Columns.Item(0).Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                Grid0.RowHeaders.TitleObject.Caption = "#";
                objform.Freeze(true);
                for (int i = 0; i <= Grid0.Columns.Count - 1; i++)
                {
                    Grid0.Columns.Item(i).TitleObject.Sortable = true;
                    if (i == 0) continue;                    
                    //Grid0.Columns.Item(i).Editable = false;                    
                    //Grid0.Columns.Item(i).BackColor = Color.Brown.ToArgb();
                     Grid0.Columns.Item(i).ForeColor = Color.Brown.ToArgb(); //DarkMagenta, Green
                    Grid0.Columns.Item(i).TextStyle = Convert.ToInt16(FontStyle.Bold);
                    
                }
                Grid0.CommonSetting.FixedColumnsCount = 7;
                Grid0.CommonSetting.EnableArrowKey = true;
                Grid0.AutoResizeColumns();
                objform.Freeze(false);
            }
            catch (Exception ex)
            {
                objform.Freeze(false);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void Load_SelectedDetails_ToRevenue(string LineDBSource)
        {
            try
            {
                SAPbouiCOM.Matrix RevMatrix;
                RevMatrix =(SAPbouiCOM.Matrix)oRevform.Items.Item("mtxcont").Specific;
                oRevform.DataSources.DBDataSources.Item(LineDBSource).Clear();
                int Row=1;

                for (int i = Grid0.Rows.SelectedRows.Count; i >= 1; i--)
                {
                    if (Convert.ToDouble(Grid0.DataTable.GetValue("Revenue Current Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))) == 0) continue;
                        var withBlock = oRevform.DataSources.DBDataSources.Item(LineDBSource);
                        withBlock.InsertRecord(0);
                        withBlock.SetValue("LineId", 0,Convert.ToString(Row));
                        withBlock.SetValue("U_Project", 0, Convert.ToString(Grid0.DataTable.GetValue("Contract Number", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_ComRes", 0, Convert.ToString(Grid0.DataTable.GetValue("Com/Res", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_Status", 0, Convert.ToString(Grid0.DataTable.GetValue("Status", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_ConDesc", 0, Convert.ToString(Grid0.DataTable.GetValue("Contract Description", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_ProjValue", 0, Convert.ToString(Grid0.DataTable.GetValue("Original Contract Value", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_VarValue", 0, Convert.ToString(Grid0.DataTable.GetValue("Variation Omission", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_RevCntVal", 0, Convert.ToString(Grid0.DataTable.GetValue("Revised Contract Value", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_PrevRev", 0, Convert.ToString(Grid0.DataTable.GetValue("Previous Revenue", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_PrevCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Previous Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurMonCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Current Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_MatCostTill", 0, Convert.ToString(Grid0.DataTable.GetValue("Material Cost Till Last Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_EstCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Estimated Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_EstProfit", 0, Convert.ToString(Grid0.DataTable.GetValue("Estimated Profit", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_EstPercent", 0, Convert.ToString(Grid0.DataTable.GetValue("% of estimated MU", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CostTill", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Till Last Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurMatCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Material Cost Current Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurLabCost", 0, Convert.ToString(Grid0.DataTable.GetValue("YAC Labour Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurCLabCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Casual Labour Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurSubCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Sub Contract Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_TotCostTill", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Till Last Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_TotCost", 0, Convert.ToString(Grid0.DataTable.GetValue("TOTAL COST - PTD", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_EstCostCnt", 0, Convert.ToString(Grid0.DataTable.GetValue("Estimated cost to complete the contract", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_TotCostCnt", 0, Convert.ToString(Grid0.DataTable.GetValue("Total Contract Cost", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_PerComp", 0, Convert.ToString(Grid0.DataTable.GetValue("Percentage of Completion", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_RevTill", 0, Convert.ToString(Grid0.DataTable.GetValue("Revenue Till Last Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurRevCost", 0, Convert.ToString(Grid0.DataTable.GetValue("Revenue Current Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_Revenue", 0, Convert.ToString(Grid0.DataTable.GetValue("Revenue - PTD", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurProfit", 0, Convert.ToString(Grid0.DataTable.GetValue("Profit", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_RevtoComp", 0, Convert.ToString(Grid0.DataTable.GetValue("Revenue to Complete", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_EstMargin", 0, Convert.ToString(Grid0.DataTable.GetValue("Estimated Margin to Complete", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_MartoComp", 0, Convert.ToString(Grid0.DataTable.GetValue("% Margin to complete", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_InvValue", 0, Convert.ToString(Grid0.DataTable.GetValue("Invoiced - PTD", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_CurInv", 0, Convert.ToString(Grid0.DataTable.GetValue("Invoiced - Current Month", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_TotInv", 0, Convert.ToString(Grid0.DataTable.GetValue("Total Invoiced", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_DueFrmCus", 0, Convert.ToString(Grid0.DataTable.GetValue("Due From Customer", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_DueToCus", 0, Convert.ToString(Grid0.DataTable.GetValue("Due To Customer", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_AdvFrmCus", 0, Convert.ToString(Grid0.DataTable.GetValue("Advance Amount", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));

                        withBlock.SetValue("U_OcrCode1", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Center 1", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_OcrCode2", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Center 2", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_OcrCode3", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Center 3", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_OcrCode4", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Center 4", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_OcrCode5", 0, Convert.ToString(Grid0.DataTable.GetValue("Cost Center 5", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_RevAcc", 0, Convert.ToString(Grid0.DataTable.GetValue("Revenue Account", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        withBlock.SetValue("U_ExpAcc", 0, Convert.ToString(Grid0.DataTable.GetValue("Expense Account", Grid0.GetDataTableRowIndex(Grid0.Rows.SelectedRows.Item(i - 1, SAPbouiCOM.BoOrderType.ot_SelectionOrder)))));
                        Row += 1;
                }
                RevMatrix.LoadFromDataSourceEx();
                RevMatrix.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Load_SelectedDetails: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                GC.Collect();
            }
        }


    }
}
