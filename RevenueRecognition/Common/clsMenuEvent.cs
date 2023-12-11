using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevenueRecognition.Common
{
    class clsMenuEvent
    {     

        SAPbouiCOM.Form objform;
        string strsql;
        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (clsModule. objaddon.objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "-392":
                    case "-393":
                    case "392":
                    case "393":
                        {
                            // Default_Sample_MenuEvent(pVal, BubbleEvent)
                            if (pVal.BeforeAction == true)
                                return;
                            objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            Default_Sample_MenuEvent(pVal, BubbleEvent);

                            break;
                        }
                    case "REVREC":
                        RevenueRecognition_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                    case "REVPRJMSTR":
                        ProjectMaster_MenuEvent(ref pVal, ref BubbleEvent);
                        break;
                }
            } 
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }

                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find
                            {
                                oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                oUDFForm.Items.Item("U_RevRecDE").Enabled = true;
                                break;
                            }
                        case "1287":
                            {
                                if (oUDFForm.Items.Item("U_RevRecDN").Enabled == false|| oUDFForm.Items.Item("U_RevRecDE").Enabled == false)
                                {
                                    oUDFForm.Items.Item("U_RevRecDN").Enabled = true;
                                    oUDFForm.Items.Item("U_RevRecDE").Enabled = true;
                                }
                                ((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_RevRecDN").Specific).String = "";
                                ((SAPbouiCOM.EditText)oUDFForm.Items.Item("U_RevRecDE").Specific).String = "";
                                break;
                            }
                        default:
                            {
                                oUDFForm.Items.Item("U_RevRecDN").Enabled = false;
                                oUDFForm.Items.Item("U_RevRecDE").Enabled = false;
                                break;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                // objaddon.objapplication.SetStatusBarMessage("Error in Standart Menu Event" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
            }
        }

        private void RevenueRecognition_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                SAPbobsCOM.Recordset objRs;
                SAPbouiCOM.DBDataSource DBSource;
                SAPbouiCOM.Matrix Matrix0;
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                DBSource = objform.DataSources.DBDataSources.Item("@AT_REV_RECO");
                Matrix0 = (SAPbouiCOM.Matrix)objform.Items.Item("mtxcont").Specific;
                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1284": //Cancel
                            if (clsModule.objaddon.objapplication.MessageBox("Cancelling of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            else
                            {
                                if (((SAPbouiCOM.EditText)objform.Items.Item("tvochid").Specific).String != "")
                                {
                                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                                    //if( Remove_JournalVoucher(objform.UniqueID, Convert.ToInt32(((SAPbouiCOM.EditText)objform.Items.Item("tvochid").Specific).String)) == false) 
                                    if (clsModule.objaddon.HANA)
                                    {
                                        strsql = "Update \"@AT_REV_RECO\" Set \"U_VoucherID\"=null Where \"DocEntry\"=" + objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry", 0) + " ";
                                    }
                                    else
                                    {
                                        strsql = "Update @AT_REV_RECO Set U_VoucherID=null Where DocEntry=" + objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry", 0) + " ";
                                    }
                                    objRs.DoQuery(strsql);
                                    //BubbleEvent = false;
                                }
                            }
                            break;
                        case "1286":
                            {
                                //clsModule.objaddon.objapplication.SetStatusBarMessage("Remove or Cancel is not allowed ", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                                //BubbleEvent = false;
                                //return;
                                break;
                            }
                        case "1293":
                            if (Matrix0.VisualRowCount == 1) BubbleEvent = false;
                            break;
                    }
                }
                else
                {                    
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                            
                                objform.Items.Item("tdocnum").Enabled = true;
                                objform.Items.Item("tglc").Enabled = true;
                                objform.Items.Item("tgln").Enabled = true;
                                objform.Items.Item("series").Enabled = true;
                                objform.Items.Item("tposdate").Enabled = true;                            
                                objform.Items.Item("mtxcont").Enabled = false;
                            objform.Items.Item("tvochid").Enabled = true;
                            objform.Items.Item("tjeid").Enabled = true;
                            objform.EnableMenu("1282", true);
                            objform.ActiveItem = "tdocnum";
                            break;
                            
                        case "1282"://Add Mode                            
                                clsModule.objaddon.objglobalmethods.LoadSeries(objform, DBSource, "AT_REVREC");
                                ((SAPbouiCOM.EditText)objform.Items.Item("tposdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                                ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                            ((SAPbouiCOM.ComboBox)objform.Items.Item("cprojtype").Specific).Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                //clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "invnum", "#");
                            objform.EnableMenu("1282", false);
                            break;
                        case "1293"://Delete Row
                            //DeleteRow(Matrix0, "@REV_RECO1");
                            break;
                                              
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void ProjectMaster_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {
                if (clsModule.objaddon.HANA == true)
                    strsql = "Select \"USERID\" from OUSR Where \"USER_CODE\"='" + clsModule.objaddon.objcompany.UserName + "'";
                else
                    strsql = "Select USERID from OUSR Where USER_CODE='" + clsModule.objaddon.objcompany.UserName + "'";
                strsql = clsModule.objaddon.objglobalmethods.getSingleValue(strsql);

                //clsModule.objaddon.objglobalmethods.Update_UserFormSettings_UDF(objform, "-REVPRJMSTR", Convert.ToInt32(strsql)); //REVPRJMSTR

                SAPbouiCOM.DBDataSource odbdsContent, odbdsBoqItem, odbdsBoqLabour;
                SAPbouiCOM.Matrix matContent,matBoqItem,matBoqLabour;
                
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                //Content Matrix
                odbdsContent = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR1");
                matContent = (SAPbouiCOM.Matrix)objform.Items.Item("mtxcont").Specific;
                //BOQ Item Matrix
                odbdsBoqItem = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR3");
                matBoqItem = (SAPbouiCOM.Matrix)objform.Items.Item("mboqitem").Specific;
                //BOQ Labour Matrix
                odbdsBoqLabour = objform.DataSources.DBDataSources.Item("@AT_PROJMSTR4");
                matBoqLabour = (SAPbouiCOM.Matrix)objform.Items.Item("mboqlab").Specific;          
                

                if (pval.BeforeAction == true)
                {
                    switch (pval.MenuUID)
                    {
                        case "1283":
                            if (clsModule.objaddon.objapplication.MessageBox("Removing of an entry cannot be reversed. Do you want to continue?", 2, "Yes", "No") != 1)
                            {
                                BubbleEvent = false;
                            }
                            break;
                        case "1293":
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fldrcont").Specific).Selected == true)
                            {
                                if (matContent.VisualRowCount == 1) { BubbleEvent = false; return; }
                            }

                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fboqitem").Specific).Selected == true)
                            {
                                if (matBoqItem.VisualRowCount == 1) { BubbleEvent = false; return; }
                            }

                            else if (((SAPbouiCOM.Folder)objform.Items.Item("fboqlab").Specific).Selected == true)
                            {
                                if (matBoqLabour.VisualRowCount == 1) { BubbleEvent = false; return; }
                            }
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fldrcont").Specific).Selected == true)
                            {
                                if (clsModule.objaddon.objapplication.MessageBox("Do you want to delete the line?", 2, "Yes", "No") != 1)
                                {
                                    BubbleEvent = false; return;
                                }
                                try
                                {
                                    bool soFlag = false;
                                    if (clsModule.objaddon.objGlobalVariables.contentMatCurRow == -1) return;
                                    matBoqItem.FlushToDataSource();
                                    for (int BOQItemRow = 0; BOQItemRow < odbdsBoqItem.Size - 1; BOQItemRow++) //BOQ Item 
                                    {
                                        if (odbdsBoqItem.GetValue("U_SOEntry", BOQItemRow) == odbdsContent.GetValue("U_SOEntry", clsModule.objaddon.objGlobalVariables.contentMatCurRow)) soFlag = true;
                                    }
                                    matBoqLabour.FlushToDataSource();
                                    for (int BOQLabourRow = 0; BOQLabourRow < odbdsBoqLabour.Size - 1; BOQLabourRow++) //BOQ Labour
                                    {
                                        if (odbdsBoqLabour.GetValue("U_SOEntry", BOQLabourRow) == odbdsContent.GetValue("U_SOEntry", clsModule.objaddon.objGlobalVariables.contentMatCurRow)) soFlag = true;
                                    }
                                    if (soFlag == true) { clsModule.objaddon.objapplication.StatusBar.SetText("Sales Order found in BOQ Lines...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error); BubbleEvent = false; clsModule.objaddon.objGlobalVariables.contentMatCurRow = -1; }
                                }
                                catch (Exception ex)
                                {

                                }
                            }
                            
                            
                            break;
                    }
                }
                else
                {
                    switch (pval.MenuUID)
                    {
                        case "1281": // Find Mode                            
                            //objform.Items.Item("tdocnum").Enabled = true;                            
                            break;
                        case "1293"://Delete Row
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fldrcont").Specific).Selected == true)  DeleteRow(matContent, "@AT_PROJMSTR1");
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fboqitem").Specific).Selected == true) DeleteRow(matBoqItem, "@AT_PROJMSTR3");
                            if (((SAPbouiCOM.Folder)objform.Items.Item("fboqlab").Specific).Selected == true) DeleteRow(matBoqLabour, "@AT_PROJMSTR4");
                            matContent.Columns.Item("netadd").Cells.Item(1).Click();
                            clsModule.objaddon.objapplication.SendKeys("{TAB}");
                            break;
                        case "1282"://Add Mode                            
                            //((SAPbouiCOM.EditText)objform.Items.Item("tdate").Specific).String = "A";//DateTime.Now.Date.ToString("dd/MM/yy");
                            ((SAPbouiCOM.EditText)objform.Items.Item("trem").Specific).String = "Created By " + clsModule.objaddon.objcompany.UserName + " on " + DateTime.Now.ToString("dd/MMM/yyyy HH:mm:ss");
                            clsModule.objaddon.objglobalmethods.Matrix_Addrow(matContent, "soentry", "#");
                            strsql = "Select * from \"@AT_PROJTYPE\"";
                            clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("cprotype").Specific), strsql, new[] { "-,-" });
                            strsql = "Select * from \"@AT_TYPE\"";
                            clsModule.objaddon.objglobalmethods.Load_Combo(objform.UniqueID, ((SAPbouiCOM.ComboBox)objform.Items.Item("ccomres").Specific), strsql, new[] { "-,-" });
                            break;                       

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public void DeleteRow(SAPbouiCOM.Matrix objMatrix, string TableName)
        {
            try
            {
                SAPbouiCOM.DBDataSource DBSource;
                // objMatrix = objform.Items.Item("20").Specific
                objMatrix.FlushToDataSource();
                DBSource = objform.DataSources.DBDataSources.Item(TableName); 
                for (int i = 1, loopTo = objMatrix.VisualRowCount; i <= loopTo; i++)
                {
                    objMatrix.GetLineData(i);
                    DBSource.Offset = i - 1;
                    DBSource.SetValue("LineId", DBSource.Offset, Convert.ToString(i));
                    objMatrix.SetLineData(i);
                    objMatrix.FlushToDataSource();
                }
                DBSource.RemoveRecord(DBSource.Size - 1);
                objMatrix.LoadFromDataSource();
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("DeleteRow  Method Failed:" + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
            finally
            {
            }
        }

        private bool Cancelling_IntBranch_RecoJournalEntry(string FormUID, string JETransId)
        {            
                string TransId;
                SAPbouiCOM.Matrix objmatrix;
                SAPbobsCOM.JournalEntries objjournalentry;
                if (string.IsNullOrEmpty(JETransId))
                    return true;
                SAPbobsCOM.Recordset objRs;
                string strSQL;
                try
                {
                    objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                    objmatrix =(SAPbouiCOM.Matrix) objform.Items.Item("mtxcont").Specific;
                    objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                    string GetStatus = clsModule.objaddon.objglobalmethods.getSingleValue("select distinct 1 as \"Status\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                    if (GetStatus == "1")
                    {
                        TransId = clsModule.objaddon.objglobalmethods.getSingleValue("select \"TransId\" from OJDT where \"StornoToTr\"=" + JETransId + "");
                        ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String= TransId;
                    //return true;
                    }
                    strSQL = "Select T0.\"Series\",T0.\"TaxDate\",T0.\"DueDate\",T0.\"RefDate\",T0.\"Ref1\",T0.\"Ref2\",T0.\"Memo\",T1.\"Account\",T1.\"Credit\",T1.\"Debit\",T1.\"BPLId\",T0.\"U_RevRecDN\",T0.\"U_RevRecDE\",T1.\"U_InvEntry\",";
                    strSQL += "\n (Select \"CardCode\" from OCRD where \"CardCode\"=T1.\"ShortName\") as \"BPCode\"";
                    strSQL += "\n from OJDT T0 join JDT1 T1 ON T0.\"TransId\"=T1.\"TransId\" where  T1.\"TransId\"='" + JETransId + "' order by T1.\"Line_ID\"";
                    objRs.DoQuery(strSQL);
                    if (objRs.RecordCount == 0)
                        return true;
                    if (!clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.StartTransaction();
                    objjournalentry = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Reversing Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objjournalentry.TaxDate = Convert.ToDateTime(objRs.Fields.Item("TaxDate").Value); // objJEHeader.GetValue("TaxDate", 0)
                    objjournalentry.DueDate = Convert.ToDateTime(objRs.Fields.Item("DueDate").Value); // objJEHeader.GetValue("DueDate", 0)
                    objjournalentry.ReferenceDate = Convert.ToDateTime(objRs.Fields.Item("RefDate").Value); // objJEHeader.GetValue("RefDate", 0)
                    objjournalentry.Reference = Convert.ToString(objRs.Fields.Item("Ref1").Value); // objJEHeader.GetValue("Ref1", 0)
                    objjournalentry.Reference2 = Convert.ToString(objRs.Fields.Item("Ref2").Value); // objJEHeader.GetValue("Ref2", 0)
                    objjournalentry.Reference3 = DateTime.Now.ToString();
                    objjournalentry.Memo = Convert.ToString(objRs.Fields.Item("Memo").Value) + "(Reversal) - " + JETransId; // objJEHeader.GetValue("Memo", 0) & " (Reversal) - " & Trim(JETransId)
                    objjournalentry.Series = Convert.ToInt32(objRs.Fields.Item("Series").Value); // objJEHeader.GetValue("Series", 0)
                    objjournalentry.UserFields.Fields.Item("U_RevRecDN").Value = Convert.ToString(objRs.Fields.Item("U_RevRecDN").Value);
                    objjournalentry.UserFields.Fields.Item("U_RevRecDE").Value = Convert.ToString(objRs.Fields.Item("U_RevRecDE").Value);
           
                for (int AccRow = 0; AccRow < objRs.RecordCount ; AccRow++)
                    {
                        if (Convert.ToString(objRs.Fields.Item("BPCode").Value) != "")
                            objjournalentry.Lines.ShortName = Convert.ToString(objRs.Fields.Item("BPCode").Value);
                        else
                            objjournalentry.Lines.AccountCode = Convert.ToString(objRs.Fields.Item("Account").Value);
                        if (Convert.ToDouble(objRs.Fields.Item("Credit").Value) != 0)
                            objjournalentry.Lines.Debit = Convert.ToDouble(objRs.Fields.Item("Credit").Value);
                        else
                            objjournalentry.Lines.Credit = Convert.ToDouble(objRs.Fields.Item("Debit").Value);
                        if(Convert.ToString(objRs.Fields.Item("BPLId").Value)!="") objjournalentry.Lines.BPLID = Convert.ToInt32(objRs.Fields.Item("BPLId").Value);
                        objjournalentry.Lines.UserFields.Fields.Item("U_InvEntry").Value =Convert.ToString( objRs.Fields.Item("U_InvEntry").Value); // Branch
                        objjournalentry.Lines.Add();
                        objRs.MoveNext();
                    }

                    if (objjournalentry.Add() != 0)
                    {
                        if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Reverse: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(objjournalentry);
                        return false;
                    }
                    // 
                    else
                    {
                      if (clsModule.objaddon.objcompany.InTransaction) clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                      TransId = clsModule.objaddon.objcompany.GetNewObjectKey();                
                        
                     ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String = TransId;
                     objRs.DoQuery("Update OJDT set \"StornoToTr\"=" + JETransId + " where \"TransId\"=" + TransId + "");
                    ((SAPbouiCOM.ComboBox)objform.Items.Item("cstatus").Specific).Select("C", SAPbouiCOM.BoSearchKey.psk_ByValue);
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    objform.Items.Item("1").Click();
                    objform.Items.Item("trvtran").Visible = true;
                    objform.Items.Item("lrvtran").Visible = true;
                    objform.Items.Item("lkrvtran").Visible = true;
                    objmatrix.Item.Enabled = false;
                    clsModule.objaddon.objapplication.StatusBar.SetText("Journal Entry Reversed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    return true;
                    }

                    //if (ErrorFlag)
                    //{
                    //    ((SAPbouiCOM.EditText)objform.Items.Item("trvtran").Specific).String = "";
                    //}
                    //else
                    //{
                    //    
                    //    clsModule.objaddon.objapplication.StatusBar.SetText("Transactions Cancelled Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    //    return true;
                    //}
                }
                catch (Exception ex)
                {
                    if (clsModule.objaddon.objcompany.InTransaction)  clsModule.objaddon.objcompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    clsModule.objaddon.objapplication.SetStatusBarMessage("Transaction Cancelling Error " + clsModule.objaddon.objcompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                    return false;
                }
            

        }

        private bool Remove_JournalVoucher(string FormUID, int JVoucherId)
        {
            SAPbobsCOM.JournalVouchers journalVoucher;
            if (string.IsNullOrEmpty(Convert.ToString(JVoucherId)))
                return true;
            SAPbobsCOM.Recordset objRs;
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.Item(FormUID);
                objRs = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset);
                string GetStatus = clsModule.objaddon.objglobalmethods.getSingleValue("select distinct 1 as \"Status\" from OBTF where \"BatchNum\"=" + JVoucherId + "");
                if (GetStatus == "") return true;               
               
                journalVoucher = (SAPbobsCOM.JournalVouchers)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalVouchers);

                clsModule.objaddon.objapplication.StatusBar.SetText("Journal Voucher Removing Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                if (journalVoucher.JournalEntries.GetByKey(JVoucherId))
                {
                    if (journalVoucher.JournalEntries.Remove() != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Remove: " + clsModule.objaddon.objcompany.GetLastErrorDescription() + "-" + clsModule.objaddon.objcompany.GetLastErrorCode(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(journalVoucher);
                        return false;
                    }
                    else
                    {
                        clsModule.objaddon.objapplication.StatusBar.SetText("Journal Voucher Removed Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        if (clsModule.objaddon.HANA)
                        {
                            strsql = "Update \"@AT_REV_RECO\" Set \"U_VoucherID\"=null Where \"DocEntry\"="+ objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry",0) + " ";
                        }
                        else
                        {
                            strsql = "Update @AT_REV_RECO Set U_VoucherID=null Where DocEntry=" + objform.DataSources.DBDataSources.Item("@AT_REV_RECO").GetValue("DocEntry", 0) + " ";
                        }
                        objRs.DoQuery(strsql);
                        return true;
                    }
                }
                return true;

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.SetStatusBarMessage("Journal Voucher Removing Error " + clsModule.objaddon.objcompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                return false;
            }


        }

    }
}
