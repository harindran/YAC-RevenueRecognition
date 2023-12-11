using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevenueRecognition.Common
{
    class clsTable
    {        
        public void FieldCreation()
        {
            AddTables("AT_PROJTYPE", "Project Type", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            AddTables("AT_TYPE", "COM RES Type", SAPbobsCOM.BoUTBTableType.bott_NoObject);
            Project_Master();
            Revenue_Recognition();
            AddFields("OPRJ", "RevAcc", "Revenue Account", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulChartOfAccounts);
            AddFields("OPRJ", "ExpAcc", "Expense Account", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulChartOfAccounts);
            AddFields("OPRJ", "PrjType", "Project Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "AT_PROJTYPE", SAPbobsCOM.BoYesNoEnum.tNO, "", false,new[] { ""});
            AddFields("OJDT", "RevRecDN", "Rev RecoNum", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("OJDT", "RevRecDE", "Rev RecoEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);            
        }
        #region Master Data Creation

        public void Project_Master()
        {
            AddFields("ORDR", "EngCode", "Engineer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulEmployeesInfo, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "" });//Engineer code in Sales Order

            AddTables("AT_PROJMSTR", "Project Master Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddTables("AT_PROJMSTR1", "Project Master Contents", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR2", "Project Master Attachments", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR3", "BOQ Item Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            AddTables("AT_PROJMSTR4", "BOQ Labour Details", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);
            
            //Header Table
            AddFields("@AT_PROJMSTR", "PrjCode", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR", "PrjName", "Project Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            //AddFields("@AT_PROJMSTR", "SONo", "SalesOrder Doc Num", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            //AddFields("@AT_PROJMSTR", "SOEntry", "SalesOrder Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            //AddFields("@AT_PROJMSTR", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            //AddFields("@AT_PROJMSTR", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PROJMSTR", "Date", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date);
         
            AddFields("@AT_PROJMSTR", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 5, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "O", false, new[] {"O,Open","W,Work In Progress","H,Hold","C,Completed" });
            //AddFields("@AT_PROJMSTR", "BusiCenter", "Business Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            //AddFields("@AT_PROJMSTR", "SalesEmp", "Sales Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR", "ProjType", "Project Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "" });//AT_PROJTYPE
            AddFields("@AT_PROJMSTR", "Type", "Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "" });//AT_TYPE
            AddFields("@AT_PROJMSTR", "ProjValue", "Initial Project Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR", "EstProjValue", "Estimated Project Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR", "PrevCost", "Previous Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR", "PrevRevenue", "Previous Revenue Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PROJMSTR", "ConDesc", "Contract Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PROJMSTR", "POC", "POC", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone,"", SAPbobsCOM.BoYesNoEnum.tNO,"N",true, new[] { "" });


            //Content Table
            AddFields("@AT_PROJMSTR1", "SONo", "SalesOrder Doc Num", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR1", "SOEntry", "SalesOrder Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR1", "NetAdd", "Net Addition", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR1", "NetDed", "Net Deduction", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR1", "EstValue", "Estimated Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR1", "CardCode", "Card Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@AT_PROJMSTR1", "CardName", "Card Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            //AddFields("@AT_PROJMSTR1", "BusiCenter", "Business Center", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            //AddFields("@AT_PROJMSTR1", "SalesEmp", "Sales Employee", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR1", "Origin", "SO Origin", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone,"", SAPbobsCOM.BoYesNoEnum.tNO,"",true, new[] { "" });
            AddFields("@AT_PROJMSTR1", "EngCode", "Engineer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PROJMSTR1", "EngName", "Engineer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
            AddFields("@AT_PROJMSTR1", "SlpCode", "Sales EmpCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PROJMSTR1", "SlpName", "Sales EmpName", SAPbobsCOM.BoFieldTypes.db_Alpha, 150);
            AddFields("@AT_PROJMSTR1", "IProjValue", "Initial Project Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR1", "Advance", "Advance Amount", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);

            //Attachment Table
            AddFields("@AT_PROJMSTR2", "TrgtPath", "Target Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PROJMSTR2", "SrcPath", "Source Path", SAPbobsCOM.BoFieldTypes.db_Memo, 200);
            AddFields("@AT_PROJMSTR2", "Date", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@AT_PROJMSTR2", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR2", "FileExt", "File Extension", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR2", "FreeText", "Free Text", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            //BOQ Item Table
            AddFields("@AT_PROJMSTR3", "ItemCode", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@AT_PROJMSTR3", "ItemName", "Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 200);
            AddFields("@AT_PROJMSTR3", "SONo", "SalesOrder Doc Num", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "SOEntry", "SalesOrder Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity);
            AddFields("@AT_PROJMSTR3", "UnitPrice", "Unit Price", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Price);
            AddFields("@AT_PROJMSTR3", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR3", "Project", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR3", "Uom", "Unit of Measurement", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR3", "OcrCode", "Cost Center 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR3", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);

            //BOQ Laour Table
            AddFields("@AT_PROJMSTR4", "AcctCode", "Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR4", "AcctName", "Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PROJMSTR4", "SONo", "SalesOrder Doc Num", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "SOEntry", "SalesOrder Doc Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "Total", "Total", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@AT_PROJMSTR4", "Project", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR4", "OcrCode", "Cost Center 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_PROJMSTR4", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_PROJMSTR4", "Cosglc", "Cost of Sales Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 20);
            AddFields("@AT_PROJMSTR4", "Cosgln", "Cost of Sales Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_PROJMSTR4", "LabType", "Labour Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "-", false, new[] { "-,-","CL,Casual Labour", "YL,YAC Labour" });

            AddUDO("AT_PROJMASTER", "Revenue Project Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, "AT_PROJMSTR", new[] { "AT_PROJMSTR1", "AT_PROJMSTR2", "AT_PROJMSTR3", "AT_PROJMSTR4" }, new[] { "Code", "Name" },false, true, false);
        }

        #endregion

        #region Document Data Creation

        public void Revenue_Recognition()
        {
            AddTables("AT_REV_RECO", "Revenue Recognition Header", SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTables("AT_REV_RECO1", "Revenue Recognition Lines", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);

            AddFields("@AT_REV_RECO", "GLCode", "GL Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@AT_REV_RECO", "GLName", "GL Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@AT_REV_RECO", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date);            
            AddFields("@AT_REV_RECO", "Month", "Month Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_REV_RECO", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_REV_RECO", "TransId", "JE TransId", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_REV_RECO", "VoucherID", "Journal Voucher ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);
            AddFields("@AT_REV_RECO", "ProjType", "Project Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulNone, "", SAPbobsCOM.BoYesNoEnum.tNO, "", false, new[] { "" });//PROJTYPE

            AddFields("@AT_REV_RECO1", "Project", "Project", SAPbobsCOM.BoFieldTypes.db_Alpha, 20); //Contract Number
            AddFields("@AT_REV_RECO1", "ComRes", "Com / Res", SAPbobsCOM.BoFieldTypes.db_Alpha, 30);//Com / Res
            AddFields("@AT_REV_RECO1", "Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 20); // Status
            AddFields("@AT_REV_RECO1", "ConDesc", "Contract Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);//Contract Description
            AddFields("@AT_REV_RECO1", "ProjValue", "Original contract Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Original contract Value
            AddFields("@AT_REV_RECO1", "VarValue", "Variation / (omision)", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Variation / (omision)
            AddFields("@AT_REV_RECO1", "RevCntVal", "Revised Contract Value", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Revised Contract Value
            AddFields("@AT_REV_RECO1", "PrevRev", "Previous Revenue", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Previous Revenue
            AddFields("@AT_REV_RECO1", "PrevCost", "Previous Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Previous Cost
            AddFields("@AT_REV_RECO1", "CurMonCost", "Cost Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Cost Current Month
            AddFields("@AT_REV_RECO1", "MatCostTill", "Material Cost Till Last Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Material Cost Till Last Month
            AddFields("@AT_REV_RECO1", "EstCost", "Estimated Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Estimated Cost
            AddFields("@AT_REV_RECO1", "EstProfit", "Estimated Profit", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Estimated Profit
            AddFields("@AT_REV_RECO1", "EstPercent", "% of estimated MU", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage); //% of estimated MU
            AddFields("@AT_REV_RECO1", "CostTill", "Cost Till Last Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Cost Till Last Month
            AddFields("@AT_REV_RECO1", "CurMatCost", "Material Cost Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Material Cost Current Month
            AddFields("@AT_REV_RECO1", "CurLabCost", "YAC Labor Cost Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //YAC Labor Cost Current Month
            AddFields("@AT_REV_RECO1", "CurCLabCost", "Casual Labor Cost Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Casual Labor Cost Current Month
            AddFields("@AT_REV_RECO1", "CurSubCost", "Sub Contract Cost Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Sub Contract Cost Current Month
            AddFields("@AT_REV_RECO1", "TotCostTill", "Total Cost Till Last Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Total Cost Till Last Month
            AddFields("@AT_REV_RECO1", "TotCost", "TOTAL COST - PTD", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //TOTAL COST - PTD
            AddFields("@AT_REV_RECO1", "EstCostCnt", "Estimated Cost contract", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Estimated cost to complete the contract
            AddFields("@AT_REV_RECO1", "TotCostCnt", "Total Contract Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Total Contract Cost
            AddFields("@AT_REV_RECO1", "PerComp", "Percentage of Completion", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage); //Percentage of Completion
            AddFields("@AT_REV_RECO1", "RevTill", "Revenue Till Last Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Revenue Till Last Month
            AddFields("@AT_REV_RECO1", "CurRevCost", "Revenue Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Revenue Current Month
            AddFields("@AT_REV_RECO1", "Revenue", "Revenue - PTD", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Revenue - PTD
            AddFields("@AT_REV_RECO1", "CurProfit", "Profit Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Profit Current Month
            AddFields("@AT_REV_RECO1", "RevtoComp", "Revenue to Complete", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Revenue to Complete
            AddFields("@AT_REV_RECO1", "EstMargin", "Estimated Margin to Complete", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Estimated Margin to Complete
            AddFields("@AT_REV_RECO1", "MartoComp", "% Margin to complete", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Percentage); //% Margin to complete
            AddFields("@AT_REV_RECO1", "InvValue", "Invoiced - PTD", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Invoiced - PTD
            AddFields("@AT_REV_RECO1", "CurInv", "Invoiced Current Month", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Invoiced Current Month
            AddFields("@AT_REV_RECO1", "TotInv", "Total Invoiced", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Total Invoiced
            AddFields("@AT_REV_RECO1", "DueFrmCus", "Due from Customer", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Due from Customer
            AddFields("@AT_REV_RECO1", "DueToCus", "Due to Customer", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Due to Customer
            AddFields("@AT_REV_RECO1", "AdvFrmCus", "Advance from customer", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Sum); //Advance from customer           
            AddFields("@AT_REV_RECO1", "OcrCode1", "Cost Center 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_REV_RECO1", "OcrCode2", "Cost Center 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_REV_RECO1", "OcrCode3", "Cost Center 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_REV_RECO1", "OcrCode4", "Cost Center 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_REV_RECO1", "OcrCode5", "Cost Center 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@AT_REV_RECO1", "RevAcc", "Revenue Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);
            AddFields("@AT_REV_RECO1", "ExpAcc", "Expense Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 15);

            AddUDO("AT_REVREC", "Revenue Recognition", SAPbobsCOM.BoUDOObjType.boud_Document, "AT_REV_RECO", new[] { "AT_REV_RECO1" }, new[] { "DocEntry", "DocNum" }, true,true, true);

        }

        #endregion

        #region Table Creation Common Functions

        private void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            // var oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception(clsModule.objaddon.objcompany.GetLastErrorDescription() + strTab);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.UDFLinkedSystemObjectTypesEnum LinkedSysObject= 0,string UDTTable="", SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

            try
            {
                // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                // If Not (strTab = "OPDN" Or strTab = "OQUT" Or strTab = "OADM" Or strTab = "OPOR" Or strTab = "OWST" Or strTab = "OUSR" Or strTab = "OSRN" Or strTab = "OSPP" Or strTab = "WTR1" Or strTab = "OEDG" Or strTab = "OHEM" Or strTab = "OLCT" Or strTab = "ITM1" Or strTab = "OCRD" Or strTab = "SPP1" Or strTab = "SPP2" Or strTab = "RDR1" Or strTab = "ORDR" Or strTab = "OWHS" Or strTab = "OITM" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OWDD" Or strTab = "OWOR" Or strTab = "OWTQ" Or strTab = "OMRV" Or strTab = "JDT1" Or strTab = "OIGN" Or strTab = "OCQG") Then
                // strTab = "@" + strTab
                // End If
                if (!IsColumnExists(strTab, strCol))
                {
                    // If Not oUserFieldMD1 Is Nothing Then
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                    // End If
                    // oUserFieldMD1 = Nothing
                    // oUserFieldMD1 = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;
                    
                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    if (LinkedSysObject != 0)
                        oUserFieldMD1.LinkedSystemObject = LinkedSysObject;// SAPbobsCOM.UDFLinkedSystemObjectTypesEnum.ulInvoices ;
                    if (UDTTable != "")
                        oUserFieldMD1.LinkedTable = UDTTable;
                    string[] split_char;
                    if (Validvalues !=null)
            {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule. objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short,true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet=null;
            string strSQL;
            try
            {
                if (clsModule. objaddon.HANA)
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                }
                else
                {
                    strSQL = "SELECT COUNT(*) FROM CUFD WHERE TableID = '" + Table + "' AND AliasID = '" + Column + "'";
                }

                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32( oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddKey(string strTab, string strColumn, string strKey, int i)
        {
            var oUserKeysMD = default(SAPbobsCOM.UserKeysMD);

            try
            {
                // // The meta-data object must be initialized with a
                // // regular UserKeys object
                oUserKeysMD =(SAPbobsCOM.UserKeysMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

                if (!oUserKeysMD.GetByKey("@" + strTab, i))
                {

                    // // Set the table name and the key name
                    oUserKeysMD.TableName = strTab;
                    oUserKeysMD.KeyName = strKey;

                    // // Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn;
                    oUserKeysMD.Elements.Add();
                    oUserKeysMD.Elements.ColumnAlias = "RentFac";

                    // // Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

                    // // Add the key
                    if (oUserKeysMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool Cancel = false, bool canlog = false, bool Manageseries = false)
        {

           SAPbobsCOM.UserObjectsMD oUserObjectMD=null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
               
                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    if(Cancel)
                        oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;

                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);
                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
            {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }

                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }

            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }

        }


        #endregion

    }
}
