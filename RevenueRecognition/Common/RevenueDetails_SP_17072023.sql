Alter Procedure "ATPL_GetRevenueDetails" (IN fdate timestamp,IN ProjectType varchar(50)) as

begin

Select B.*,C.*, 
(B."Estimated Cost"-B."TOTAL COST - PTD") "Estimated cost to complete the contract",
(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD")) "Total Contract Cost",
Round(((B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD")))*100) ,2) "Percentage of Completion",
Round((B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")-Round((B."Total Cost Current Month"/(B."Total Cost Current Month"+(B."Estimated Cost"-B."Total Cost Current Month"))* B."Revised Contract Value"- B."Previous Revenue"),2),2) "Revenue Till Last Month",

--Round((B."Total Cost Current Month"/(B."Total Cost Current Month"+(B."Estimated Cost"-B."Total Cost Current Month"))* B."Revised Contract Value"- B."Previous Revenue"),2) "Revenue Current Month Old",

Round( (B."TOTAL COST - PTD" /(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value")  - Round((B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")-Round((B."Total Cost Current Month"/(B."Total Cost Current Month"+(B."Estimated Cost"-B."Total Cost Current Month"))* B."Revised Contract Value"- B."Previous Revenue"),2),2),2) "Revenue Current Month",


Round((B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")),2) "Revenue - PTD",
Round(((B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")-B."Total Cost Current Month") ,2) "Profit",
Round((B."Revised Contract Value"-(B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue"))) ,2) "Revenue to Complete",
Round(((B."Revised Contract Value"-(B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")))-
(B."Estimated Cost"-B."TOTAL COST - PTD")) ,2) "Estimated Margin to Complete",

Round(((((B."Revised Contract Value"-(B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")))-
(B."Estimated Cost"-B."TOTAL COST - PTD"))/(B."Revised Contract Value"-(B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue"))))*100) ,2) "% Margin to complete",

Case When  B."Total Invoiced" < Round((B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")),2) and B."Advance Amount"=0 Then (Round((B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")),2) - B."Total Invoiced") Else 0 END "Due From Customer",
Case When  B."Total Invoiced" > Round((B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")),2) and B."Advance Amount"=0 Then (B."Total Invoiced" - Round((B."Previous Revenue"+(B."TOTAL COST - PTD"/(B."TOTAL COST - PTD"+(B."Estimated Cost"-B."TOTAL COST - PTD"))* B."Revised Contract Value"- B."Previous Revenue")),2)) Else 0 END "Due To Customer"


from (Select A.*,(A."Invoiced - PTD" + A."Invoiced - Current Month") "Total Invoiced",(A."Revised Contract Value"-A."Estimated Cost") "Estimated Profit",Round((((A."Revised Contract Value"-A."Estimated Cost")/A."Estimated Cost")*100) ,2) "% of estimated MU",
(A."Material Cost Current Month" + A."YAC Labour Cost" + A."Casual Labour Cost" + A."Sub Contract Cost") "Total Cost Current Month",
((A."Cost Till Last Month" )+ ifnull((A."Material Cost Current Month" + A."YAC Labour Cost" + A."Casual Labour Cost" + A."Sub Contract Cost"),0) ) "TOTAL COST - PTD"

from (Select T1."Code" "Contract Number",T1."U_ProjType"  "Type",T1."U_Type" "Com/Res",T1."U_Status"  "Status",
T1."U_ConDesc" "Contract Description", (Select Sum("U_IProjValue") from "@AT_PROJMSTR1" Where "Code"=T1."Code" and "U_Origin"='Y')  "Original Contract Value",
((Select Sum("U_NetAdd") from "@AT_PROJMSTR1" Where "Code"=T1."Code")-(Select Sum("U_NetDed") from "@AT_PROJMSTR1" Where "Code"=T1."Code"))  "Variation Omission",
(((Select Sum("U_IProjValue") from "@AT_PROJMSTR1" Where "Code"=T1."Code" and "U_Origin"='Y') +(Select Sum("U_NetAdd") from "@AT_PROJMSTR1" Where "Code"=T1."Code"))-(Select Sum("U_NetDed") from "@AT_PROJMSTR1" Where "Code"=T1."Code")) "Revised Contract Value",
(T1."U_EstProjValue") "Estimated Cost",ifnull(T1."U_PrevRevenue",0) "Previous Revenue",ifnull(T1."U_PrevCost",0)  "Previous Cost",

(ifnull(T1."U_PrevCost",0)+
(Select (ifnull(sum(A."Debit"),0)-ifnull(sum(A."Credit"),0)) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."StornoToTr" is null and A."Project"=T1."Code"
and A."Account" in (Select "AcctCode" from OACT Where "GroupMask"=5) and B."RefDate"<= LAST_DAY(ADD_MONTHS(:fdate,-1)) )
) "Cost Till Last Month",

(Select (ifnull(sum(A."Debit"),0)-ifnull(sum(A."Credit"),0)) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."StornoToTr" is null and A."Project"=T1."Code"
and A."Account" in (Select "AcctCode" from OACT Where "GroupMask"=5) and B."RefDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate)) as "Cost Current Month",

(Select (ifnull(sum(A."Debit"),0)-ifnull(sum(A."Credit"),0)) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."StornoToTr" is null and A."Project"=T1."Code"
and A."Account" in (Select "AcctCode" from OACT Where "GroupMask"=5) and B."RefDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate) and A."TransType" in (18,19,59,60)) "Material Cost Current Month",

(Select (ifnull(sum(A."Debit"),0)-ifnull(sum(A."Credit"),0)) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."StornoToTr" is null and A."Project"=T1."Code"
and A."Account" in (Select "AcctCode" from OACT Where "GroupMask"=5) and B."RefDate"<= LAST_DAY(ADD_MONTHS(:fdate,-1)) and A."TransType" in (18,19,59,60)) "Material Cost Till Last Month",

(Select ifnull(Sum(A."Debit"),0) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."TransType"='30' and B."StornoToTr" is null and A."Project"=T1."Code" and B."TransCode"='YL'
and A."Account" in (Select "U_Cosglc" From "@AT_PROJMSTR4" Where "Code"=T1."Code" and "U_LabType"='YL') and B."RefDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate) ) "YAC Labour Cost",

(Select ifnull(Sum(A."Debit"),0) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."TransType"='30' and B."StornoToTr" is null and A."Project"=T1."Code" and B."TransCode"='CL'
and A."Account" in (Select "U_Cosglc" From "@AT_PROJMSTR4" Where "Code"=T1."Code" and "U_LabType"='CL') and B."RefDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate) ) "Casual Labour Cost",

(Select ifnull(SUM(A."Debit"),0) from JDT1 A Left Join OJDT B On A."TransId"=B."TransId" Where B."TransType"='30' and B."StornoToTr" is null 
and A."Project"=T1."Code" and B."RefDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate) and UPPER(A."OcrCode2") like '%SUBCONTRACT' ) "Sub Contract Cost",

(Select ifnull(sum(B."DocTotal"),0) from INV1 A Left Join OINV B On A."DocEntry"=B."DocEntry" Where B."CANCELED"='N' and A."Project"=T1."Code" and B."DocDate"<= LAST_DAY(ADD_MONTHS(:fdate,-1))) "Invoiced - PTD",
(Select ifnull(sum(B."DocTotal"),0) from INV1 A Left Join OINV B On A."DocEntry"=B."DocEntry" Where B."CANCELED"='N' and A."Project"=T1."Code" and B."DocDate" between ADD_MONTHS(NEXT_DAY(LAST_DAY(:fdate)),-1) and LAST_DAY(:fdate)) "Invoiced - Current Month",

(Select IFNULL(SUM(A."CashSum"+ A."CreditSum" + A."CheckSum" + A."TrsfrSum"),0) from ORCT A join RCT2 B On A."DocEntry"=B."DocNum" Where A."Canceled"='N'
and B."DocEntry" in (Select "DocEntry" from DPI1 Where "BaseType"='17' and "BaseEntry"=(Select "U_SOEntry" From "@AT_PROJMSTR1" Where "U_Origin"='Y' and "Code"=T1."Code")) ) "Advance Amount"

from "@AT_PROJMSTR" T1 Left Join OPRJ T2 On T1."Code"=T2."PrjCode"

---Where UPPER(T2."U_PrjType")='COMMERCIAL'
Where UPPER(T2."U_PrjType")=:ProjectType and T1."U_Status"<>'C' and T1."U_POC"='Y'  and :fdate between T1."U_Date" and CURRENT_DATE

and T1."Code" not in (Select A."U_Project" from "@AT_REV_RECO1" A Left Join "@AT_REV_RECO" B On A."DocEntry"=B."DocEntry" 
Where B."Status"='C' and B."U_VoucherID" is not null and B."U_Year"= Cast(YEAR(:fdate) as varchar) 
and B."U_Month"=(Case when LENGTH(MONTH(:fdate))=1 Then Cast ('0' || MONTH(:fdate) as varchar) Else Cast (MONTH(:fdate) as varchar) End))

) A) B 

Left Join 

(Select distinct T."Project",T."DocEntry" "SO DocEntry",T."OcrCode" "Cost Center 1",T."OcrCode2" "Cost Center 2",T."OcrCode3" "Cost Center 3",
T."OcrCode4" "Cost Center 4",T."OcrCode5" "Cost Center 5",(Select "U_RevAcc" from OPRJ Where "PrjCode"=T."Project") "Revenue Account",
(Select "U_ExpAcc" from OPRJ Where "PrjCode"=T."Project") "Expense Account" 
from RDR1 T Where T."DocEntry"=(Select "U_SOEntry" from "@AT_PROJMSTR1" Where "Code"=T."Project" and "U_Origin"='Y')) C On B."Contract Number"=C."Project";

End;