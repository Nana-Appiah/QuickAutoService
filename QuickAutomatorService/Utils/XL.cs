using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Diagnostics;
using QuickAutomatorService.Models;

using ClosedXML.Excel;

namespace QuickAutomatorService.Utils
{
    public class XL
    {
        public static bool Create(List<Quick> qData, string filePath)
        {
            IXLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add(@"Quick_Data");

            ws.Cell("A1").Value = @"BRANCH";
            ws.Cell("B1").Value = @"CIF";
            ws.Cell("C1").Value = @"ACC";
            ws.Cell("D1").Value = @"KYCSTATUS";
            ws.Cell("E1").Value = @"FULLNAME";
            ws.Cell("F1").Value = @"TELEPHONE";
            ws.Cell("G1").Value = @"MONTHS_LASTDEPOSIT";
            ws.Cell("H1").Value = @"AGE_CLIENT";
            ws.Cell("I1").Value = @"AGE_ACC";
            ws.Cell("J1").Value = @"AVG_DEPOSIT_IN3MTH";
            ws.Cell("K1").Value = @"AVG_WITHDRAWN_IN3MTH";
            ws.Cell("L1").Value = @"LOANAMT";
            ws.Cell("M1").Value = @"CYCLE";
            ws.Cell("N1").Value = @"AVGDAYS12MO";
            ws.Cell("O1").Value = @"BLACKLIST";
            ws.Cell("P1").Value = @"AVGNUMDEPOSIT_IN3MTH";
            ws.Cell("Q1").Value = @"AVGNUMWITHDRAWN_IN3MTH";
            ws.Cell("R1").Value = @"GENDER";
            ws.Cell("S1").Value = @"TOTALNUMDEP";
            ws.Cell("T1").Value = @"TOTALNUMWDRAW";
            ws.Cell("U1").Value = @"ISSTAFF";
            ws.Cell("V1").Value = @"DPD";
            ws.Cell("W1").Value = @"ISLOANCLIENT";

            //creating a range for styling
            IXLRange range = ws.Range(ws.Cell(1, 1).Address, ws.Cell(1, 24).Address);
            range.Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
            range.Style.Font.Bold = true;

            //iteration
            int rowId = 2;

            foreach(var q in qData)
            {
                ws.Cell(rowId, 1).Value = q.BranchCode;
                ws.Cell(rowId, 2).Value = q.Cif;
                ws.Cell(rowId,3).SetValue(q.Acct).Style.NumberFormat.SetNumberFormatId((int)XLPredefinedFormat.Number.Text);
                ws.Cell(rowId, 4).SetValue(q.Kycstatus);
                ws.Cell(rowId, 5).Value = q.Fullname;
                ws.Cell(rowId, 6).Value = q.Telephone;
                ws.Cell(rowId, 7).Value = q.MonthsLastDeposit;
                ws.Cell(rowId, 8).Value = q.AgeClient;
                ws.Cell(rowId, 9).Value = q.AgeAcc;
                ws.Cell(rowId, 10).Value = q.AvgDepositIn3Months;
                ws.Cell(rowId, 11).Value = q.AvgWithdrawalIn3Months;
                ws.Cell(rowId, 12).Value = q.LoanAmt;
                ws.Cell(rowId, 13).Value = q.Cycle;
                ws.Cell(rowId, 14).Value = q.AvgDaysIn12Months;
                ws.Cell(rowId, 15).Value = q.BlackList;
                ws.Cell(rowId, 16).Value = q.AvgNumDepositIn3Months;
                ws.Cell(rowId, 17).Value = q.AvgNumWithdrawnIn3Months;
                ws.Cell(rowId, 18).Value = q.Gender;
                ws.Cell(rowId, 19).Value = q.TotalNumDep;
                ws.Cell(rowId, 20).Value = q.TotalNumWithdraw;
                ws.Cell(rowId, 21).Value = q.IsStaff;
                ws.Cell(rowId, 22).Value = q.Dpd;
                ws.Cell(rowId, 23).Value = q.IsLoanClient;

                rowId += 1;
            }

            ws.Columns().AdjustToContents();
            ws.Rows().AdjustToContents();

            //saving workbook
            wb.SaveAs(filePath);

            return true;
        }
    }
}
