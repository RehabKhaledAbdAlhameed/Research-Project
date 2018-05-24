using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelProject
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            #region A1
            CurrentSheet.Range["A1"].Value = "Refernce Data";
            CurrentSheet.Range["A1"].Font.Bold = true;
            #endregion
            #region A2
            CurrentSheet.Range["A2"].Value = "Company Name";
            #endregion
            #region A3
            CurrentSheet.Range["A3"].Value = "Reuter Code";
            #endregion
            #region A4
            CurrentSheet.Range["A4"].Value = "Trading Currency";
            #endregion
            #region A5
            CurrentSheet.Range["A5"].Value = "Current outstanding Shares";
            #endregion
            #region A6
            CurrentSheet.Range["A6"].Value = "Industry Name";
            #endregion
            #region A7
            CurrentSheet.Range["A7"].Value = "Sector Name";
            #endregion
          
            #region A9
            CurrentSheet.Range["A9"].Value = "Non-periodic Data";
            CurrentSheet.Range["A9"].Font.Bold = true;
            #endregion
            #region A10
            // dropdown + - neutral
            CurrentSheet.Range["A10"].Value = "Rating *";
            #endregion
            #region A11 & c11
            CurrentSheet.Range["A11"].Value = "Target Price *";
            CurrentSheet.Range["A11"].Font.Bold = true;
            CurrentSheet.Range["A11"].Borders.Color = System.Drawing.Color.Black;

            CurrentSheet.Range["B11"].Font.Bold = true;
            CurrentSheet.Range["B11"].Borders.Color = System.Drawing.Color.Black;

            #endregion

            #region A12
            CurrentSheet.Range["A12"].Value = "Investment Thesis(Text) *";
            #endregion
            #region A13
            CurrentSheet.Range["A13"].Value = "Valuation and Risks (Text) *";
            #endregion

            #region A15 & A16 & B15 & C15 & D15 &E15 & F15 & G15 & H15 & I15 

            CurrentSheet.get_Range("A15", "A16").Merge();
            CurrentSheet.Range["A15"].Value = "Valuation - Current";
            CurrentSheet.Range["A15"].VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            CurrentSheet.Range["A15"].Font.Bold = true;
            CurrentSheet.Range["A15"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Gray);



            CurrentSheet.Range["B15"].Value = "2013";
            CurrentSheet.Range["B15"].Font.Bold = true;
            CurrentSheet.Range["B15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            CurrentSheet.Range["C15"].Value = "2014";
            CurrentSheet.Range["C15"].Font.Bold = true;
            CurrentSheet.Range["C15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["D15"].Value = "2015";
            CurrentSheet.Range["D15"].Font.Bold = true;
            CurrentSheet.Range["D15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["E15"].Value = "2016";
            CurrentSheet.Range["E15"].Font.Bold = true;
            CurrentSheet.Range["E15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["F15"].Value = "2017a";
            CurrentSheet.Range["F15"].Font.Bold = true;
            CurrentSheet.Range["F15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["G15"].Value = "2018e";
            CurrentSheet.Range["G15"].Font.Bold = true;
            CurrentSheet.Range["G15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["H15"].Value = "2019e";
            CurrentSheet.Range["H15"].Font.Bold = true;
            CurrentSheet.Range["H15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["I15"].Value = "2020e";
            CurrentSheet.Range["I15"].Font.Bold = true;
            CurrentSheet.Range["I15"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;





            #endregion

            #region B16 & C16 & D16 &E16 & F16 & G16 & H16 & I16  

            CurrentSheet.Range["B16"].Value = "31/12/2013";
            CurrentSheet.Range["B16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["B16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;



            CurrentSheet.Range["C16"].Value = "31/12/2014";
            CurrentSheet.Range["C16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["C16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["D16"].Value = "31/12/2015";
            CurrentSheet.Range["D16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["D16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["E16"].Value = "31/12/2016";
            CurrentSheet.Range["E16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["E16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["F16"].Value = "31/12/2017";
            CurrentSheet.Range["F16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["F16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["G16"].Value = "31/12/2018";
            CurrentSheet.Range["G16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["G16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["H16"].Value = "31/12/2019";
            CurrentSheet.Range["H16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["H16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            CurrentSheet.Range["I16"].Value = "31/12/2020";
            CurrentSheet.Range["I16"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);
            CurrentSheet.Range["I16"].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;




            #endregion

            #region A17

            CurrentSheet.Range["A17"].Value = "Share Items";
            CurrentSheet.Range["A17"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A17"].Font.Bold = true;
            CurrentSheet.Range["A17"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A17"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
          
            #endregion

            #region A18 &  A19 & A20
            CurrentSheet.Range["A18"].Value = "Outstanding Shares (m)";
            CurrentSheet.Range["A19"].Value = "Total Number of Shares (m)";
            CurrentSheet.Range["A20"].Value = "Weighted Average Number of Shares (m)";
            #endregion
            
            #region A21

            CurrentSheet.Range["A21"].Value = "Ratios";
            CurrentSheet.Range["A21"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A21"].Font.Bold = true;
            CurrentSheet.Range["A21"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A21"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A22 &  A23 & A24
            CurrentSheet.Range["A22"].Value = "EPS *";
            CurrentSheet.Range["A23"].Value = "Dividend Per Share *";
            CurrentSheet.Range["A24"].Value = "Book Value Per Share *";
            #endregion

            #region A25

            CurrentSheet.Range["A25"].Value = "Income Statement";
            CurrentSheet.Range["A25"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A25"].Font.Bold = true;
            CurrentSheet.Range["A25"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A25"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A26 & A27 & A28 & A29 & A30 & A31 & A32 & A33 & A34 
            CurrentSheet.Range["A26"].Value = "Total Revenue (m) *";
            CurrentSheet.Range["A27"].Value = "COGS (m) *";
            CurrentSheet.Range["A28"].Value = "Gross Profit (m) *";
            CurrentSheet.Range["A29"].Value = "SG&A (m)";
            CurrentSheet.Range["A30"].Value = "Other Operating Income (Expense) (m)";
            CurrentSheet.Range["A31"].Value = "EBITDA (m) *";
            CurrentSheet.Range["A32"].Value = "Depreciation & Amortisation (m) *";
            CurrentSheet.Range["A33"].Value = "Net Operating Profit (m) *";
            CurrentSheet.Range["A34"].Value = "Attributable Net Income (m)";
            #endregion

            #region A35

            CurrentSheet.Range["A35"].Value = "Balance Sheet";
            CurrentSheet.Range["A35"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A35"].Font.Bold = true;
            CurrentSheet.Range["A35"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A35"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A36 & A37 & A38 & A39 & A40 & A41 
            CurrentSheet.Range["A36"].Value = "Cash & Cash Equivalents (m)*";
            CurrentSheet.Range["A37"].Value = "Total Current Assets (m)*";
            CurrentSheet.Range["A38"].Value = "Accumulated Depreciation (m)";
            CurrentSheet.Range["A39"].Value = "Total Non-Current Assets (m)*";
            CurrentSheet.Range["A40"].Value = "Total Assets (m)*";
            CurrentSheet.Range["A41"].Value = "Total Equity (m)*";
            #endregion

            #region A42

            CurrentSheet.Range["A42"].Value = "Cash Flow";
            CurrentSheet.Range["A42"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A42"].Font.Bold = true;
            CurrentSheet.Range["A42"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A42"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A43 & A44 & A45 & A46 
            CurrentSheet.Range["A43"].Value = "Cash Operating Profit After Taxes (m)*";
            CurrentSheet.Range["A44"].Value = "Change in Working Capital (m)";
            CurrentSheet.Range["A45"].Value = "Capex (m)";
            CurrentSheet.Range["A46"].Value = "Free Cash Flow (m)*";
            #endregion

            #region A47

            CurrentSheet.Range["A47"].Value = "KPIs";
            CurrentSheet.Range["A47"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A47"].Font.Bold = true;
            CurrentSheet.Range["A47"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A47"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A48 & A49 & A50 & A51 
            CurrentSheet.Range["A48"].Value = "Dividends (m) *";
            CurrentSheet.Range["A49"].Value = "Total Cash (m) *";
            CurrentSheet.Range["A50"].Value = "Total Revenue (Proportionate) (m) *";
            CurrentSheet.Range["A51"].Value = "EBITDA (Proportionate) (m) *";
            #endregion

            #region A52

            CurrentSheet.Range["A52"].Value = "Ratios";
            CurrentSheet.Range["A52"].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            CurrentSheet.Range["A52"].Font.Bold = true;
            CurrentSheet.Range["A52"].Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            CurrentSheet.Range["A52"].Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            #endregion

            #region A53 & A54 
            CurrentSheet.Range["A53"].Value = "ROAE (%) *";
            CurrentSheet.Range["A54"].Value = "ROAIC (%) *";
            #endregion



            CurrentSheet.Columns.AutoFit();
           
        }
        ExcelModel db = new ExcelModel();
        Company CompanyDetails = new Company();


        private void NonPeriodic(Worksheet CurrentSheet)
        {
            CompanyDetails.Comp_Name = CurrentSheet.Range["C2"].Value2.ToString();
            if (double.TryParse(CurrentSheet.Range["C3"].Value2.ToString(), out double Reu_code))
                CompanyDetails.Reu_Code = Reu_code;

            CompanyDetails.Trd_Curr = CurrentSheet.Range["C4"].Value2.ToString();
            if (decimal.TryParse(CurrentSheet.Range["C5"].Value2.ToString(), out decimal Curr_out_shares))
            {
                CompanyDetails.Curr_Out_Shares = Curr_out_shares;
            }
            if (int.TryParse(CurrentSheet.Range["C6"].Value2.ToString(), out int IndId))
            {
                CompanyDetails.Ind_id = IndId;
            }
            if (int.TryParse(CurrentSheet.Range["C7"].Value2.ToString(), out int SecId))
            {
                CompanyDetails.Sec_id = SecId;
            }
            if (!string.IsNullOrEmpty(CompanyDetails.Comp_Name))
            {
                db.Companies.Add(CompanyDetails);
                db.SaveChanges();
            }


            else
            {
                MessageBox.Show("Insert Company Name !!");
            }
        }

        private void Periodic(Worksheet CurrentSheet)
        {
            List<double> OutstandingShares;
            GetLists(out OutstandingShares, 17, CurrentSheet);

            List<double> TotalNumOfShares;
            GetLists(out TotalNumOfShares, 18, CurrentSheet);


        }
        private void GetLists(out List<double> OurList, int CellNum, Worksheet CurrentSheet)
        {
            OurList = new List<double>();
            if (double.TryParse(CurrentSheet.Range[$"C{CellNum}"].Value2.ToString(), out double cell1))
                OurList.Add(cell1);

            if (double.TryParse(CurrentSheet.Range[$"D{CellNum}"].Value2.ToString(), out double cell2))
                OurList.Add(cell2);

            if (double.TryParse(CurrentSheet.Range[$"E{CellNum}"].Value2.ToString(), out double cell3))
                OurList.Add(cell3);
            if (double.TryParse(CurrentSheet.Range[$"F{CellNum}"].Value2.ToString(), out double cell4))
                OurList.Add(cell4);

            if (double.TryParse(CurrentSheet.Range[$"G{CellNum}"].Value2.ToString(), out double cell5))
                OurList.Add(cell5);

            if (double.TryParse(CurrentSheet.Range[$"H{CellNum}"].Value2.ToString(), out double cell6))
                OurList.Add(cell6);

            if (double.TryParse(CurrentSheet.Range[$"I{CellNum}"].Value2.ToString(), out double cell7))
                OurList.Add(cell7);

            if (double.TryParse(CurrentSheet.Range[$"J{CellNum}"].Value2.ToString(), out double cell8))
                OurList.Add(cell8);

        }
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //    //Test();
            //    Worksheet CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            //    int x = int.Parse(CurrentSheet.Range["A1"].Value2.ToString());

            Worksheet CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            NonPeriodic(CurrentSheet);
            Periodic(CurrentSheet);
        }
    }
}
