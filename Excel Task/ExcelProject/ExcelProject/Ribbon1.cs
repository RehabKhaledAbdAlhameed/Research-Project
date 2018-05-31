using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;
using System.Data.Entity;

namespace ExcelProject
{
    public partial class Ribbon1
    {

        #region Declarations
        int RateCode = 1003, TargetPriceCode = 1004, ITCode = 1005, VRCode = 1005;
        Worksheet CurrentSheet;
        Company MyCompanyData=new Company();
        private bool Flag = true;bool IsThereLastPrice = false;bool Has_NonPeriodic_Data_Draft = false; bool UploadBtnFlag = false;
        int index = 0;double OldTargetPrice = 0.0; int ListIndex = 0;

        ExcelModel db = new ExcelModel();
        Company CompanyDetails = new Company();
        Data_item items = new Data_item();
        Data_type type = new Data_type();
        NonPeriodic_Data_Draft MyData = new NonPeriodic_Data_Draft();
        NonPeriodic_Data_Draft CurrentNonPeriodicData= new NonPeriodic_Data_Draft();

        List<NonPeriodic_Data_Draft> List_NonPeriodic_DataDraft_Data = new List<NonPeriodic_Data_Draft>();

        List<Data_item> DataItems = new List<Data_item>();

        #endregion
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            MyCompanyData = new Company();
             

        }

        #region 1.Create ExcelSheet

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet(); 
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
        #endregion

        #region 2.GetCompanyData from DB

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
               UploadBtnFlag = true;
               CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            //string r = CurrentSheet.Range["B3"].Value?.ToString();
            if (CurrentSheet.Range["B3"].Value==null)
            {
                MessageBox.Show("Enter The Company Reuter Code Please ");
                CurrentSheet.Range["B3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                CurrentSheet.Range["A3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

            }
            else
            {
                string ReuterCode= CurrentSheet?.Range["B3"]?.Value?.ToString();
                if (!string.IsNullOrEmpty(ReuterCode))//double.TryParse(CurrentSheet.Range["B3"].Value.ToString(), out ReuterCode)
                {

                    MyCompanyData = db.Companies.FirstOrDefault(x => x.Reu_Code == ReuterCode);
                    if (MyCompanyData != null)
                    {
                        CurrentSheet.Range["B3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        CurrentSheet.Range["A3"].Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                        //Fill All Excel Sheet
                        FillCompanyData(MyCompanyData);
                        Fill_Required_NonPeriodic_Data(MyCompanyData);
                        Fill_NonPeriodic_Data(MyCompanyData);

                    }
                    else
                    {
                        MessageBox.Show("This Company doesn't exist in our Data base please try another ReuterCode ");
                    }
                }
            }
           

        }

        private void FillCompanyData(Company company)
        {
            CurrentSheet.Range["B2"].Value2 = company.Comp_Name;
            CurrentSheet.Range["B4"].Value2 = company.Trd_Curr;
            CurrentSheet.Range["B5"].Value2 = company.Curr_Out_Shares;
            CurrentSheet.Range["B6"].Value2 = company.Industry.Ind_Name;
            CurrentSheet.Range["B7"].Value2 = company.Sector.Sec_Name;
           
        }

        private void Fill_Required_NonPeriodic_Data(Company company)
        {

            List_NonPeriodic_DataDraft_Data = (from e in db.NonPeriodic_Data_Draft
                                        where e.comp_id == company.Comp_id
                                        select e).ToList();
            if (List_NonPeriodic_DataDraft_Data?.Count != 0)
            {
                Has_NonPeriodic_Data_Draft = true;
                MyData = List_NonPeriodic_DataDraft_Data.FirstOrDefault(item => item.item_code == RateCode);//Rating
                if (!string.IsNullOrEmpty(MyData?.op_value))
                {
                    CurrentSheet.Range["B10"].Value2 = MyData.op_value;

                }
                MyData = List_NonPeriodic_DataDraft_Data.FirstOrDefault(item => item.item_code == TargetPriceCode);//Target Price
                if (!string.IsNullOrEmpty(MyData?.op_value))
                {
                    IsThereLastPrice = true;
                    OldTargetPrice=double.Parse(MyData.op_value);
                    CurrentSheet.Range["B11"].Value2 = MyData.op_value;

                }
                MyData = List_NonPeriodic_DataDraft_Data.FirstOrDefault(item => item.item_code == ITCode);//Investment Thesis(Text)
                if (!string.IsNullOrEmpty(MyData?.op_value))
                {
                    CurrentSheet.Range["B12"].Value2 = MyData.op_value;

                }
                MyData = List_NonPeriodic_DataDraft_Data.FirstOrDefault(item => item.item_code == VRCode);//Valuation and Risks (Text)
                if (!string.IsNullOrEmpty(MyData?.op_value))
                {
                    CurrentSheet.Range["B13"].Value2 = MyData.op_value;

                }

            }
            else
            {
                MessageBox.Show("You didn't enter any (Non Periodic Data) Before");
            }
           

        }

    
        private void Fill_NonPeriodic_Data(Company company)
        {

            if (List_NonPeriodic_DataDraft_Data?.Count!=0)
            {
                Microsoft.Office.Interop.Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
                Microsoft.Office.Interop.Excel.Range Range_ = activeSheet.UsedRange;
                int index = 4;//Skip 4 rows (Rating,Target Price,Investment Thesis(Text),Valuation and Risks (Text)
                for (int a = 1; a <= Range_.Areas.Count; a++)
                {
                    Microsoft.Office.Interop.Excel.Range area = Range_.Areas[a];

                    for (int r = 18; r <= 54; r++)
                    {

                        for (int c = 2; c <= 9; c++)
                        { // Cell value 
                            if (r == 21 || r == 25 || r == 29 || r == 30 || r == 34 || r == 35 || r == 38 || r == 42 || r == 44 || r == 45 || r == 47 || r == 52)
                            {
                                continue;
                            }

                           ((Microsoft.Office.Interop.Excel.Range)area[r, c]).Value2 = List_NonPeriodic_DataDraft_Data[index++].op_value;




                        }
                    }
                }


            }
        }


        #endregion


        #region 3.SaveCompanyData To DB


        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            CurrentSheet = Globals.ThisAddIn.GetActiveWorkSheet();

            Flag = true;index = 0;ListIndex = 4;

            if (List_NonPeriodic_DataDraft_Data?.Count != 0)
            {
                Has_NonPeriodic_Data_Draft = true;
            }
           


                CheckRequired();

            if (Flag == true && UploadBtnFlag == true)
            {
                //Periodic(CurrentSheet);
                Required_Nonperiodic_data(CurrentSheet);
                NonPeriodic(CurrentSheet);
                UploadBtnFlag = false;
                MessageBox.Show("All Data Saved Thanks ");

            }
            else
            {
                MessageBox.Show("Follow all instructions please :) "); 
            }

        }



        #region Periodic data Function
        private void Periodic(Worksheet CurrentSheet)
        {
            #region B2
            if (CurrentSheet.Range["B2"].Value != null)
            {
                CompanyDetails.Comp_Name = CurrentSheet.Range["B2"].Value2.ToString();
            }
            else
            {
                MessageBox.Show("Please Enter Value in Company Name Field !! ");
            }

            #endregion

            #region B3
            string RouterCode = CurrentSheet.Range["B3"].Value?.ToString();

            if (!string.IsNullOrEmpty(RouterCode) && !RouterCode.Any(x=>x=='-') )
            {
                float r;
                if (RouterCode.StartsWith("(") && RouterCode.EndsWith(")"))
                {
                    string RouterCode2 = RouterCode.Remove(0, 1);
                    string RouterCode3 = RouterCode2.Remove(RouterCode2.Length - 1, 1);

                    if (float.TryParse(RouterCode3, out r))
                        CompanyDetails.Reu_Code = $"{-r}";
                    else
                        MessageBox.Show("Enter A valid Value between The brackets In Reuter Code Field");


                }

                else if (float.TryParse(RouterCode, out r))
                    CompanyDetails.Reu_Code = $"{-r}";
                else
                {
                    MessageBox.Show("Enter A real  valid Value In Reuter Code Field");
                }
            }

            else
            {
                MessageBox.Show("Enter A valid Value In Reuter Code Field");
            }

            #endregion

            #region B4
            if (CurrentSheet.Range["B4"].Value != null)
            {
                CompanyDetails.Trd_Curr = CurrentSheet.Range["B4"].Value2.ToString();
            }
            else
            {
                MessageBox.Show("Please Enter Value in Trading Currency Field !! ");
            }

            #endregion

            #region B5

            if (CurrentSheet.Range["B5"].Value != null && decimal.TryParse(CurrentSheet.Range["B5"].Value2.ToString(), out decimal Curr_out_shares))
            {
                CompanyDetails.Curr_Out_Shares = decimal.Parse(CurrentSheet.Range["B5"].Value2.ToString());
            }

            else
            {
                MessageBox.Show("Enter A valid Value In Current Out Shares Field");
            }

            #endregion

            #region B6

            if (CurrentSheet.Range["B6"].Value2 != null)
            {
                string IndustryName = CurrentSheet.Range["B6"].Value2.ToString();
                switch (IndustryName)
                {
                    case "Mobile Phones":
                        CompanyDetails.Ind_id = 3;
                        break;
                    case "Face book":
                        CompanyDetails.Ind_id = 4;
                        break;

                    default:
                        MessageBox.Show("Enter One Of These Industry Name with the same spelling " + "\n" + "Mobile Phones" + "\n" + "Face book");
                        break;
                }

            }

            else
            {
                MessageBox.Show("Enter A valid Value In Industry Name Field");
            }

            #endregion

            #region B7

            if (CurrentSheet.Range["B7"].Value2 != null)
            {
                string SectorName = CurrentSheet.Range["B7"].Value2.ToString();
                switch (SectorName)
                {
                    case "Communication":
                        CompanyDetails.Sec_id = 1;
                        break;
                    case "Air Crafts":
                        CompanyDetails.Sec_id = 2;
                        break;
                    case "Consumer Dictionary":
                        CompanyDetails.Sec_id = 3;

                        break;

                    default:
                        MessageBox.Show("Enter valid Sector Name please");

                        Flag = false;
                        break;

                }
                if (Flag == true)
                {
                    db.Companies.Add(CompanyDetails);
                    db.SaveChanges();
                }

            }

            else
            {
                MessageBox.Show("Enter One Of These Sector Name with the same spelling " + "\n" + "Communication" + "\n" + "Air Crafts" + "\n" + "Consumer Dictionary");

            }

            #endregion


        }
        #endregion


        #region Required_Nonperiodic_data_from A10 to A13
        private void Required_Nonperiodic_data(Worksheet CurrentSheet)
        {
                    // MyData = new NonPeriodic_Data_Draft();
                    string CompanyName = CurrentSheet?.Range["B2"].Value?.ToString();
                    if (!string.IsNullOrEmpty(MyCompanyData.Comp_Name)) { 
                   /* var Company = (from y in db.Companies
                                   where y.Comp_Name == CompanyName
                                   select y).FirstOrDefault();*/

                    int CompanyID = MyCompanyData.Comp_id;


                    string Rate = CurrentSheet?.Range["B10"].Value?.ToString();
                    string targetPrice = CurrentSheet?.Range["B11"].Value?.ToString();
                    string inv_tehs = CurrentSheet?.Range["B12"].Value?.ToString();
                    string valrisks = CurrentSheet?.Range["B13"].Value?.ToString();

                
                    #region Rate

                  //while (string.IsNullOrEmpty(Rate))
                  //  {
                  //      MessageBox.Show("Enter A Valid Value in Rating Cell");

                  //  }
               
                Rate = Rate.ToLower();
                /*  while (!Rate.Equals("neutral") || !Rate.Equals("sell") || !Rate.Equals("buy"))
                  {
                      MessageBox.Show("Rating Value Should be (Neutral or Buy or Sell");

                  }
                  */
                //Update current Data
                if (Has_NonPeriodic_Data_Draft)
                {
                    MyData = List_NonPeriodic_DataDraft_Data[0];//Rating First Row
                    MyData.op_date = DateTime.Now.ToLocalTime();
                    if (Rate == "neutral" || Rate == "buy"|| Rate == "sell")
                    {

                        MyData.op_value = Rate;
                        db.Entry(MyData).State = EntityState.Modified;


                    }
            
                    else
                    {
                        MessageBox.Show("Rating Value Should be (Neutral or Buy or Sell");
                    }

                }

                //Add New
                else
                {
                    MyData.op_date = DateTime.Now.ToLocalTime();

                    if (Rate == "neutral" || Rate == "buy" || Rate == "sell")
                        {
                            MyData.op_value = Rate;
                            MyData.item_code = RateCode;
                            MyData.comp_id = int.Parse(CompanyID.ToString());
                            db.NonPeriodic_Data_Draft.Add(MyData);

                        }
                    
                        else
                        {
                            MessageBox.Show("Rating Value Should be (Neutral or Buy or Sell");
                        }
                    }
                    db.SaveChanges();

                    #endregion

                    #region Target Price
                    double targetPriceDouble= double.Parse(targetPrice);
                    //while (string.IsNullOrEmpty(targetPrice))
                    //{
                    //    MessageBox.Show("Enter  a value in Target Price Cell please");

                    //}
                    //while (!double.TryParse(targetPrice, out targetPriceDouble))
                    //{
                    //    MessageBox.Show("Enter Valid Number in Target Price Cell please");

                    //}
                    if (IsThereLastPrice)
                    {/*
                
                       ex:OldTargetPrice=1000    =>  10% from 1000= 100
                       then targetPriceDouble should be between 900 and 1100
                       */


                    //    double PrecentageValue = OldTargetPrice * .1;
                    //    double MaxValue = OldTargetPrice + PrecentageValue;
                    //    double MinValue = OldTargetPrice - PrecentageValue;

                    //if(targetPriceDouble > MaxValue || targetPriceDouble < MinValue) {

                       
                    //    MessageBox.Show($"Target price should be between {MinValue} and {MaxValue} please");

                    //}//Target Price Second Row
                    MyData = List_NonPeriodic_DataDraft_Data[1];
                    MyData.op_date = DateTime.Now.ToLocalTime();
                    MyData.op_value = targetPrice;
                    db.Entry(MyData).State = EntityState.Modified;

                 

             }
                    //New Target Price
             else
                    {
                    MyData.op_date = DateTime.Now.ToLocalTime();
                    MyData.op_value = targetPrice;
                    MyData.item_code = TargetPriceCode;//in table Data_item
                    MyData.comp_id = int.Parse(CompanyID.ToString());
                    db.NonPeriodic_Data_Draft.Add(MyData);
                }
                  
                    db.SaveChanges();

                    #endregion

                    #region Investment

                    while (string.IsNullOrEmpty(inv_tehs))
                    {
                        MessageBox.Show("Enter Valid Value in Valuation And Risks Thesis Cell");

                    }
                    if (Has_NonPeriodic_Data_Draft)
                    {
                        MyData = List_NonPeriodic_DataDraft_Data[2];//Investment Thesis(Text) Third Row
                        MyData.op_date = DateTime.Now.ToLocalTime();
                        MyData.op_value = inv_tehs;
                        db.Entry(MyData).State = EntityState.Modified;


                }
                else
                    {
                        MyData.op_value = inv_tehs;
                        MyData.op_date = DateTime.Now.ToLocalTime();
                        MyData.item_code = ITCode;//in table Data_item
                        MyData.comp_id = int.Parse(CompanyID.ToString());
                        db.NonPeriodic_Data_Draft.Add(MyData);
                    }
                    db.SaveChanges();

                    #endregion

                    #region valrisks

                    while (string.IsNullOrEmpty(valrisks))
                    {
                        MessageBox.Show("Enter Valid Value in Valuation and Risks  Cell");

                    }
                    if (Has_NonPeriodic_Data_Draft)
                    {

                        MyData = List_NonPeriodic_DataDraft_Data[3];//Valuation and Risks (Text)  Fourth Row
                        MyData.op_date = DateTime.Now.ToLocalTime();
                        MyData.op_value = valrisks;
                        db.Entry(MyData).State = EntityState.Modified;

                }
                else
                    {
                        MyData.op_value = valrisks;
                        MyData.op_date = DateTime.Now.ToLocalTime();
                        MyData.item_code = VRCode;//in table Data_item
                        MyData.comp_id = int.Parse(CompanyID.ToString());
                        db.NonPeriodic_Data_Draft.Add(MyData);
                    }
                        db.SaveChanges();

                    #endregion



        }
    }
        #endregion
       

        #region Non-periodic Function
        private void NonPeriodic(Worksheet CurrentSheet)
        {
            //31
            /*Company LastCompany = (from x in db.Companies
                                   orderby x.Comp_id descending
                                   select x).FirstOrDefault();
            */
            for (int row = 18;row < 55; row++)
            {

                if (row == 21|| row == 25|| row == 35|| row == 42|| row == 47|| row == 52)
                {
                    continue;
                }

                if (Flag != false)
                    GetLists(row, CurrentSheet , MyCompanyData);
                else
                    return;

            }

        }
        #endregion


        #region Get lists Function
        private void GetLists(int RowNum, Worksheet CurrentSheet, Company CurrentCompany)
        {
            //Data can be Empty for some cells , it's ok 
            //if (CurrentSheet.Range[$"B{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'B', CurrentCompany, index);
                ListIndex++;

            //}

            //if (CurrentSheet.Range[$"C{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'C', CurrentCompany, index);
                ListIndex++;

            //}

            //if (CurrentSheet.Range[$"D{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'D', CurrentCompany, index);
                ListIndex++;

            //}

            //if (CurrentSheet.Range[$"E{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'E', CurrentCompany, index);
                ListIndex++;

            //}
            //if (CurrentSheet.Range[$"F{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'F', CurrentCompany, index);
                ListIndex++;

            //}
            //if (CurrentSheet.Range[$"G{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'G', CurrentCompany, index);
                ListIndex++;

            //}
            //if (CurrentSheet.Range[$"H{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'H', CurrentCompany, index);
                ListIndex++;

            //}
            //if (CurrentSheet.Range[$"I{RowNum}"].Value2 != null && Flag != false)
            //{
                addCellForRow(CurrentSheet, RowNum, 'I', CurrentCompany, index);
                ListIndex++;

            //}

            index++;
            //Index (Number of Rows) to use it in   :             
            //MyData.item_code = db.Data_item.ToList()[index].item_code ;//DataItems[index].item_code;


        }
        #endregion


        #region Add Every cell for everyRow
        void addCellForRow(Worksheet CurrentSheet, int CellNum, char CellChar, Company CurrentCompany, int index)
        {
            string cell = CurrentSheet?.Range[$"{CellChar}{CellNum}"].Value2?.ToString();
            if (Has_NonPeriodic_Data_Draft)
            {
                MyData = List_NonPeriodic_DataDraft_Data[ListIndex];
                MyData.op_value = cell;
                db.Entry(MyData).State = EntityState.Modified;
                db.SaveChanges();
            }
            else
            {

                MyData.op_value = cell;
                DateTime date;
                //if (DateTime.TryParse(CurrentSheet.Range[$"{CellChar}{16}"].Value2.ToString(), out date))
                //    MyData.op_date = date;

                MyData.op_date = DateTime.ParseExact(CurrentSheet.Range[$"{CellChar}{16}"].Value2.ToString(), "dd/MM/yyyy", CultureInfo.InvariantCulture);

               
                MyData.item_code = db.Data_item.ToList()[index].item_code ;//DataItems[index].item_code;
                




                MyData.comp_id = CurrentCompany.Comp_id;

                db.NonPeriodic_Data_Draft.Add(MyData);
                db.SaveChanges();

            }



        }
        #endregion



        private void CheckRequired()
        {
            Microsoft.Office.Interop.Excel.Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;
            Microsoft.Office.Interop.Excel.Range Range_ = activeSheet.UsedRange;

            for (int a = 1; a <= Range_.Areas.Count; a++)
            {
                Microsoft.Office.Interop.Excel.Range area = Range_.Areas[a];

                for (int r = 22; r <= 54; r++)
                {
                    if (Flag == false)
                        break;
                    for (int c = 2; c <= 9; c++)
                    { // Cell value 
                        if ( r == 25 || r == 29 || r == 30 || r == 34 || r == 35 || r == 38 || r == 42 || r == 44 || r == 45 || r == 47 || r == 52)
                        {
                            continue;
                        }

                        //CellValue = ((Microsoft.Office.Interop.Excel.Range)area[r, c]).Value2;
                        if (((Microsoft.Office.Interop.Excel.Range)area[r, c]).Value2 == null)
                        {
                            Flag = false;
                            MessageBox.Show("Enter All Required Fields Please :)");
                           
                            return;

                        }

                        string v = ((Microsoft.Office.Interop.Excel.Range)area[r, c]).Value2.ToString();
                        if (v.Any(x=>x=='-'))
                        {
                            Flag = false;
                            MessageBox.Show("Enter Positive Number Please");
                            
                            return;
                        }

                    }
                }
                for (int r = 18; r <= 21; r++)
                {
                    if (Flag == false)
                        break;
                    for (int c = 2; c <= 9; c++)
                    { // Cell value 
                        
                   

                        string v = ((Microsoft.Office.Interop.Excel.Range)area[r, c]).Value2?.ToString();

                        if (!string.IsNullOrEmpty(v))
                        {
                            if (v.Any(x => x == '-'))
                            {
                                Flag = false;
                                MessageBox.Show("Enter Positive Number Please");

                                return;
                            }

                            if (r == 18)
                            {

                            }
                        
                        }
                      

                    }
                }


                for (int r = 10; r <= 13; r++)
                {
                    if (Flag == false)
                        return;
                  


                    string v = ((Microsoft.Office.Interop.Excel.Range)area[r, 2]).Value2?.ToString();

                    if (!string.IsNullOrEmpty(v))
                    {
                          

                        if (r == 10)
                        {
                                if (v == "neutral" || v == "buy" || v == "sell")
                                {
                              
                                }

                                else
                                {
                                    MessageBox.Show("Rating Value Should be (Neutral or Buy or Sell");
                                }
                        }


                    }
                    else
                    {
                        Flag = false;
                        MessageBox.Show("Enter All Required Fields Please :)");
                        return;
                    }
     }
  }



            #region TargetPrice

            double targetPriceDouble;
            string targetPrice = CurrentSheet?.Range["B11"].Value?.ToString();

            if (!double.TryParse(targetPrice, out targetPriceDouble))
            {
                MessageBox.Show("Enter Valid Number in Target Price Cell please");
                Flag = false;
                return;
            }
            if (IsThereLastPrice)
                {/*
                
                           ex:OldTargetPrice=1000    =>  10% from 1000= 100
                           then targetPriceDouble should be between 900 and 1100
                           */


                    double PrecentageValue = OldTargetPrice * .1;
                    double MaxValue = OldTargetPrice + PrecentageValue;
                    double MinValue = OldTargetPrice - PrecentageValue;

                    if (targetPriceDouble > MaxValue || targetPriceDouble < MinValue)
                    {


                        MessageBox.Show($"Target price should be between {MinValue} and {MaxValue} please");
                        Flag = false;
                        return;

                    }//Target Price Second Row
                  


                }
            #endregion
        }


       
     

  #endregion
    }
}
