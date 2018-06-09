using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Spire.Doc;
namespace WordAddIn1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
      
        private void Form1_Load(object sender, EventArgs e)
        {
            using (ResearchModel db = new ResearchModel())
            {
                var q = db.Companies.Select(c => new { c.Comp_id,c.Comp_Name });
                combobox.DataSource = q.ToList();
                combobox.ValueMember= "Comp_id";
                combobox.DisplayMember = "Comp_Name";

            }

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void combobox_SelectedIndexChanged(object sender, EventArgs e)
        {
          


        }



        private void btnCreate_Click(object sender, EventArgs e)
        {
            //  _Application word = new Word.Application();
            //  Document doc = word.Documents.Open(@"C:\Users\Nouran\Desktop\5-6word\WordAddIn1\WordAddIn1\WordAddIn1\Resources\Company Report (Portrait).docx");

            //  doc.Bookmarks["CName"].Select();
            //  word.Selection.TypeText("ay 7agaaaaaaaa22222222");

            //  Microsoft.Office.Interop.Word.Application application = new Microsoft.Office.Interop.Word.Application();
            //  Microsoft.Office.Interop.Word.Document document = new Microsoft.Office.Interop.Word.Document();
            // // string tempFile = Path.GetTempFileName();
            // //// File.WriteAllBytes(tempFile, Resources.Company_Report__Portrait_);
            ////  document = application.Documents.Add(Template: tempFile); 
            //  application.Visible = true; 
            //  application.Activate();

            //  ((_Application)word).Quit(WdSaveOptions.wdSaveChanges, WdOriginalFormat.wdOriginalDocumentFormat);
            //    var application = new Word.Application();

            //    object o = Missing.Value;
            //    object oFalse = false;
            //    object oTrue = true;

            //    Word._Application app = null;
            //    Word.Documents docs = null;
            //    Word.Document doc = null;

            //    try
            //    {
            //        app = new Word.Application();
            //        app.Visible = false;
            //        app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            //        docs = app.Documents;
            //        //string path = Path.GetDirectoryName(Path.GetFullPath(@"C:\Users\Nouran\Desktop\5-6word\WordAddIn1\WordAddIn1\WordAddIn1\"));

            //        //object path_YourDocsName = path + @"Company-Report-Portrait.docx";

            //        doc = docs.Open(@"C:\Users\Nouran\Desktop\5-6word\WordAddIn1\WordAddIn1\WordAddIn1\Company-Report-Portrait.docx", ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o, ref o);
            //        doc.Activate();

            //        foreach (Word.Range range in doc.StoryRanges)
            //        {

            //            Word.Find find = range.Find;
            //            object findText = "[ABC Bank]";
            //            //  object findText = { "[Todays date]","[]" };
            //            object replacText = "ay7agaaaaaaaaaaa";   //gets todays date and time to doc
            //            object replace = Word.WdReplace.wdReplaceAll;
            //            object findWrap = Word.WdFindWrap.wdFindContinue;
            //            find.Execute(ref findText, ref o, ref o, ref o, ref oFalse, ref o,
            //                ref o, ref findWrap, ref o, ref replacText,
            //                ref replace, ref o, ref o, ref o, ref o);

            //            Word.Find find1 = range.Find;
            //            object findText1 = "[doc content]";
            //            object replacText1 = Name;
            //            object replace1 = Word.WdReplace.wdReplaceAll;
            //            object findWrap1 = Word.WdFindWrap.wdFindContinue;
            //            find.Execute(ref findText1, ref o, ref o, ref o, ref oFalse, ref o,
            //                ref o, ref findWrap1, ref o, ref replacText1,
            //                ref replace1, ref o, ref o, ref o, ref o);

            //            Word.Find find2 = range.Find;
            //            object findText2 = "[doc content]";
            //            object replacText2 = "Somestringyouneed";
            //            object replace2 = Word.WdReplace.wdReplaceAll;
            //            object findWrap2 = Word.WdFindWrap.wdFindContinue;
            //            find.Execute(ref findText2, ref o, ref o, ref o, ref oFalse, ref o,
            //                ref o, ref findWrap2, ref o, ref replacText2,
            //                ref replace2, ref o, ref o, ref o, ref o);

            //            Word.Find find3 = range.Find;
            //            object findText3 = "[Doc content]";
            //            object replacText3 = "somesecondstringyouneed";
            //            object replace3 = Word.WdReplace.wdReplaceAll;
            //            object findWrap3 = Word.WdFindWrap.wdFindContinue;
            //            find.Execute(ref findText3, ref o, ref o, ref o, ref oFalse, ref o,
            //                ref o, ref findWrap3, ref o, ref replacText3,
            //                ref replace3, ref o, ref o, ref o, ref o);


            //            Marshal.FinalReleaseComObject(find);
            //            Marshal.FinalReleaseComObject(find1);
            //            Marshal.FinalReleaseComObject(find2);
            //            Marshal.FinalReleaseComObject(find3);

            //            Marshal.FinalReleaseComObject(range);
            //        }
            //        // var path_YourDocsName = path + @"\folder\YourDocsName.doc";
            //        //  Console.WriteLine(path_YourDocsName);
            //        doc.SaveAs(@"C:\Users\Nouran\Desktop\5-6word\WordAddIn1\WordAddIn1\WordAddIn1\Company-Report-Portrait.docx");
            //        ((Word._Document)doc).Close(ref o, ref o, ref o);
            //       // doc.Close();
            //        app.Quit(ref o, ref o, ref o);
            //    }
            //    catch (Exception ex)
            //    {
            //        MessageBox.Show (ex.StackTrace);
            //    }
            //    finally
            //    {
            //        if (doc != null) Marshal.FinalReleaseComObject(doc);

            //        if (docs != null) Marshal.FinalReleaseComObject(docs);

            //        if (app != null) Marshal.FinalReleaseComObject(app);
            //    


            //string strVal = System.IO.File.ReadAllText(Server.MapPath("~/TestFile.doc")); //--open the file and read it in a string
            //strVal = strVal.Replace("SearchText", "ReplaceText"); //--write the words one by one with the words you want to replace with.
            //System.IO.File.WriteAllText(Server.MapPath("~/TestFile.doc"), strVal); //--finally writing the file back with replaced words


            Document document = new Document();
            document.LoadFromFile(@"C:\Users\Nouran\Desktop\5-6word\WordAddIn1\WordAddIn1\WordAddIn1\Company-Report-Portrait.docx");
            ResearchModel db = new ResearchModel();
            var id = combobox.SelectedValue.ToString();
            int _id = int.Parse(id.ToString());
            Company c = db.Companies.FirstOrDefault(x => x.Comp_id == _id);
            document.Replace("ABC Bank", c.Comp_Name, false, true);

            //var result = db.NonPeriodic_Data_Draft.FirstOrDefault(r =>r.op_value =)
            //var result = (from NonPeriodic_Data_Draft in db.NonPeriodic_Data_Draft
            //              where NonPeriodic_Data_Draft.comp_id == _id && NonPeriodic_Data_Draft.item_code == 1002
            //              select NonPeriodic_Data_Draft.op_value).ToList();

            NonPeriodic_Data_Draft n = db.NonPeriodic_Data_Draft.FirstOrDefault(r => r.comp_id == _id && r.item_code == 1002);
             document.Replace("Buy", n.op_value , false, true);
            NonPeriodic_Data_Draft pt = db.NonPeriodic_Data_Draft.FirstOrDefault(r => r.comp_id == _id && r.item_code == 1003);

            document.Replace("EGP 102.0",pt.op_value , false, true);
            document.Replace("EMFD EY / EMFD.CA", c.Reu_Code , false, true);

            NonPeriodic_Data_Draft IT = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 1004);
            document.Replace("This is a test for Investment Thesis", IT.op_value , false, true);
            NonPeriodic_Data_Draft V = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 1005);
            document.Replace("This is a test for Valuation and Risks", V.op_value, false, true);

            // table 




            NonPeriodic_Data_Draft IR = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 7 && i.op_date.Value.Year == 2015);
            if (IR.op_value != null)
            {
                document.Replace("11", IR.op_value, false, true);
            }
            else
            {
                document.Replace("11", "-", false, true);

            }


            NonPeriodic_Data_Draft IR2 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 7 && i.op_date.Value.Year == 2016);
            if (IR2.op_value != null)
            {
                document.Replace("12", IR2.op_value, false, true);
            }
            else
            {
                document.Replace("12", "-", false, true);

            }

            NonPeriodic_Data_Draft IR3 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 7 && i.op_date.Value.Year == 2017);
            if (IR3.op_value != null)
            {
                document.Replace("13", IR3.op_value, false, true);
            }
            else
            {
                document.Replace("13", "-", false, true);

            }

            NonPeriodic_Data_Draft IR4 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 7 && i.op_date.Value.Year == 2018);

            if (IR4.op_value != null)
            {
                document.Replace("14", IR4.op_value, false, true);
            }
            else
            {
                document.Replace("14", "-", false, true);

            }

            // 21
            NonPeriodic_Data_Draft IR21 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 8 && i.op_date.Value.Year == 2015);
            if (IR21.op_value != null)
            {
                document.Replace("21", IR21.op_value, false, true);
            }
            else
            {
                document.Replace("21", "-", false, true);

            }

            NonPeriodic_Data_Draft IR22 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 8 && i.op_date.Value.Year == 2016);
            if (IR22.op_value != null)
            {
                document.Replace("22", IR22.op_value, false, true);
            }
            else
            {
                document.Replace("22", "-", false, true);

            }

            NonPeriodic_Data_Draft IR23 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 8 && i.op_date.Value.Year == 2017);
            if (IR23.op_value != null)
            {
                document.Replace("23", IR23.op_value, false, true);
            }
            else
            {
                document.Replace("23", "-", false, true);

            }

            NonPeriodic_Data_Draft IR24 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 8 && i.op_date.Value.Year == 2018);
            if (IR24.op_value != null)
            {
                document.Replace("24", IR23.op_value, false, true);
            }
            else
            {
                document.Replace("24", "-", false, true);

            }




            //31

            NonPeriodic_Data_Draft IR31 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 9 && i.op_date.Value.Year == 2015);
            if (IR31.op_value != null)
            {
                document.Replace("31", IR31.op_value, false, true);
            }
            else
            {
                document.Replace("31", "-", false, true);

            }


            NonPeriodic_Data_Draft IR32 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 9 && i.op_date.Value.Year == 2016);
            if (IR32.op_value != null)
            {
                document.Replace("32", IR32.op_value, false, true);
            }
            else
            {
                document.Replace("32", "-", false, true);

            }



            NonPeriodic_Data_Draft IR33 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 9 && i.op_date.Value.Year == 2017);
            if (IR33.op_value != null)
            {
                document.Replace("33", IR33.op_value, false, true);
            }
            else
            {
                document.Replace("33", "-", false, true);

            }


            NonPeriodic_Data_Draft IR34 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 9 && i.op_date.Value.Year == 2018);
            if (IR34.op_value != null)
            {
                document.Replace("34", IR34.op_value, false, true);
            }
            else
            {
                document.Replace("34", "-", false, true);

            }


            //41

            NonPeriodic_Data_Draft IR41 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 10 && i.op_date.Value.Year == 2015);
            if (IR41.op_value != null)
            {
                document.Replace("41", IR31.op_value, false, true);
            }
            else
            {
                document.Replace("41", "-", false, true);

            }


            NonPeriodic_Data_Draft IR42 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 10 && i.op_date.Value.Year == 2016);
            if (IR42.op_value != null)
            {
                document.Replace("42", IR32.op_value, false, true);
            }
            else
            {
                document.Replace("42", "-", false, true);

            }



            NonPeriodic_Data_Draft IR43 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 10 && i.op_date.Value.Year == 2017);
            if (IR43.op_value != null)
            {
                document.Replace("43", IR43.op_value, false, true);
            }
            else
            {
                document.Replace("43", "-", false, true);

            }


            NonPeriodic_Data_Draft IR44 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 10 && i.op_date.Value.Year == 2018);
            if (IR44.op_value != null)
            {
                document.Replace("44", IR44.op_value, false, true);
            }
            else
            {
                document.Replace("44", "-", false, true);

            }
            //51


            NonPeriodic_Data_Draft IR51 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 11 && i.op_date.Value.Year == 2015);
            if (IR51.op_value != null)
            {
                document.Replace("51", IR31.op_value, false, true);
            }
            else
            {
                document.Replace("51", "-", false, true);

            }


            NonPeriodic_Data_Draft IR52 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 11 && i.op_date.Value.Year == 2016);
            if (IR52.op_value != null)
            {
                document.Replace("52", IR52.op_value, false, true);
            }
            else
            {
                document.Replace("52", "-", false, true);

            }



            NonPeriodic_Data_Draft IR53 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 11 && i.op_date.Value.Year == 2017);
            if (IR53.op_value != null)
            {
                document.Replace("53", IR53.op_value, false, true);
            }
            else
            {
                document.Replace("53", "-", false, true);

            }


            NonPeriodic_Data_Draft IR54 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 11 && i.op_date.Value.Year == 2018);
            if (IR54.op_value != null)
            {
                document.Replace("54", IR54.op_value, false, true);
            }
            else
            {
                document.Replace("54", "-", false, true);

            }



            //61


            NonPeriodic_Data_Draft IR61 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 12 && i.op_date.Value.Year == 2015);
            if (IR61.op_value != null)
            {
                document.Replace("61", IR61.op_value, false, true);
            }
            else
            {
                document.Replace("61", "-", false, true);

            }


            NonPeriodic_Data_Draft IR62 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 12 && i.op_date.Value.Year == 2016);
            if (IR62.op_value != null)
            {
                document.Replace("62", IR62.op_value, false, true);
            }
            else
            {
                document.Replace("62", "-", false, true);

            }



            NonPeriodic_Data_Draft IR63 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 12 && i.op_date.Value.Year == 2017);
            if (IR63.op_value != null)
            {
                document.Replace("63", IR63.op_value, false, true);
            }
            else
            {
                document.Replace("63", "-", false, true);

            }


            NonPeriodic_Data_Draft IR64 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 12 && i.op_date.Value.Year == 2018);
            if (IR64.op_value != null)
            {
                document.Replace("64", IR64.op_value, false, true);
            }
            else
            {
                document.Replace("64", "-", false, true);

            }
            //71


            NonPeriodic_Data_Draft IR71 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 13 && i.op_date.Value.Year == 2015);
            if (IR71.op_value != null)
            {
                document.Replace("71", IR71.op_value, false, true);
            }
            else
            {
                document.Replace("71", "-", false, true);

            }


            NonPeriodic_Data_Draft IR72 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 13 && i.op_date.Value.Year == 2016);
            if (IR72.op_value != null)
            {
                document.Replace("72", IR72.op_value, false, true);
            }
            else
            {
                document.Replace("72", "-", false, true);

            }



            NonPeriodic_Data_Draft IR73 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 13 && i.op_date.Value.Year == 2017);
            if (IR73.op_value != null)
            {
                document.Replace("73", IR73.op_value, false, true);
            }
            else
            {
                document.Replace("73", "-", false, true);

            }


            NonPeriodic_Data_Draft IR74 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 13 && i.op_date.Value.Year == 2018);
            if (IR74.op_value != null)
            {
                document.Replace("74", IR74.op_value, false, true);
            }
            else
            {
                document.Replace("74", "-", false, true);

            }

            //81


            NonPeriodic_Data_Draft IR81 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 14 && i.op_date.Value.Year == 2015);
            if (IR81.op_value != null)
            {
                document.Replace("81", IR81.op_value, false, true);
            }
            else
            {
                document.Replace("81", "-", false, true);

            }


            NonPeriodic_Data_Draft IR82 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 14 && i.op_date.Value.Year == 2016);
            if (IR82.op_value != null)
            {
                document.Replace("82", IR82.op_value, false, true);
            }
            else
            {
                document.Replace("82", "-", false, true);

            }



            NonPeriodic_Data_Draft IR83 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 14 && i.op_date.Value.Year == 2017);
            if (IR83.op_value != null)
            {
                document.Replace("83", IR83.op_value, false, true);
            }
            else
            {
                document.Replace("83", "-", false, true);

            }


            NonPeriodic_Data_Draft IR84 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 14 && i.op_date.Value.Year == 2018);
            if (IR84.op_value != null)
            {
                document.Replace("84", IR84.op_value, false, true);
            }
            else
            {
                document.Replace("84", "-", false, true);

            }

            //91


            NonPeriodic_Data_Draft IR91 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 15 && i.op_date.Value.Year == 2015);
            if (IR91.op_value != null)
            {
                document.Replace("91", IR91.op_value, false, true);
            }
            else
            {
                document.Replace("91", "-", false, true);

            }


            NonPeriodic_Data_Draft IR92 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 15 && i.op_date.Value.Year == 2016);
            if (IR92.op_value != null)
            {
                document.Replace("92", IR92.op_value, false, true);
            }
            else
            {
                document.Replace("92", "-", false, true);

            }



            NonPeriodic_Data_Draft IR93 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 15 && i.op_date.Value.Year == 2017);
            if (IR93.op_value != null)
            {
                document.Replace("93", IR93.op_value, false, true);
            }
            else
            {
                document.Replace("93", "-", false, true);

            }


            NonPeriodic_Data_Draft IR94 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 15 && i.op_date.Value.Year == 2018);
            if (IR94.op_value != null)
            {
                document.Replace("94", IR94.op_value, false, true);
            }
            else
            {
                document.Replace("94", "-", false, true);

            }

            //101


            NonPeriodic_Data_Draft IR101 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 16 && i.op_date.Value.Year == 2015);
            if (IR101.op_value != null)
            {
                document.Replace("101", IR101.op_value, false, true);
            }
            else
            {
                document.Replace("101", "-", false, true);

            }


            NonPeriodic_Data_Draft IR102 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 16 && i.op_date.Value.Year == 2016);
            if (IR102.op_value != null)
            {
                document.Replace("102", IR102.op_value, false, true);
            }
            else
            {
                document.Replace("102", "-", false, true);

            }



            NonPeriodic_Data_Draft IR103 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 16 && i.op_date.Value.Year == 2017);
            if (IR103.op_value != null)
            {
                document.Replace("103", IR103.op_value, false, true);
            }
            else
            {
                document.Replace("103", "-", false, true);

            }


            NonPeriodic_Data_Draft IR104 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 16 && i.op_date.Value.Year == 2018);
            if (IR104.op_value != null)
            {
                document.Replace("104", IR104.op_value, false, true);
            }
            else
            {
                document.Replace("104", "-", false, true);

            }

            ///111
            NonPeriodic_Data_Draft IR111 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 17 && i.op_date.Value.Year == 2015);
            if (IR111.op_value != null)
            {
                document.Replace("111", IR111.op_value, false, true);
            }
            else
            {
                document.Replace("111", "-", false, true);

            }

            NonPeriodic_Data_Draft IR112 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 17 && i.op_date.Value.Year == 2016);
            if (IR112.op_value != null)
            {
                document.Replace("112", IR112.op_value, false, true);
            }
            else
            {
                document.Replace("112", "-", false, true);

            }


            NonPeriodic_Data_Draft IR113 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 17 && i.op_date.Value.Year == 2017);
            if (IR113.op_value != null)
            {
                document.Replace("113", IR113.op_value, false, true);
            }
            else
            {
                document.Replace("113", "-", false, true);

            }

            NonPeriodic_Data_Draft IR114 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 17 && i.op_date.Value.Year == 2018);
            if (IR114.op_value != null)
            {
                document.Replace("114", IR114.op_value, false, true);
            }
            else
            {
                document.Replace("114", "-", false, true);

            }

            //121


            NonPeriodic_Data_Draft IR121 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 18 && i.op_date.Value.Year == 2015);
            if (IR121.op_value != null)
            {
                document.Replace("121", IR121.op_value, false, true);
            }
            else
            {
                document.Replace("121", "-", false, true);

            }


            NonPeriodic_Data_Draft IR122 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 18 && i.op_date.Value.Year == 2016);
            if (IR122.op_value != null)
            {
                document.Replace("122", IR122.op_value, false, true);
            }
            else
            {
                document.Replace("122", "-", false, true);

            }


            NonPeriodic_Data_Draft IR123 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 18 && i.op_date.Value.Year == 2017);
            if (IR123.op_value != null)
            {
                document.Replace("123", IR123.op_value, false, true);
            }
            else
            {
                document.Replace("123", "-", false, true);

            }

            NonPeriodic_Data_Draft IR124 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 18 && i.op_date.Value.Year == 2018);
            if (IR124.op_value != null)
            {
                document.Replace("124", IR124.op_value, false, true);
            }
            else
            {
                document.Replace("124", "-", false, true);

            }



            //131


            NonPeriodic_Data_Draft IR131 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 19 && i.op_date.Value.Year == 2015);
            if (IR131.op_value != null)
            {
                document.Replace("131", IR131.op_value, false, true);
            }
            else
            {
                document.Replace("131", "-", false, true);

            }


            NonPeriodic_Data_Draft IR132 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 19 && i.op_date.Value.Year == 2016);
            if (IR132.op_value != null)
            {
                document.Replace("132", IR132.op_value, false, true);
            }
            else
            {
                document.Replace("132", "-", false, true);

            }


            NonPeriodic_Data_Draft IR133 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 19 && i.op_date.Value.Year == 2017);
            if (IR133.op_value != null)
            {
                document.Replace("133", IR133.op_value, false, true);
            }
            else
            {
                document.Replace("133", "-", false, true);

            }

            NonPeriodic_Data_Draft IR134 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 19 && i.op_date.Value.Year == 2018);
            if (IR134.op_value != null)
            {
                document.Replace("134", IR134.op_value, false, true);
            }
            else
            {
                document.Replace("134", "-", false, true);

            }


            //141


            NonPeriodic_Data_Draft IR141 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 20 && i.op_date.Value.Year == 2015);
            if (IR141.op_value != null)
            {
                document.Replace("141", IR141.op_value, false, true);
            }
            else
            {
                document.Replace("141", "-", false, true);

            }


            NonPeriodic_Data_Draft IR142 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 20 && i.op_date.Value.Year == 2016);
            if (IR142.op_value != null)
            {
                document.Replace("142", IR142.op_value, false, true);
            }
            else
            {
                document.Replace("142", "-", false, true);

            }


            NonPeriodic_Data_Draft IR143 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 20 && i.op_date.Value.Year == 2017);
            if (IR143.op_value != null)
            {
                document.Replace("143", IR143.op_value, false, true);
            }
            else
            {
                document.Replace("143", "-", false, true);

            }

            NonPeriodic_Data_Draft IR144 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 20 && i.op_date.Value.Year == 2018);
            if (IR144.op_value != null)
            {
                document.Replace("144", IR144.op_value, false, true);
            }
            else
            {
                document.Replace("144", "-", false, true);

            }



            //151
            NonPeriodic_Data_Draft IR151 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 21 && i.op_date.Value.Year == 2015);
            if (IR151.op_value != null)
            {
                document.Replace("151", IR151.op_value, false, true);
            }
            else
            {
                document.Replace("151", "-", false, true);

            }


            NonPeriodic_Data_Draft IR152 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 21 && i.op_date.Value.Year == 2016);
            if (IR152.op_value != null)
            {
                document.Replace("152", IR152.op_value, false, true);
            }
            else
            {
                document.Replace("152", "-", false, true);

            }


            NonPeriodic_Data_Draft IR153 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 21 && i.op_date.Value.Year == 2017);
            if (IR153.op_value != null)
            {
                document.Replace("153", IR153.op_value, false, true);
            }
            else
            {
                document.Replace("153", "-", false, true);

            }

            NonPeriodic_Data_Draft IR154 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 21 && i.op_date.Value.Year == 2018);
            if (IR154.op_value != null)
            {
                document.Replace("154", IR154.op_value, false, true);
            }
            else
            {
                document.Replace("154", "-", false, true);

            }


            //161
            NonPeriodic_Data_Draft IR161 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 22 && i.op_date.Value.Year == 2015);
            if (IR161.op_value != null)
            {
                document.Replace("161", IR161.op_value, false, true);
            }
            else
            {
                document.Replace("161", "-", false, true);

            }


            NonPeriodic_Data_Draft IR162 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 22 && i.op_date.Value.Year == 2016);
            if (IR162.op_value != null)
            {
                document.Replace("162", IR162.op_value, false, true);
            }
            else
            {
                document.Replace("162", "-", false, true);

            }


            NonPeriodic_Data_Draft IR163 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 22 && i.op_date.Value.Year == 2017);
            if (IR163.op_value != null)
            {
                document.Replace("163", IR163.op_value, false, true);
            }
            else
            {
                document.Replace("163", "-", false, true);

            }

            NonPeriodic_Data_Draft IR164 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 22 && i.op_date.Value.Year == 2018);
            if (IR164.op_value != null)
            {
                document.Replace("164", IR164.op_value, false, true);
            }
            else
            {
                document.Replace("164", "-", false, true);

            }


            //171

            NonPeriodic_Data_Draft IR171 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 23 && i.op_date.Value.Year == 2015);
            if (IR171.op_value != null)
            {
                document.Replace("171", IR171.op_value, false, true);
            }
            else
            {
                document.Replace("171", "-", false, true);

            }


            NonPeriodic_Data_Draft IR172 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 23 && i.op_date.Value.Year == 2016);
            if (IR172.op_value != null)
            {
                document.Replace("172", IR172.op_value, false, true);
            }
            else
            {
                document.Replace("172", "-", false, true);

            }



            NonPeriodic_Data_Draft IR173 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 23 && i.op_date.Value.Year == 2017);
            if (IR173.op_value != null)
            {
                document.Replace("173", IR173.op_value, false, true);
            }
            else
            {
                document.Replace("173", "-", false, true);

            }


            NonPeriodic_Data_Draft IR174 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 23 && i.op_date.Value.Year == 2018);
            if (IR174.op_value != null)
            {
                document.Replace("174", IR174.op_value, false, true);
            }
            else
            {
                document.Replace("174", "-", false, true);

            }



            //181


            NonPeriodic_Data_Draft IR181 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 24 && i.op_date.Value.Year == 2015);
            if (IR181.op_value != null)
            {
                document.Replace("181", IR181.op_value, false, true);
            }
            else
            {
                document.Replace("181", "-", false, true);

            }



            NonPeriodic_Data_Draft IR182 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 24 && i.op_date.Value.Year == 2016);
            if (IR182.op_value != null)
            {
                document.Replace("182", IR182.op_value, false, true);
            }
            else
            {
                document.Replace("182", "-", false, true);

            }


            NonPeriodic_Data_Draft IR183 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 24 && i.op_date.Value.Year == 2017);
            if (IR183.op_value != null)
            {
                document.Replace("183", IR183.op_value, false, true);
            }
            else
            {
                document.Replace("183", "-", false, true);

            }

            NonPeriodic_Data_Draft IR184 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 24 && i.op_date.Value.Year == 2018);
            if (IR184.op_value != null)
            {
                document.Replace("184", IR184.op_value, false, true);
            }
            else
            {
                document.Replace("184", "-", false, true);

            }


            //191


            NonPeriodic_Data_Draft IR191 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 25 && i.op_date.Value.Year == 2015);
            if (IR191.op_value != null)
            {
                document.Replace("191", IR191.op_value, false, true);
            }
            else
            {
                document.Replace("191", "-", false, true);

            }


            NonPeriodic_Data_Draft IR192 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 25 && i.op_date.Value.Year == 2016);
            if (IR192.op_value != null)
            {
                document.Replace("192", IR192.op_value, false, true);
            }
            else
            {
                document.Replace("192", "-", false, true);

            }


            NonPeriodic_Data_Draft IR193 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 25 && i.op_date.Value.Year == 2017);
            if (IR193.op_value != null)
            {
                document.Replace("193", IR193.op_value, false, true);
            }
            else
            {
                document.Replace("193", "-", false, true);

            }

            NonPeriodic_Data_Draft IR194 = db.NonPeriodic_Data_Draft.FirstOrDefault(i => i.comp_id == _id && i.item_code == 25 && i.op_date.Value.Year == 2018);
            if (IR194.op_value != null)
            {
                document.Replace("194", IR194.op_value, false, true);
            }
            else
            {
                document.Replace("194", "-", false, true);

            }

            document.SaveToFile("Replace.docx", FileFormat.Docx);

            System.Diagnostics.Process.Start("Replace.docx");


        }
    }
}