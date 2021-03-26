using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPnt = Microsoft.Office.Interop.PowerPoint;
using WinWord = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace insertGuaXingtoPowerpnt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> lb = new List<string>{"64卦圖","行書","小篆","甲骨文"
                    ,"金文","隸書"};
            listBox1.DataSource = lb;
            List<string> lb2 = new List<string> { "PowerPoint", "Word", "Excel" };
            listBox2.DataSource = lb2;

        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            go();
        }

        private void go()
        {

            officeEnum ofcE = officeEnum.PowerPoint;
            switch (listBox2.SelectedItem)
            {
                case "PowerPoint":
                    ofcE = officeEnum.PowerPoint;
                    break;
                case "Word":
                    ofcE = officeEnum.Word;
                    break;
                case "Excel":
                    ofcE = officeEnum.Excel;
                    break;
                default:
                    break;
            }
            switch (listBox1.SelectedValue)
            {
                case "64卦圖":
                    guaXing(ofcE);
                    break;
                case "行書":
                    GuWenZi(picEnum.行書, ofcE);
                    break;
                case "小篆":
                    GuWenZi(picEnum.小篆, ofcE);
                    break;
                case "甲骨文":
                    break;
                case "金文":
                    break;
                case "隸書":
                    GuWenZi(picEnum.隸書, ofcE);
                    break;
                default:
                    break;
            }
        }

        void guaXing(officeEnum oE)
        {
            string dir = getDir(picEnum.卦圖64) + "\\";
            if (dir != "")
            {
                switch (oE)
                {
                    case officeEnum.PowerPoint:
                        runPPTGuaXing(dir);
                        break;
                    case officeEnum.Word:
                        runDOCGuaXing(dir);
                        break;
                    case officeEnum.Excel:
                        break;
                    default:
                        break;
                }
            }
        }

        void runPPTGuaXing(string dir)
        {
            PowerPnt.Application pptApp = (PowerPnt.Application)getOffice(officeEnum.PowerPoint);
            PowerPnt.Presentation ppt = pptApp.ActivePresentation;
            PowerPnt.Selection sel = ppt.Application.ActiveWindow.Selection;
            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
            {
                PowerPnt.Slide sld = ppt.Application.ActiveWindow.View.Slide;
                string f = dir + sel.TextRange.Text + ".png";
                if (System.IO.File.Exists(f))
                {
                    PowerPnt.TextRange tr = sel.TextRange.InsertAfter(sel.TextRange.Text);
                    tr.Select();
                    sel = sel.Application.ActiveWindow.Selection;
                    float lf = sel.TextRange.BoundLeft;
                    float tp = sel.TextRange.BoundTop;
                    int selCt = sel.TextRange2.Characters.Count;
                    float h = sel.TextRange.BoundHeight;
                    if (sel.ShapeRange.TextFrame.Orientation == Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationVerticalFarEast)
                    {
                        h = sel.TextRange.BoundHeight / selCt;
                    }
                    float w = sel.TextRange.BoundWidth;
                    if (sel.ShapeRange.TextFrame.Orientation == Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal)
                    {
                        w = sel.TextRange.BoundWidth / selCt;
                    }
                    PowerPnt.Shape sp = sld.Shapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                        , Microsoft.Office.Core.MsoTriState.msoTrue,
                        lf, tp, w, h);
                    sel.TextRange.Text = "　";
                    spTransp(ref sp, sel.TextRange2);
                    sel.Unselect();
                }

            }
            ppt.Application.Activate();
        }

        void runDOCGuaXing(string dir)//Word插入卦形
        {
            WinWord.Application docApp = (WinWord.Application)getOffice(officeEnum.Word);
            WinWord.Document doc = docApp.ActiveDocument;
            WinWord.Selection sel = doc.Application.ActiveWindow.Selection;
            WinWord.InlineShape sp;
            WinWord.Range rng = sel.Range;
            string f = dir + sel.Text + ".png";
            if (!System.IO.File.Exists(f) && rng.Characters.Count == 1)
            {
                rng.SetRange(rng.Start, rng.Characters[1].Next().End);
                f = dir + rng.Text + ".png";
            }
            if (System.IO.File.Exists(f))
            {
                docApp.ScreenUpdating = false;
                WinWord.WdColorIndex c = sel.Range.HighlightColorIndex;
                sp = sel.InlineShapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                    , Microsoft.Office.Core.MsoTriState.msoTrue);
                sp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                sp.Height = (float)0.9 * (15 + sel.Range.Font.Size - 12);
                sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                sp.PictureFormat.TransparencyColor = 16777215;
                sp.Range.HighlightColorIndex = c;
                if (sel.ParagraphFormat.BaseLineAlignment != WinWord.WdBaselineAlignment.wdBaselineAlignCenter)
                {
                    sel.ParagraphFormat.BaseLineAlignment = WinWord.WdBaselineAlignment.wdBaselineAlignCenter;
                }
                docApp.ScreenUpdating = true;
            }
            doc.Application.Activate();
        }
        void GuWenZi(picEnum pE, officeEnum ofE)
        {
            string dir = getDir(pE) + "\\";
            if (dir != "")
            {
                switch (ofE)
                {
                    case officeEnum.PowerPoint:
                        runPPT(dir, pE);
                        break;
                    case officeEnum.Word:
                        runDOC(dir, pE);
                        break;
                    case officeEnum.Excel:
                        break;
                    default:
                        break;
                }

            }
        }

        void runDOC(string dir, picEnum pE)//Word插入字圖
        {
            WinWord.Application docApp = (WinWord.Application)getOffice(officeEnum.Word);
            WinWord.Document doc = docApp.ActiveDocument;
            WinWord.Selection sel = doc.Application.ActiveWindow.Selection;
            WinWord.InlineShape sp; WinWord.Shape s;
            string extName = ".png";
            bool inserted = false;
            switch (pE)
            {
                case picEnum.行書:
                    extName = ".jpg";
                    break;
                case picEnum.小篆:
                    break;
                case picEnum.甲骨文:
                    break;
                case picEnum.金文:
                    break;
                case picEnum.隸書:
                    break;
                default:
                    break;
            }
            doc.Application.Activate();            
            foreach (WinWord.Range item in sel.Characters)
            {
                Delay(Convert.ToInt32(numericUpDown1.Value * 1000));
                //wait();                
                string f = dir + item.Text + extName;
                if (pE == picEnum.小篆)
                {
                    f = getFullNameNTUswxz(dir, item.Text);
                }
                if (System.IO.File.Exists(f))
                {
                    if (!inserted)
                    {
                        inserted = true;
                    }
                    docApp.ScreenUpdating = false;
                    WinWord.WdColorIndex c = item.HighlightColorIndex;                    
                    sp = item.InlineShapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                        , Microsoft.Office.Core.MsoTriState.msoTrue, item);
                    sp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    //sp.Height = (float)0.9 * (15 + item.Font.Size - 12);
                    sp.Height = (float)1 * (15 + item.Font.Size - 12); 
                    sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                    sp.PictureFormat.TransparencyColor = 16777215;
                    sp.Range.HighlightColorIndex = c;
                    doc.ActiveWindow.ScrollIntoView(sp);
                    if (checkBox1.Checked != true)
                    {
                        //http://www.exceloffice.net/archives/3643
                        item.Font.Fill.Transparency = 1;
                        s = sp.ConvertToShape();
                        s.WrapFormat.Type = WinWord.WdWrapType.wdWrapFront;//文繞圖 文字在後
                        //https://social.msdn.microsoft.com/Forums/zh-TW/b6f28a4f-be91-4b67-9dfc-378a6809eeb0/22914203092103329992vba23559word2003272843504122294292553034037197?forum=232
                        //https://docs.microsoft.com/zh-tw/office/vba/api/word.wdwraptypemerged
                    }
                    if (docApp.Selection.Type != WinWord.WdSelectionType.wdSelectionIP)
                    {

                        item.SetRange(item.Start + 1, item.End);
                        item.Delete();
                    }
                    docApp.ScreenUpdating = true;
                }
            }
            if (inserted)
            {
                if (sel.ParagraphFormat.BaseLineAlignment != WinWord.WdBaselineAlignment.wdBaselineAlignCenter)
                {
                    sel.ParagraphFormat.BaseLineAlignment = WinWord.WdBaselineAlignment.wdBaselineAlignCenter;
                }
                sel.Collapse(WinWord.WdCollapseDirection.wdCollapseEnd);
            }
        }

        void runPPT(string dir, picEnum pE)
        {
            PowerPnt.Application oPPTapp = (PowerPnt.Application)getOffice(officeEnum.PowerPoint);
            PowerPnt.Presentation ppt = oPPTapp.ActivePresentation;
            PowerPnt.Selection sel = ppt.Application.ActiveWindow.Selection;
            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
            {
                ppt.Application.Activate();

                PowerPnt.Slide sld = ppt.Application.ActiveWindow.View.Slide;

                var tr = sel.TextRange2;
                foreach (Microsoft.Office.Core.TextRange2 item in tr.Characters)
                {
                    Delay(Convert.ToInt32(numericUpDown1.Value * 1000));
                    //wait();
                    string f = dir + item.Text + ".png";
                    if (pE == picEnum.行書)
                    {
                        f = dir + item.Text + ".jpg";
                    }

                    if (System.IO.File.Exists(f))
                    {
                        float lf = item.BoundLeft;
                        float tp = item.BoundTop;
                        float h = item.BoundHeight;
                        float w = item.BoundWidth;
                        PowerPnt.Shape sp = sld.Shapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                            , Microsoft.Office.Core.MsoTriState.msoTrue,
                            lf, tp, w, h);
                        spTransp(ref sp, item);
                    }
                }
            }
        }

        object getOffice(officeEnum ofE)
        {
            string CLSID = ""; object office;
            switch (ofE)
            {
                case officeEnum.PowerPoint:
                    CLSID = "PowerPoint.Application";
                    //機碼：HKEY_CLASSES_ROOT\PowerPoint.Application
                    break;
                case officeEnum.Word:
                    CLSID = "Word.Application";//HKEY_CLASSES_ROOT\Word.Application\CLSID
                    break;
                case officeEnum.Excel:
                    CLSID = "Excel.Application";
                    break;
                default:
                    break;
            }
            //https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance
            office = System.Runtime.InteropServices.Marshal.GetActiveObject(CLSID);
            return office;
        }

        string getDir(picEnum pE)
        {
            string subFolder = "";
            switch (pE)
            {
                case picEnum.卦圖64:
                    subFolder = "\\Macros\\64卦圖";
                    break;
                case picEnum.行書:
                    subFolder = "\\Macros\\古文字\\行書";
                    break;
                case picEnum.小篆:
                    subFolder = "\\Macros\\古文字\\台大說文小篆字圖";
                    break;
                case picEnum.甲骨文:
                    subFolder = "\\Macros\\古文字\\甲骨文";
                    break;
                case picEnum.金文:
                    subFolder = "\\Macros\\古文字\\金文";
                    break;
                case picEnum.隸書:
                    subFolder = "\\Macros\\古文字\\隸書";
                    break;
                default:
                    break;
            }
            string dir = "";
            List<string> dirs = new List<string> { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
            foreach (string item in dirs)
            {
                dir = item + ":\\@@@華語文工具及資料@@@" + subFolder;
                if (System.IO.Directory.Exists(dir))
                {

                    return dir;
                }
            }
            return "";
        }
        string getFullNameNTUswxz(string dir, string x)
        {
<<<<<<<<< Temporary merge branch 1
=========
            if (x=="/")
            {
                return "";
            }
            string s = dir.Substring(0, dir.IndexOf("古文字"));
            ADODB.Recordset rst = new ADODB.Recordset();
            ADODB.Connection cnt = new ADODB.Connection();
            cnt.Open("Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" +
                s + "\\說文資料庫原造字取代為系統字參照用.mdb");
            rst.Open("SELECT 台大說文小篆字圖卷數表.檔名, Format([卷],\"00\") " +
                "AS V FROM 台大說文小篆字圖卷數表 WHERE (((" +
                "InStr([檔名],\"" + x + "\"))>0))", cnt, ADODB.CursorTypeEnum.adOpenKeyset,
                ADODB.LockTypeEnum.adLockReadOnly);
            if (rst.RecordCount > 0)
            {
                return dir + "\\說文卷" + rst.Fields["V"].Value + "\\" + rst.Fields["檔名"].Value;
            }
            return "";
        }


        void spTransp(ref PowerPnt.Shape sp, Microsoft.Office.Core.TextRange2 tr)
        {
            sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
            sp.PictureFormat.TransparencyColor = 16777215; //Microsoft.VisualBasic.Information.RGB(255, 255, 255);
                                                           //if (checkBox1.Checked != true)
            tr.Font.Fill.Transparency = 1;
        }

        #region 毫秒延时 界面不会卡死
        //果然要完美解決卡頓的問題還是要藉由多執行緒代理的方法，未必也；蓋是用BackgroundWorker 類別比較對，詳部件篩選器實作 //https://docs.microsoft.com/zh-tw/dotnet/api/system.componentmodel.backgroundworker?view=netframework-4.0&f1url=%3FappId%3DDev15IDEF1%26l%3DZH-TW%26k%3Dk(System.ComponentModel.BackgroundWorker);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.0);k(DevLang-csharp)%26rd%3Dtrue 
        public static void Delay(int mm)
        {//https://my.oschina.net/u/4419355/blog/3452446            
            DateTime current = DateTime.Now;
            //Application.DoEvents();
            while (current.AddMilliseconds(mm) >= DateTime.Now)
            {
                Application.DoEvents();
            }
            return;
        }
        #endregion

        void wait()
        {
            //https://www.itread01.com/content/1547889140.html
            //https://www.google.com/search?q=c%23+%E4%B8%8D%E5%8D%A1%E6%AD%BB+&sxsrf=ALeKk018ozK2YgezqvoGdvu0dgRhsw77Gw%3A1616674700814&ei=jH9cYNmcMfuJr7wPur2m6Aw&oq=c%23+%E4%B8%8D%E5%8D%A1%E6%AD%BB+&gs_lcp=Cgdnd3Mtd2l6EAMyBQghEKABOgUIABCwAzoECCMQJzoKCAAQsQMQgwEQQzoECAAQQzoCCAA6BAgAEB46CAgAEAgQChAeOgYIABAIEB46CAgAELEDEIMBULnzC1i_sgxgo7QMaARwAHgAgAGqA4gB1QeSAQU4LjQtMZgBAKABAaoBB2d3cy13aXrIAQHAAQE&sclient=gws-wiz&ved=0ahUKEwjZkonKtsvvAhX7xIsBHbqeCc0Q4dUDCA0&uact=5

            decimal fl = numericUpDown1.Value;
            if (fl > 0)
            {
                //System.Threading.CountdownEvent w = new System.Threading.CountdownEvent(0);
                //w.Wait(Convert.ToInt32(1000 * fl));
                System.Threading.Thread.Sleep(Convert.ToInt32(1000 * fl));
            }
        }





        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if (numericUpDown1.Value < 0)
            {
                numericUpDown1.Value = 0;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.go();
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }
    }
    enum picEnum : byte
    {
        卦圖64, 行書, 小篆, 甲骨文, 金文, 隸書
    }

    enum officeEnum
    {
        PowerPoint, Word, Excel
    }
}
