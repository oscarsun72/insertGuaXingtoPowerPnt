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
using System.Text.RegularExpressions;

namespace insertGuaXingtoPowerpnt
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        PowerPnt.Application pptApp;
        PowerPnt.Presentation ppt;
        PowerPnt.Selection sel;
        PowerPnt.Slide sld;
        WinWord.Application docApp;
        WinWord.Document doc;
        WinWord.Selection selDoc;
        WinWord.InlineShape inlSp;
        WinWord.Range rng;
        WinWord.WdSelectionType selDocType;

        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> lb = new List<string>{"64卦圖","行書","小篆","甲骨文"
                    ,"金文","隸書"};
            listBox1.DataSource = lb;
            List<string> lb2 = new List<string> { "PowerPoint", "Word", "Excel" };
            listBox2.DataSource = lb2;
            checkBox1.Enabled = false;//在上一行給定listBox2.DataSource值時就會觸發事件            

        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            go();
        }

        private void go()
        {
            listBox1.Enabled = false; listBox2.Enabled = false; numericUpDown1.Focus(); button1.Enabled = false; checkBox1.Enabled = false;
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
            listBox1.Enabled = true; listBox2.Enabled = true; button1.Enabled = true;
            if ((string)listBox2.SelectedValue == "Word")
            {
                checkBox1.Enabled = true;
            }
        }

        void guaXing(officeEnum ofE)
        {
            string dir = getDir(picEnum.卦圖64) + "\\";
            if (dir != "")
            {
                switch (ofE)
                {
                    case officeEnum.PowerPoint:
                        pptApp = (PowerPnt.Application)getOffice(ofE);
                        ppt = pptApp.ActivePresentation;
                        sel = pptApp.ActiveWindow.Selection;
                        selDocType = selDoc.Type;
                        sld = ppt.Application.ActiveWindow.View.Slide;
                        runPPTGuaXing(dir);
                        break;
                    case officeEnum.Word:
                        docApp = (WinWord.Application)getOffice(ofE);
                        doc = docApp.ActiveDocument;
                        selDoc = doc.ActiveWindow.Selection;
                        rng = selDoc.Range;
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
            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
            {
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
                    spTransp(sp, sel.TextRange2);
                    sel.Unselect();
                }

            }
            ppt.Application.Activate();
        }

        void runDOCGuaXing(string dir)//Word插入卦形
        {
            string f = dir + selDoc.Text + ".png";
            if (!System.IO.File.Exists(f) && rng.Characters.Count == 1)
            {
                rng.SetRange(rng.Start, rng.Characters[1].Next().End);
                f = dir + rng.Text + ".png";
            }
            if (System.IO.File.Exists(f))
            {
                docApp.ScreenUpdating = false;
                WinWord.WdColorIndex c = selDoc.Range.HighlightColorIndex;
                inlSp = selDoc.InlineShapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                    , Microsoft.Office.Core.MsoTriState.msoTrue);
                inlSp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                inlSp.Height = (float)0.9 * (15 + selDoc.Range.Font.Size - 12);
                spTransp(inlSp, selDoc.Range, checkBox1.Checked);
                inlSp.Range.HighlightColorIndex = c;
                if (selDoc.ParagraphFormat.BaseLineAlignment != WinWord.WdBaselineAlignment.wdBaselineAlignCenter)
                {
                    selDoc.ParagraphFormat.BaseLineAlignment = WinWord.WdBaselineAlignment.wdBaselineAlignCenter;
                }
                if (checkBox1.Checked != true)//shape
                {
                    inlSp.ConvertToShape().WrapFormat.Type = WinWord.WdWrapType.wdWrapFront;//文繞圖 文字在後
                }
                else
                {//inlineshape
                    if (selDoc.Type != WinWord.WdSelectionType.wdSelectionIP)
                    {
                        selDoc.Range.SetRange(selDoc.Start + 1, selDoc.End);
                        selDoc.Delete();
                    }
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
                        pptApp = (PowerPnt.Application)getOffice(ofE);
                        ppt = pptApp.ActivePresentation;
                        sel = pptApp.ActiveWindow.Selection;
                        sld = pptApp.ActiveWindow.View.Slide;
                        runPPT(dir, pE);
                        break;
                    case officeEnum.Word:
                        docApp = (WinWord.Application)getOffice(ofE);
                        doc = docApp.ActiveDocument;
                        selDoc = doc.ActiveWindow.Selection;
                        selDocType = selDoc.Type;
                        rng = selDoc.Range;
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
            string extName = ".png";
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
            if (selDoc.ParagraphFormat.BaseLineAlignment != WinWord.WdBaselineAlignment.wdBaselineAlignCenter)
            {
                selDoc.ParagraphFormat.BaseLineAlignment = WinWord.WdBaselineAlignment.wdBaselineAlignCenter;
            }

            if (selDoc.Information[WinWord.WdInformation.wdWithInTable])
            {//若在表格中
                foreach (WinWord.Cell c in selDoc.Cells)
                {
                    charBycharDoc(dir, pE, extName, c.Range);
                }
            }
            else
                charBycharDoc(dir, pE, extName,selDoc.Range);
            selDoc.Collapse(WinWord.WdCollapseDirection.wdCollapseEnd);
        }

        private void charBycharDoc(string dir, picEnum pE, string extName,
            WinWord.Range rng)
        {
            foreach (WinWord.Range item in rng.Characters)
            {
                if (!checkCharsValid(item.Text))
                {
                    continue;
                }
                Delay(Convert.ToInt32(numericUpDown1.Value * 1000));
                //wait();                
                string f = dir + item.Text + extName;
                if (pE == picEnum.小篆)
                {
                    f = getFullNameNTUswxz(dir, item.Text);
                }
                if (System.IO.File.Exists(f))
                {
                    docApp.ScreenUpdating = false;
                    WinWord.WdColorIndex c = item.HighlightColorIndex;
                    inlSp = item.InlineShapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                        , Microsoft.Office.Core.MsoTriState.msoTrue, item);
                    inlSp.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    //sp.Height = (float)0.9 * (15 + item.Font.Size - 12);
                    inlSp.Height = (float)1 * (15 + item.Font.Size - 12);
                    spTransp(inlSp, item, checkBox1.Checked);
                    inlSp.Range.HighlightColorIndex = c;
                    doc.ActiveWindow.ScrollIntoView(inlSp.Range);
                    if (checkBox1.Checked != true)
                    {
                        inlSp.ConvertToShape().WrapFormat.Type = WinWord.WdWrapType.wdWrapFront;//文繞圖 文字在後
                        //https://social.msdn.microsoft.com/Forums/zh-TW/b6f28a4f-be91-4b67-9dfc-378a6809eeb0/22914203092103329992vba23559word2003272843504122294292553034037197?forum=232
                        //https://docs.microsoft.com/zh-tw/office/vba/api/word.wdwraptypemerged
                    }
                    else//inlineShape
                    {
                        //if (item.Information[WinWord.WdInformation.wdWithInTable])
                        //{
                        //    item.SetRange(item.Start + 1, item.End);
                        //    item.Delete();
                        //}
                        //else { 
                        //if (docApp.Selection.Type != WinWord.WdSelectionType.wdSelectionIP)
                        if (selDocType != WinWord.WdSelectionType.wdSelectionIP)
                        {
                            item.SetRange(item.Start + 1, item.End);
                            item.Delete();
                        }
                        //}
                    }
                    docApp.ScreenUpdating = true;
                    //docApp.ScreenRefresh();//若有逐字展示的需求才需要此行 2021/3/27
                }
            }
        }

        void runPPT(string dir, picEnum pE)
        {
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
                    switch (pE)
                    {
                        case picEnum.卦圖64:
                            break;
                        case picEnum.行書:
                            f = dir + item.Text + ".jpg";
                            break;
                        case picEnum.小篆:
                            f = getFullNameNTUswxz(dir, item.Text);
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


                    if (System.IO.File.Exists(f))
                    {
                        float lf = item.BoundLeft;
                        float tp = item.BoundTop;
                        float h = item.BoundHeight;
                        float w = item.BoundWidth;
                        PowerPnt.Shape sp = sld.Shapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                            , Microsoft.Office.Core.MsoTriState.msoTrue,
                            lf, tp, w, h);
                        spTransp(sp, item);
                    }
                }
                sel.Unselect();
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
            //https://docs.microsoft.com/zh-tw/dotnet/standard/base-types/best-practices-strings
            //https://docs.microsoft.com/zh-tw/dotnet/standard/base-types/character-classes-in-regular-expressions
            //https://walterinuniverse.wordpress.com/2014/09/03/asp-net-c-%E5%88%A4%E6%96%B7%E5%AD%97%E4%B8%B2-%E6%98%AF%E5%90%A6%E7%94%B1%E8%8B%B1%E6%96%87%E8%88%87%E6%95%B8%E5%AD%97%E7%B5%84%E6%88%90/
            System.Text.RegularExpressions.Regex re = new System.Text.RegularExpressions.Regex("[^A-Za-z0-9() 　/]");
            if (!re.IsMatch(x))
                return "";
            if (x == "" || x == "/")
                return "";

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
                string v = rst.Fields["V"].Value, f = rst.Fields["檔名"].Value;
                rst.Close(); cnt.Close();
                return dir + "\\說文卷" + v + "\\" + f;
            }
            rst.Close(); cnt.Close();
            return "";
        }

        //http://www.exceloffice.net/archives/3643
        void spTransp(PowerPnt.Shape sp, Microsoft.Office.Core.TextRange2 tr)
        {//圖片、字型透明化
            /*
             * System.InvalidCastException
            HResult=0x80004002
            Message=無法將類型 'System.__ComObject' 的 COM 物件轉換為介面類型 'Microsoft.Office.Interop.Word.Range'。由於發生下列錯誤，介面 (IID 為 '{0002095E-0000-0000-C000-000000000046}') 之 COM 元件上的 QueryInterface 呼叫失敗而導致作業失敗: 不支援此種介面 (發生例外狀況於 HRESULT: 0x80004002 (E_NOINTERFACE))。
            所以必須用多載的方式，函式（方法）多載（重載）的需求也應運而生
             …… */
            sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
            sp.PictureFormat.TransparencyColor = 16777215; //Microsoft.VisualBasic.Information.RGB(255, 255, 255);
                                                           //if (checkBox1.Checked != true)
            tr.Font.Fill.Transparency = 1;

        }

        void spTransp(WinWord.InlineShape inlsp, WinWord.Range tr,
            bool inlineShpape = false)
        {//圖片、字型透明化 for MS Word
            if (!inlineShpape)
            {
                WinWord.Shape sp = inlsp.ConvertToShape();
                sp.PictureFormat.TransparencyColor = 16777215;
                sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
                tr.Font.Fill.Transparency = 1;
            }
            else
            {
                inlsp.PictureFormat.TransparencyColor = 16777215;
                inlsp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
            }
        }
        #region 毫秒延时 界面不会卡死
        //果然要完美解決卡頓的問題還是要藉由多執行緒代理的方法，未必也；蓋是用BackgroundWorker 類別比較對，詳部件篩選器實作 //https://docs.microsoft.com/zh-tw/dotnet/api/system.componentmodel.backgroundworker?view=netframework-4.0&f1url=%3FappId%3DDev15IDEF1%26l%3DZH-TW%26k%3Dk(System.ComponentModel.BackgroundWorker);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.0);k(DevLang-csharp)%26rd%3Dtrue 
        //https://my.oschina.net/u/4419355/blog/3452446            
        //public static void Delay(int mm)
        void Delay(int mm)
        {
            DateTime current = DateTime.Now;
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


        bool checkCharsValid(string character)
        {
            Regex r = new Regex("[\r\a ]");
            if (r.IsMatch(character))
            {
                return false;
            }
            return true;
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



        private void listBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            if ((string)listBox2.SelectedValue == "Word")
                checkBox1.Enabled = true;//inlineShape？
            else
            {
                checkBox1.Checked = false;//N/A inlineShape
                checkBox1.Enabled = false;
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
