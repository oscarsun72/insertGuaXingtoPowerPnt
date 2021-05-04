using CharacterConverttoCharacterPics;
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using PowerPnt = Microsoft.Office.Interop.PowerPoint;
using WinWord = Microsoft.Office.Interop.Word;
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
        officeEnum officE;
        internal static picEnum PicE;
        ADODB.Connection cnt;
        ADODB.Recordset rst;


        private void Form1_Load(object sender, EventArgs e)
        {
            List<string> lb = FontsOpsDoc.fontPicsList;/* new List<string>{"64卦圖","行書",
                "小篆","甲骨文","金文","隸書","文鼎隸書B","文鼎隸書DB","文鼎隸書HKM","文鼎隸書M",

                "華康行書體","文鼎行楷L","DFGGyoSho-W7","DFPGyoSho-W7","文鼎魏碑B","文鼎行楷碑體B","文鼎鋼筆行楷M","DFPOYoJun-W5","DFPPenJi-W4",

                "FangSong","Adobe 仿宋 Std R","文鼎仿宋B","文鼎仿宋L",

                "教育部標準楷書","Adobe 楷体 Std R","KaiTi","文鼎標準楷體ProM",
                "文鼎顏楷H","文鼎顏楷U","文鼎毛楷B","文鼎毛楷EB","文鼎毛楷H",
                "DFMinchoP-W5",
                "DFGothicP-W5",
                "DFGKanTeiRyu-W11","文鼎古印體B",
                "文鼎雕刻體B","DFKinBun-W3",
                "DFGFuun-W7"};*/
            if (lb.Count > 0)
            {
                listBox1.DataSource = lb; listBox1.SetSelected(2, true);// 設定預設值為"行書";the zero-based index of the currently selected item in a ListBox. 
                listbox1itme = listBox1.Items;
                PicE = picEnum.行書;
            }
            List<string> lb2 = new List<string> { "PowerPoint", "Word", "Excel" };
            listBox2.DataSource = lb2;
            checkBox1.Enabled = false;//在上一行給定listBox2.DataSource值時就會觸發事件
            officE = officeEnum.PowerPoint;
        }

        private void listBox2_DoubleClick(object sender, EventArgs e)
        {
            this.go();
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            this.go();
        }
        private void go()
        {
            if (getOffice(officE) == null) return;
            //為免誤按，故使控制項失效
            listBox1.Enabled = false; listBox2.Enabled = false; numericUpDown1.Focus(); button1.Enabled = false; checkBox1.Enabled = false; button2.Enabled = false;
            switch (PicE)
            {
                case picEnum.卦圖64:// "64卦圖":
                    guaXing(officE);
                    break;
                case picEnum.卦形8:// 8卦圖及陽爻陰爻:
                    guaXing(officE);
                    break;
                default:
                    GuWenZi(PicE, officE);
                    break;
            }
            //執行完後恢復控制項狀態
            listBox1.Enabled = true; listBox2.Enabled = true; button1.Enabled = true; button2.Enabled = true;
            if ((string)listBox2.SelectedValue == "Word")
            {
                checkBox1.Enabled = true;
            }
            Activate();
        }

        void guaXing(officeEnum ofE)
        {
            string dir = "";
            switch (PicE)
            {
                case picEnum.卦圖64:
                    dir = DirFiles.getPicDir(picEnum.卦圖64) + "\\";
                    break;
                case picEnum.卦形8:
                    dir = DirFiles.getPicDir(picEnum.卦形8) + "\\";
                    break;
                default:
                    break;
            }
            if (dir != "")
            {
                switch (ofE)
                {
                    case officeEnum.PowerPoint:
                        pptApp = (PowerPnt.Application)getOffice(ofE);
                        ppt = pptApp.ActivePresentation;
                        sel = pptApp.ActiveWindow.Selection;
                        ppt.Application.Activate();
                        sld = ppt.Application.ActiveWindow.View.Slide;
                        if (PicE == picEnum.卦形8)
                        {
                            switch (sel.TextRange.Font.NameFarEast)
                            {
                                case "標楷體":
                                    dir += "楷體用\\";
                                    break;
                                default:
                                    dir += "細明體用\\";
                                    break;
                            }
                        }
                        runPPTGuaXing(dir);
                        break;
                    case officeEnum.Word:
                        docApp = (WinWord.Application)getOffice(ofE);
                        doc = docApp.ActiveDocument;
                        selDoc = doc.ActiveWindow.Selection;
                        selDocType = selDoc.Type;
                        rng = selDoc.Range;
                        if (PicE == picEnum.卦形8)
                        {
                            switch (selDoc.Font.NameFarEast)
                            {
                                case "標楷體":
                                    dir += "楷體用\\";
                                    break;
                                default:
                                    dir += "細明體用\\";
                                    break;
                            }
                        }
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
                if (rng.Characters[1].Next() == null)
                {
                    MessageBox.Show("請將插入點放在要插入卦形圖的卦名前位置！", "", MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
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
            string dir = DirFiles.getPicDir(pE) + "\\";
            if (dir != "")
            {
                switch (ofE)
                {
                    case officeEnum.PowerPoint:
                        pptApp = (PowerPnt.Application)getOffice(ofE);
                        pptApp.Activate();
                        ppt = pptApp.ActivePresentation;
                        //ppt.Application.Activate();
                        sel = pptApp.ActiveWindow.Selection;
                        sld = pptApp.ActiveWindow.View.Slide;
                        if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
                            if (sel.TextRange2.Characters.Count == 0)
                                sel.ShapeRange.TextFrame.TextRange.Select();
                        //選取數張投影片以逐張執行插入字圖功能
                        List<PowerPnt.Slide> slds = new List<PowerPnt.Slide>();
                        if (sel.Type == PowerPnt.PpSelectionType.ppSelectionSlides)
                            foreach (PowerPnt.Slide sld in sel.SlideRange)
                                slds.Add(sld);
                        else
                            slds.Add(sld);
                        runPPT(dir, pE, slds);
                        //runPPT(dir, pE);
                        break;
                    case officeEnum.Word:
                        docApp = (WinWord.Application)getOffice(ofE);
                        if (docApp.Documents.Count == 0)
                        { doc = docApp.Documents.Add(); doc.ActiveWindow.Visible = true; }
                        else doc = docApp.ActiveDocument;
                        selDoc = doc.ActiveWindow.Selection;
                        selDocType = selDoc.Type;
                        if (selDocType == WinWord.WdSelectionType.wdSelectionIP)
                        {
                            doc.Content.Select();
                        }
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
            doc.Application.Activate();
            if (selDoc.ParagraphFormat.BaseLineAlignment != WinWord.WdBaselineAlignment.wdBaselineAlignCenter)
            {
                selDoc.ParagraphFormat.BaseLineAlignment = WinWord.WdBaselineAlignment.wdBaselineAlignCenter;
            }

            if (selDoc.Information[WinWord.WdInformation.wdWithInTable])
            {//若在表格中
                foreach (WinWord.Cell c in selDoc.Cells)
                {
                    charBycharDoc(dir, pE, c.Range);
                }
            }
            else
                charBycharDoc(dir, pE, selDoc.Range);
            selDoc.Collapse(WinWord.WdCollapseDirection.wdCollapseEnd);
        }

        private void charBycharDoc(string dir, picEnum pE, WinWord.Range rng)
        {
            foreach (WinWord.Range item in rng.Characters)
            {
                if (!checkCharsValid(item.Text))
                {
                    continue;
                }
                Delay(Convert.ToInt32(numericUpDown1.Value * 1000));
                //wait();
                string f = picFullName(dir, pE, item.Text);
                if (System.IO.File.Exists(f))
                {
                    docApp.ScreenUpdating = false;
                    WinWord.WdColorIndex c = item.HighlightColorIndex;
                    inlSp = item.InlineShapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                        , Microsoft.Office.Core.MsoTriState.msoTrue, item);//照線上說明所說,item若未collapsed，則當可取代，然也未能被取代，與所說不同！2021/3/28
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
                        if (selDocType != WinWord.WdSelectionType.wdSelectionIP)
                        {
                            item.Characters[2].Delete();
                        }
                    }
                    docApp.ScreenUpdating = true;
                    //docApp.ScreenRefresh();//若有逐字展示的需求才需要此行 2021/3/27
                }
            }
        }

        void runPPT(string dir, picEnum pE, List<PowerPnt.Slide> slds)
        //void runPPT(string dir, picEnum pE)
        {
            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionNone) return;
            foreach (PowerPnt.Slide sld in slds)            {
                
                for (int i=1;i<=sld.Application.SlideShowWindows.Count;i++)
                {
                    if (sld.Application.SlideShowWindows[i].Active == MsoTriState.msoTrue)
                    {
                        sld.Parent.Windows[1].Activate();
                        sld.Application.ActiveWindow.ViewType = PowerPnt.PpViewType.ppViewNormal;
                        break;
                    }
                }
                if (slds.Count > 1 || sel.Type == PowerPnt.PpSelectionType.ppSelectionSlides)
                {
                    sld.Select();
                    bool hasTextFrame = false;
                    foreach (PowerPnt.Shape sp in sld.Shapes)
                    {
                        if (sp.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                        {
                            sp.Select();//改變 sel 值                            
                            //如果sel可以因Select方法而即時變動，即不用此行:sel = sel.Application.ActiveWindow.Selection;
                            hasTextFrame = true;
                            break;
                        }
                    }
                    if (hasTextFrame == false)
                    //if (sel.Type == PowerPnt.PpSelectionType.ppSelectionSlides)
                    {
                        MessageBox.Show("請先選取要處理的文字方塊！再執行……");
                        return;
                    }
                }
                if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText ||
                        sel.Type == PowerPnt.PpSelectionType.ppSelectionShapes &
                        sel.ShapeRange.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    //ppt.Application.Activate();                
                    //PowerPnt.Slide sld = ppt.Application.ActiveWindow.View.Slide;
                    runSlideShow(sld);
                    if (sel.ShapeRange.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                    {//有表格
                        /* PowerPnt.CellRange cr = (PowerPnt.CellRange)sel.ShapeRange;//轉型失敗，改用下方「.Selected」屬性來判斷應用
                        for (int i = 0; i < cr.Count; i++)
                        {
                            if (cr[i].Selected)
                            {
                                cr[i].Select();
                                PowerPnt.Selection s = cr.Application.ActiveWindow.Selection;
                                charBycharPpt(dir, pE, sld, s.TextRange2,
                                    true, s.ShapeRange.Table.Parent.left, s.ShapeRange.Table.Parent.top);
                            }
                        } */

                        PowerPnt.Table tb = sel.ShapeRange.Table;
                        int r = tb.Rows.Count;
                        int c = tb.Columns.Count;
                        List<PowerPnt.Cell> cels = new List<PowerPnt.Cell> { };//list容器初始化
                        for (int i = 1; i <= r; i++)//這是表格中所有儲存格都處理
                        {//https://docs.microsoft.com/zh-tw/office/vba/api/powerpoint.table.cell
                            for (int j = 1; j <= c; j++)
                            {//https://docs.microsoft.com/zh-tw/office/vba/api/powerpoint.cellrange?f1url=%3FappId%3DDev11IDEF1%26l%3Dzh-TW%26k%3Dk(vbapp10.chm627000);k(TargetFrameworkMoniker-Office.Version%3Dv16)%26rd%3Dtrue
                                if (tb.Cell(i, j).Selected)//判斷儲存格是否有被選取
                                {
                                    cels.Add(tb.Cell(i, j));//記下已選取的儲存格，以備用
                                }
                            }
                        }
                        foreach (PowerPnt.Cell item in cels)
                        {
                            item.Select();
                            PowerPnt.Selection selCell = tb.Application.ActiveWindow.Selection;
                            charBycharPpt(dir, pE, sld, selCell.TextRange2, true, tb.Parent.left, tb.Parent.top);
                        }

                    }
                    else//純文字方塊
                    {
                        Microsoft.Office.Core.TextRange2 tr = sel.TextRange2;
                        charBycharPpt(dir, pE, sld, tr);
                    }
                    sel.Unselect();

                }

            }
        }

        private void runSlideShow(PowerPnt.Slide sld)
        {
            if (checkBox2.Checked)
            {//http://www.exceloffice.net/archives/4127
             //執行後即播放投影片
                if (ppt == null)
                {
                    pptApp = (PowerPnt.Application)getOffice(officeEnum.PowerPoint);
                    ppt = pptApp.ActivePresentation;
                    ppt.Application.Activate();
                    sld = pptApp.ActiveWindow.View.Slide;
                }
                PowerPnt.SlideShowSettings oSSS = ppt.SlideShowSettings;
                PowerPnt.SlideShowWindow ssw = oSSS.Run();
                ssw.View.GotoSlide(sld.SlideIndex);
            }
            else ppt.Application.Activate();
        }

        private void charBycharPpt(string dir, picEnum pE, PowerPnt.Slide sld,
            Microsoft.Office.Core.TextRange2 tr,
            bool inTable = false, float tbLeft = 0, float tbTop = 0)
        {
            foreach (Microsoft.Office.Core.TextRange2 item in tr.Characters)
            {
                if (!checkCharsValid(item.Text))
                {
                    continue;
                }
                string f = picFullName(dir, pE, item.Text);
                Delay(Convert.ToInt32(numericUpDown1.Value * 1000));
                //wait();
                if (System.IO.File.Exists(f))
                {
                    float lf = item.BoundLeft;
                    float tp = item.BoundTop;
                    float h = item.BoundHeight;
                    float w = item.BoundWidth;
                    if (inTable)
                    {
                        lf += tbLeft; tp += tbTop;
                    }
                    PowerPnt.Shape sp = sld.Shapes.AddPicture(f, Microsoft.Office.Core.MsoTriState.msoFalse
                        , Microsoft.Office.Core.MsoTriState.msoTrue,
                        lf, tp, w, h);
                    spTransp(sp, item);
                }
            }
        }

        private string picFullName(string dir, picEnum pE, string itemText)
        {
            string f = dir + itemText + ".png";
            switch (pE)
            {
                //case picEnum.卦圖64:
                //    break;
                case picEnum.行書:
                    f = dir + itemText + ".jpg";
                    break;
                case picEnum.小篆:
                    f = DirFiles.getFullNameNTUswxz(dir, itemText);
                    break;
                default:
                    break;
            }
            return f;
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
            try
            {
                office = System.Runtime.InteropServices.Marshal.GetActiveObject(CLSID);
                if (office == null) throw new Exception("Null");//https://ithelp.ithome.com.tw/articles/10254045?sc=rss.qu
                return office;
            }
            catch
            {
                MessageBox.Show("沒有開啟" + CLSID.Substring(
                    0, CLSID.IndexOf(".")), "", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                //throw new Exception("Null");
                return null;
            }
        }


        #region ADODB
        /*ADODB參考：操作方法使用 ADO 和 Jet OLE DB 提供者尋找記錄: https://docs.microsoft.com/zh-tw/office/troubleshoot/access/find-record-by-ado-and-jet-ole-db-provider
         *與各種資料庫的連線字串:http://web12.ravs.ntct.edu.tw/know/show.asp?QUESTIONID=47 
         *Create an ADO connection string:https://docs.microsoft.com/zh-tw/office/vba/access/concepts/activex-data-objects/create-an-ado-connection-string
         *MSAccess 資料庫 ADO source provider*/
        string getFullNameNTUswxzADODB(string dir, string x)
        {
            //https://docs.microsoft.com/zh-tw/dotnet/standard/base-types/best-practices-strings
            //https://docs.microsoft.com/zh-tw/dotnet/standard/base-types/character-classes-in-regular-expressions
            //https://walterinuniverse.wordpress.com/2014/09/03/asp-net-c-%E5%88%A4%E6%96%B7%E5%AD%97%E4%B8%B2-%E6%98%AF%E5%90%A6%E7%94%B1%E8%8B%B1%E6%96%87%E8%88%87%E6%95%B8%E5%AD%97%E7%B5%84%E6%88%90/

            if (!checkCharsValid(x))
                return "";
            if (x == "")
                return "";

            string s = dir.Substring(0, dir.IndexOf("古文字"));
            if (cnt == null)
                cnt = new ADODB.Connection();
            //ADODB.ObjectStateEnum.adStateClosed
            if (cnt.State == 0) cnt.Open("Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source=" +
                    s + "\\說文資料庫原造字取代為系統字參照用.mdb");
            if (rst == null)
                rst = new ADODB.Recordset();
            rst.Open("SELECT 台大說文小篆字圖卷數表.檔名, Format([卷],\"00\") " +
                    "AS V FROM 台大說文小篆字圖卷數表 WHERE (((" +
                    "InStr([檔名],\"" + x + "\"))>0))", cnt, ADODB.CursorTypeEnum.adOpenKeyset,
                    ADODB.LockTypeEnum.adLockReadOnly);
            if (rst.RecordCount > 0)
            {
                string v = rst.Fields["V"].Value, f = rst.Fields["檔名"].Value;
                rst.Close();// cnt.Close();//改為欄位則應不用再開開關關了
                return dir + "\\說文卷" + v + "\\" + f;
            }
            rst.Close();// cnt.Close();//改為欄位則應不用再開開關關了
            return "";

        }
        #endregion
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
        {   //由[\r\a... 改成 [\r\r\a ...果然解決跑出刪不掉的半形空格了
            Regex r = new Regex("[\r\r\a A-Za-z0-9()　/]");
            //Regex re = new System.Text.RegularExpressions.Regex("[^A-Za-z0-9() 　/]");
            if (r.IsMatch(character))
            {
                return false;
            }
            return true;
        }

        void resetClearAllPicsandFontTranspSel(officeEnum ofE)
        {
            switch (ofE)//自判別是哪個Office應用程式要操作
            {
                case officeEnum.PowerPoint:
                    pptApp = (PowerPnt.Application)getOffice(ofE);
                    pptApp.Activate();
                    sld = pptApp.ActiveWindow.View.Slide;
                    PowerPnt.Shape sp;
                    List<PowerPnt.Shape> spFrame = new List<PowerPnt.Shape>();//記下包括要取消/清除掉的字圖所在文字方塊
                    pptApp.Activate();
                    sel = pptApp.ActiveWindow.Selection;
                    //PowerPnt.ShapeRange spRng;
                    if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText ||
                        sel.Type == PowerPnt.PpSelectionType.ppSelectionShapes)
                    {//如果有選取範圍則僅處理選取的範圍文字框內的字圖
                        ShapsOps<PowerPnt.ShapeRange> spRngPtt =
                             new ShapsOps<PowerPnt.ShapeRange>(sel.ShapeRange,
                             "PowerPnt.ShapeRange");
                        foreach (PowerPnt.Shape item in sel.SlideRange.Shapes)
                        {
                            if (item.HasTextFrame == MsoTriState.msoTrue)
                            {//因為是插入字圖，故要文字方塊才處理

                                if (spRngPtt.isShapeContainsMeShapeRng(item))
                                {
                                    spFrame.Add(item);
                                    //break; 不能break 有可能是跨文字方塊複選的
                                }
                            }
                        }
                    }//否則就清除在投影片內的全部字圖

                    //for (int i = 1; i <=  sld.Shapes.Count; i++)
                    for (int i = 1; i <= sld.Shapes.Range().Count; i++)
                    {
                        sp = sld.Shapes[i];
                        if (sp.Type == Microsoft.Office.Core.MsoShapeType.msoPicture &&
                            sp.Title == "" && sp.AlternativeText.Length < 2 &&
                            sp.ActionSettings[PowerPnt.PpMouseActivation.ppMouseClick]
                                .Hyperlink.Address == null)
                        {
                            if (spFrame.Count > 0)
                            {
                                foreach (PowerPnt.Shape item in spFrame)
                                {
                                    if (new ShapsOps<PowerPnt.Shape>(sp,
                                        "PowerPnt.Shape").
                                        isShapeContainsMeShape(item))
                                    {
                                        sp.Delete();
                                        i--; break;//sp被刪除後便不能再比對了
                                    }
                                }
                            }
                            else
                            {//沒有包含的文字框就清除投影片內的全部字圖（非字圖，應設定「AlternativeText」屬性以供識別。如前判斷式所陳列）
                                sp.Delete();
                                i--;
                            }
                        }
                    }

                    switch (sel.Type)
                    {//復原原來字型文字狀態
                        case PowerPnt.PpSelectionType.ppSelectionText:
                            if (sel.TextRange2.Characters.Count == 0)
                                sel.ShapeRange.TextFrame.TextRange.Select();
                            {
                                sel.ShapeRange.TextFrame2.TextRange.Font.Fill.Transparency = 0;
                                //sel.TextRange2.Select();
                                //sel.TextRange2.Font.Fill.Transparency = 0;
                            }
                            break;
                        case PowerPnt.PpSelectionType.ppSelectionShapes:
                            if (sel.ShapeRange.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                            {
                                sel.ShapeRange.TextFrame2.TextRange.Font.Fill.Transparency = 0;
                                //sel.TextRange2.Select();
                                //sel.TextRange2.Font.Fill.Transparency = 0;
                            }
                            break;
                        default:
                            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionNone ||
                                sel.Type == PowerPnt.PpSelectionType.ppSelectionSlides)
                            {


                                int shpCnt = sld.Shapes.Count;
                                if (shpCnt > 0)
                                {
                                    for (int i = 1; i <= shpCnt; i++)
                                    {
                                        //if (sld.Shapes[i].Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)//不是TextBox，卻有TextFrame(msoPlaceholder即有TextFrame)
                                        if (sld.Shapes[i].HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                                        {
                                            if (sld.Shapes[i].TextFrame2.TextRange.Font.Fill.Transparency != 0)
                                            {
                                                if (spFrame.Count > 0)
                                                {
                                                    foreach (PowerPnt.Shape item in spFrame)
                                                    {
                                                        if (item.Id == sld.Shapes[i].Id)
                                                        {
                                                            sld.Shapes[i].TextFrame2.TextRange.Font.Fill.Transparency = 0;
                                                            sld.Parent.Windows[1].Activate();
                                                            sld.Shapes[i].Select();
                                                            break;//20210429改良，回復就而不必再比對，繼續找投影片中其他的shape，看有沒有其他文字方塊涵蓋了執行前選取的字圖
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    //這裡不要 break; 讓所有文字變透明的都恢復不透明就好20210410
                                                    sld.Shapes[i].TextFrame2.TextRange.Font.Fill.Transparency = 0;
                                                    sld.Parent.Windows[1].Activate();
                                                    sld.Shapes[i].Select();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            break;
                    }
                    break;
                //如果是對象是MS Word（winword）
                case officeEnum.Word:
                    docApp = (WinWord.Application)getOffice(ofE);
                    selDoc = docApp.ActiveWindow.Selection;
                    selDocType = selDoc.Type;
                    if (selDocType == WinWord.WdSelectionType.wdSelectionIP)
                    {
                        selDoc.Document.Content.Select();
                    }
                    rng = selDoc.Range;
                    docApp.Activate();
                    while (rng.InlineShapes.Count > 0)
                        rng.InlineShapes[1].Delete();
                    if (rng.ShapeRange.Count > 0)
                    { rng.ShapeRange.Select(); selDoc.Delete(); } //rng.ShapeRange[1].Delete();
                    rng.Font.Fill.Transparency = 0;
                    break;
                case officeEnum.Excel:
                    break;
                default:
                    break;
            }

            #region clearAllPics

            #endregion
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
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    this.Close();
                    break;
                case Keys.R:
                    resetClearAllPicsandFontTranspSel(officE);
                    break;
                case Keys.Enter:
                    go();
                    break;
                default:
                    break;
            }
        }



        private void listBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            switch ((string)listBox2.SelectedValue)
            {
                case "PowerPoint":
                    officE = officeEnum.PowerPoint;
                    checkBox1.Checked = false;//N/A inlineShape
                    checkBox1.Enabled = false;
                    checkBox2.Enabled = true;
                    break;
                case "Word":
                    officE = officeEnum.Word;
                    checkBox1.Enabled = true;//inlineShape？
                    checkBox2.Enabled = false;
                    break;
                case "Excel":
                    officE = officeEnum.Excel;
                    checkBox1.Checked = false;//N/A inlineShape
                    checkBox1.Enabled = false;
                    checkBox2.Enabled = false;
                    break;
                default:
                    break;
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            resetClearAllPicsandFontTranspSel(officE);
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            runSlideShow(sld);
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {//設定欄位picE的值
         //picE = (picEnum)listBox1.SelectedIndex;//https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.listbox.selectedindex?view=net-5.0
         //https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.listbox.selectedindex?view=net-5.0
            switch (listBox1.SelectedValue)
            {
                case "64卦圖":
                    PicE = picEnum.卦圖64;
                    break;
                case "8卦圖":
                    PicE = picEnum.卦形8;
                    break;
                case "行書":
                    PicE = picEnum.行書;
                    break;
                case "小篆":
                    PicE = picEnum.小篆;
                    break;
                default:
                    PicE = picEnum.Default;
                    break;
            }

            showFontPreview();
            showToolTip();
        }

        private void showToolTip()
        {//https://bit.ly/3gs04Ya        
         //https://bit.ly/32tdqLF
            ToolTip ttp = new ToolTip();
            string listBox1SelectedItem = listBox1.SelectedItem.ToString();
            Regex rx = new Regex("[a-zA-Z0-9]");
            if (rx.IsMatch(listBox1SelectedItem.Substring(0, 1)))
            {
                if (listBox1SelectedItem.Length > 11)
                    ttp.SetToolTip(listBox1, listBox1.SelectedItem.ToString());
                else
                    ttp.SetToolTip(listBox1, "");
            }
            else
            {
                if (listBox1SelectedItem.Length > 7)
                    ttp.SetToolTip(listBox1, listBox1.SelectedItem.ToString());
                else
                    ttp.SetToolTip(listBox1, "");
            }
        }

        private void showFontPreview()
        {
            string ext = "png"; string w = textBox1.Text == "打個字看看：" ? "真" : textBox1.Text;
            string dir = DirFiles.getPicDir(PicE);
            string picsFullname;
            if (PicE == picEnum.行書)
                ext = "jpg";
            picsFullname = dir + "\\" + w + "." + ext;
            if (PicE == picEnum.小篆)
                picsFullname = DirFiles.getFullNameNTUswxz(dir, w);

            if (System.IO.File.Exists(picsFullname))
            {
                Bitmap pic = new Bitmap(picsFullname);//https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.picturebox.image?view=netframework-4.6.1&f1url=%3FappId%3DDev16IDEF1%26l%3DEN-US%26k%3Dk(System.Windows.Forms.PictureBox.Image);k(TargetFrameworkMoniker-.NETFramework,Version%253Dv4.6.1);k(DevLang-csharp)%26rd%3Dtrue
                pictureBox1.SizeMode = PictureBoxSizeMode.Zoom;//https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.pictureboxsizemode?view=net-5.0
                pictureBox1.Image = pic;
            }
            else
                pictureBox1.Image = null;//https://stackoverflow.com/questions/5856196/clear-image-on-picturebox //https://www.codeproject.com/Questions/1205981/How-to-reset-the-image-in-a-picture-box-in-Csharp
            //throw new NotImplementedException();
        }



        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Process prc = new Process();
            prc.StartInfo.FileName = DirFiles.getPicDir(PicE);//開啟古文字內的各字型字圖存放的資料夾20210419
            if (Directory.Exists(prc.StartInfo.FileName))
            {
                prc.Start();
            }
            else
            {
                MessageBox.Show("資料夾不存在，請檢查！", "",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                prc.StartInfo.FileName = DirFiles.PicsRootFolder;
                if (Directory.Exists(prc.StartInfo.FileName))
                {
                    prc.Start();
                }
            }
        }

        private void listBox1_MouseHover(object sender, EventArgs e)
        {
        }

        ListBox.ObjectCollection listbox1itme;
        ListBox.ObjectCollection Listbox1Itme { get => listBox1.Items; }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {//篩選字型
            string text = textBox2.Text;
            if (text == "")
            {
                listBox1.DataSource = listbox1itme;
                return;
            }
            List<string> ls = new List<string>();
            foreach (string item in listbox1itme)
            {/* c# indexof 不分大小寫
              * http://ezbo.blogspot.com/2012/05/c-stringindexof.html
                https://blog.csdn.net/amohan/article/details/12649533
                */
                if (item.IndexOf(text, StringComparison.CurrentCultureIgnoreCase) > -1)
                {
                    ls.Add(item);
                }
            }
            listBox1.DataSource = ls;
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            if (textBox2.Text != "") textBox2.Text = "";
        }

        private void textBox1_Enter(object sender, EventArgs e)
        {

        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            if (textBox1.Text == "") textBox1.Text = "打個字看看：";
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string x = textBox1.Text;
            if (string.IsNullOrEmpty(x) || string.IsNullOrWhiteSpace(x)) return;
            Regex rg = new Regex("[0-9a-zA-Z]");
            if (rg.IsMatch(x)) return;
            StringInfo si = new StringInfo(textBox1.Text);
            if (si.LengthInTextElements == 1)
            {
                string picFilefullname = new PicsOps().getPicFullname(
                    si.String);
                if (File.Exists(picFilefullname))
                {
                    pictureBox1.Image = new Bitmap(picFilefullname);
                }
            }
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }
        bool doNotEntered = false;
        private void listBox1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (ModifierKeys)
            {
                case Keys.Shift:
                    break;
                case Keys.Control:
                    if (listBox1.SelectedItems.Count > 0)
                        if (e.KeyCode == Keys.C)
                        {
                            Clipboard.SetText
                                  (listBox1.SelectedItem.ToString());
                            doNotEntered = true;
                        }
                    break;
                case Keys.Alt:
                    break;
                default:
                    break;
            }
        }

        private void listBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (doNotEntered)
            {
                e.Handled = true; doNotEntered = false;
            }
        }
    }

    enum officeEnum
    {
        PowerPoint, Word, Excel
    }
}
