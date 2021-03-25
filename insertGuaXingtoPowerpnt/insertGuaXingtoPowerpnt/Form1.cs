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
using WinWord= Microsoft.Office.Interop.Word;
using Excel= Microsoft.Office.Interop.Excel;

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
            List<string> lb2 = new List<string>{"PowerPoint","Word","Excel"};
            listBox2.DataSource = lb2;

        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            officeEnum ofE=officeEnum.PowerPoint;
            switch (listBox2.SelectedItem)
            {
                case "PowerPoint":
                    ofE = officeEnum.PowerPoint;
                    break;
                case "Word":
                    ofE = officeEnum.Word;
                    break;
                case "Excel":
                    ofE = officeEnum.Excel;
                    break;
                default:
                    break;
            }
            switch (listBox1.SelectedValue)
            {
                case "64卦圖":
                    guaXing(ofE);
                    break;
                case "行書":
                    GuWenZi(picEnum.行書,ofE);
                    break;
                case "小篆":
                    break;
                case "甲骨文":
                    break;
                case "金文":
                    break;
                case "隸書":
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
            PowerPnt.Application pptApp =(PowerPnt.Application)getOffice(officeEnum.PowerPoint);
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

        void GuWenZi(picEnum pE, officeEnum ofE)
        {
            string dir = getDir(pE) + "\\";
            if (dir != "")
            {
                switch (ofE)
                {
                    case officeEnum.PowerPoint:
                        runPPT(dir,pE);
                        break;
                    case officeEnum.Word:
                        break;
                    case officeEnum.Excel:
                        break;
                    default:
                        break;
                }

            }
        }

        void runPPT(string dir,picEnum pE)
        {
            PowerPnt.Application oPPTapp =(PowerPnt.Application)getOffice(officeEnum.PowerPoint);
            PowerPnt.Presentation ppt = oPPTapp.ActivePresentation;
            PowerPnt.Selection sel = ppt.Application.ActiveWindow.Selection;
            if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
            {
                ppt.Application.Activate();

                PowerPnt.Slide sld = ppt.Application.ActiveWindow.View.Slide;

                var tr = sel.TextRange2;
                foreach (Microsoft.Office.Core.TextRange2 item in tr.Characters)
                {

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

        void spTransp(ref PowerPnt.Shape sp, Microsoft.Office.Core.TextRange2 tr)
        {
            sp.PictureFormat.TransparentBackground = Microsoft.Office.Core.MsoTriState.msoTrue;
            sp.PictureFormat.TransparencyColor = 16777215; //Microsoft.VisualBasic.Information.RGB(255, 255, 255);
            tr.Font.Fill.Transparency = 1;
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
