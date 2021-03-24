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

        }

        private void Form1_MouseClick(object sender, MouseEventArgs e)
        {
            string dir = getDir() + "\\";
            if (dir != "")
            {
                PowerPnt.Presentation ppt = getPPT();
                PowerPnt.Selection sel = ppt.Application.ActiveWindow.Selection;
                if (sel.Type == PowerPnt.PpSelectionType.ppSelectionText)
                {
                    PowerPnt.Slide sld = ppt.Application.ActiveWindow.View.Slide;
                    string f = dir + sel.TextRange.Text + ".png";
                    if (System.IO.File.Exists(f))
                    {
                        PowerPnt.TextRange tr= sel.TextRange.InsertAfter(sel.TextRange.Text);
                        tr.Select();
                        sel = sel.Application.ActiveWindow.Selection;
                        float lf = sel.TextRange.BoundLeft;
                        float tp = sel.TextRange.BoundTop;
                        int selCt = sel.TextRange2.Characters.Count;
                        float h = sel.TextRange.BoundHeight;
                        if (sel.ShapeRange.TextFrame.Orientation==Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationVerticalFarEast)
                        {
                            h = sel.TextRange.BoundHeight / selCt;
                        }
                        float w = sel.TextRange.BoundWidth;                        
                        if (sel.ShapeRange.TextFrame.Orientation==Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal)
                        {
                            w= sel.TextRange.BoundWidth/selCt  ;
                        }
                        PowerPnt.Shape spr = sld.Shapes.AddPicture(f,Microsoft.Office.Core.MsoTriState.msoFalse
                            ,Microsoft.Office.Core.MsoTriState.msoTrue,
                            lf,tp,w,h);
                        sel.TextRange.Text = "　";
                        
                    }

                }
                ppt.Application.Activate();
            }
        }

        PowerPnt.Presentation getPPT()
        {
            //https://docs.microsoft.com/en-us/previous-versions/office/troubleshoot/office-developer/use-visual-c-automate-run-program-instance
            PowerPnt.Application oPPntApp =
                (PowerPnt.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            //機碼：HKEY_CLASSES_ROOT\PowerPoint.Application
            PowerPnt.Presentation ppt = oPPntApp.ActivePresentation;

            return ppt;
        }

        string getDir()
        {
            string dir = "";
            List<string> dirs = new List<string> { "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z" };
            foreach (string item in dirs)
            {
                dir = item + ":\\@@@華語文工具及資料@@@\\Macros\\64卦圖";
                if (System.IO.Directory.Exists(dir))
                {

                    return dir;
                }
            }
            return "";
        }
    }
}
