using System.Collections.Generic;
using System.IO;
using System.Media;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class FontsOpsDoc
    {
        public static List<string> fontOkList//這是製作字圖參考用的，表示是否需要做此字型之字圖
        {

            get
            {
                List<string> fontoklist = new List<string>();
                StreamReader sr = new StreamReader(
                DirFiles.getFontOkList_txt().FullName);
                string fontname;
                while (!sr.EndOfStream)
                {
                    fontname = sr.ReadLine();                    
                    if (fontname != ""&& fontname.IndexOf("::")== -1)
                        fontoklist.Add(fontname);
                }
                return fontoklist;
            }
        }
        public static List<string> fontPicsList//這是給執行插圖用的，表示真有此字型字圖
        {

            get
            {
                List<string> fontoklist = new List<string>();
                StreamReader sr = new StreamReader(
                DirFiles.getFontOkList_txt().FullName);
                string fontname;
                while (!sr.EndOfStream)
                {
                    fontname = sr.ReadLine();
                    if (fontname.IndexOf("::") > -1) break;//以「::」記號作為選取中止
                    if (fontname != "")
                        fontoklist.Add(fontname);
                }
                return fontoklist;
            }
            /*get =>
            new List<string>{
                "標楷體", "新細明體", "微軟正黑體", "新細明體 (本文中文字型)", "+本文中文字型"
                , "細明體_HKSCS", "細明體", "細明體_HKSCS-ExtB", "細明體-ExtB",
                 "教育部隸書",

                "64卦圖", "行書",
                "小篆", "甲骨文", "金文", "隸書", "文鼎隸書B", "文鼎隸書DB", "文鼎隸書HKM", "文鼎隸書M",
                "華康行書體", "文鼎行楷L", "DFGGyoSho-W7", "DFPGyoSho-W7", "DFPOYoJun-W5", "DFPPenJi-W4",
                "文鼎魏碑B", "文鼎行楷碑體B", "文鼎鋼筆行楷M",

                "FangSong", "Adobe 仿宋 Std R", "文鼎仿宋B", "文鼎仿宋L",
                "教育部標準楷書", "Adobe 楷体 Std R", "KaiTi", "文鼎標準楷體ProM",
                "文鼎顏楷H", "文鼎顏楷U", "文鼎毛楷B", "文鼎毛楷EB", "文鼎毛楷H",
                "DFMinchoP-W5",
                "DFGothicP-W5",
                "DFGKanTeiRyu-W11", "文鼎古印體B",
                "文鼎雕刻體B", "DFKinBun-W3",
                "DFGFuun-W7",

                "華康行書體(P)", "DFPFuun-W7", "DFGyoSho-W7" //華康行書體(P)以下為沒必要做的

        };*/
        }
        internal static void removeNoFont(winWord.Document ThisDocument, string fontname)
        {
            foreach (winWord.Range a in ThisDocument.Characters)
            {
                if (a.Font.NameFarEast == fontname &&
                               a.Font.Name == fontname) {; }
                else
                    a.Delete();
            }
            ThisDocument.Save();
            warnings.playBeep();
            warnings.playSound();
        }



        void FontIterator()
        {
            foreach (string fnt in App.AppDoc.FontNames)
            {
                //if (fnt.ind "隸") || InStr(1, fnt, "li", vbTextCompare)) And InStr(1, fnt, "@", vbTextCompare) = 0 And InStr(1, fnt, "lian", vbTextCompare) = 0 And InStr(1, fnt, "Libre", vbTextCompare) = 0 And InStr(1, fnt, "Lith", vbTextCompare) = 0 And InStr(1, fnt, "Liber", vbTextCompare) = 0 And InStr(1, fnt, "light", vbTextCompare) = 0 And InStr(1, fnt, "Franklin", vbTextCompare) = 0 And InStr(1, fnt, "Italic", vbTextCompare) = 0 {
                //    ThisDocument.Range.Font.Name = fnt
                //    //Debug.Print fnt
                //    //Stop
                //}
            }
            warnings.playSound();
            SystemSounds.Beep.Play();
            //'Dim strFont As String
            //'Dim intResponse As Integer
            //'
            //'For Each strFont In FontNames
            //' intResponse = MsgBox(Prompt:=strFont, Buttons:=vbOKCancel)
            //' If intResponse = vbCancel Then Exit For
            //'Next strFont
        }


        void FontsListView(winWord.Document ThisDocument)
        {
            int fontCount, i = 0; string x, xp = "";
            fontCount = App.AppDoc.FontNames.Count;
            x = "\r\n" + ThisDocument.Paragraphs[1].Range.Text.Substring
                (0, ThisDocument.Paragraphs[1].Range.Text.Length - 1);
            for (int j = 2; j <= fontCount; j++)
                xp += x;
            ThisDocument.Range().InsertAfter(xp);
            foreach (string ft in App.AppDoc.FontNames)
            {
                i++;
                ThisDocument.Paragraphs[i].Range.Font.Name = ft;
            }
            //var e;
            //fontokList
            //For Each fnt In ThisDocument.Paragraphs
            //        For Each e In fontOk
            //            If e = fnt.Range.Font.NameFarEast Then fnt.Range.Delete
            //        Next e
            //Next fnt
            //playSound
            //Beep
        }

    }
}
