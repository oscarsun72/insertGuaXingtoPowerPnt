using System.Runtime.InteropServices;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class App//負責取得應用程式相關之業務
    {
        //static 表示若未加設定為null則本應用程式還開啟時，其生命週期就一直延續著20210419
        static winWord.Application appDoc;
        static powerPnt.Application appPpt;
        static object appOb; static string appClassName;

        static bool pptAppOpenbyCode = false;
        static bool docAppOpenbyCode = false;
        //public App(app app)
        //{
        //    //switch (app)
        //    //{
        //    //    case app.Word:
        //    //        break;
        //    //    case app.PowerPoint:

        //    //        break;
        //    //    default:
        //    //        break;
        //    //}
        //}
        public static winWord.Application AppDoc
        {
            get
            {
                if (appDoc == null)
                {
                    appClassName = "Word.Application";
                    appOb = getApp(appClassName);
                    if (appOb == null)
                    {
                        docAppOpenbyCode = true;
                        appDoc = new winWord.Application();
                        return appDoc;
                    }
                    docAppOpenbyCode = false;
                    appDoc = (winWord.Application)appOb;
                    return appDoc;
                }
                return appDoc;
            }
        }
        public static powerPnt.Application AppPpt
        {
            get
            {
                if (appPpt== null)
                {
                    appClassName = "PowerPoint.Application";
                    appOb = getApp(appClassName);
                    if (appOb == null)
                    {
                        pptAppOpenbyCode = true;//不如此則由程式啟動的powerpoint
                                                //似乎無法以使用者手動關閉20210419
                        appPpt = new powerPnt.Application(); 
                        return appPpt;
                    }
                    pptAppOpenbyCode = false;
                    appPpt = (powerPnt.Application)appOb;
                    return appPpt;
                }
                return appPpt;
            }
            set { appOb = value; }
        }
        public static bool PptAppOpenByCode { get => pptAppOpenbyCode; }
        public static bool DocAppOpenByCode { get => docAppOpenbyCode; }
        static object getApp(string appClassName)
        {
            try
            {
                return Marshal.GetActiveObject(appClassName);
            }
            catch (global::System.Exception)
            {
                return null;
                //throw;
            }

        }
    }
    public enum app : byte
    {
        Word, PowerPoint
    }
}
