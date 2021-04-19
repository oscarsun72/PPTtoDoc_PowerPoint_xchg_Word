using System.Runtime.InteropServices;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class App//負責取得應用程式相關之業務
    {
        //static winWord.Application appDoc;
        //static powerPnt.Application appPpt;
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
                appClassName = "Word.Application";
                appOb = getApp(appClassName);
                if (appOb == null)
                {
                    docAppOpenbyCode = true;
                    return new winWord.Application();
                }
                docAppOpenbyCode = false;
                return (winWord.Application)appOb;
            }
        }
        public static powerPnt.Application AppPpt
        {
            get
            {
                appClassName = "PowerPoint.Application";
                appOb = getApp(appClassName);
                if (appOb == null)
                {
                    pptAppOpenbyCode = true;//不如此則由程式啟動的powerpoint
                                            //似乎無法以使用者手動關閉20210419
                    return new powerPnt.Application();
                }
                pptAppOpenbyCode = false;
                return (powerPnt.Application)appOb;
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
