using powerPnt=Microsoft.Office.Interop.PowerPoint;
using winWord=Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace CharacterConverttoCharacterPics
{
    public class App//負責取得應用程式相關之業務
    {
        //static winWord.Application appDoc;
        //static powerPnt.Application appPpt;
        static object appOb;static string appClassName;
        public App(app app)
        {
            //switch (app)
            //{
            //    case app.Word:
            //        break;
            //    case app.PowerPoint:

            //        break;
            //    default:
            //        break;
            //}
        }
        public static winWord.Application AppDoc
        {
            get {
                appClassName = "Word.Application";
                appOb = getApp(appClassName);
                if (appOb == null)
                {
                    return new winWord.Application();
                }
                return (winWord.Application)appOb;
            } 
        }
            public static powerPnt.Application AppPpt
        {
            get {
                appClassName = "PowerPoint.Application";
                appOb = getApp(appClassName);
                if (appOb == null)
                {
                    return new powerPnt.Application();
                }
                return (powerPnt.Application)appOb;
            } 
        }


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
