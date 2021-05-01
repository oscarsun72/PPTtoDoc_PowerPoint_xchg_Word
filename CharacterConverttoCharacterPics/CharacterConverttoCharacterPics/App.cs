using System.Runtime.InteropServices;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using winWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public class App//負責取得應用程式相關之業務
    {
        //static 表示若未加設定為null則本應用程式還開啟時，其生命週期就一直延續著20210419
        winWord.Application appDoc;//用靜態的（static ）一直會當，改用動態的，每次使用，即一個執行個體，用完即丟
        powerPnt.Application appPpt;
        object appOb; string appClassName;

        bool pptAppOpenbyCode = false;//個人管人個開的程式，故不應是靜態的
        bool docAppOpenbyCode = false;
        public App(app app=app.Default)
        {
            switch (app)
            {
                case app.Word:
                    appDoc = null;
                    break;
                case app.PowerPoint:
                    appPpt = null;
                    break;
                default:
                    appOb = null;
                    appDoc = null;
                    appPpt = null;
                    break;
            }
        }
        public winWord.Application AppDoc
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
            set { AppDoc= value; appOb = value; }
        }
        public  powerPnt.Application AppPpt
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
            set { appOb = value;appPpt = value; }
        }
        public bool PptAppOpenByCode { get => pptAppOpenbyCode; set => pptAppOpenbyCode=value; }
        public bool DocAppOpenByCode { get => docAppOpenbyCode; set => pptAppOpenbyCode = value; }
        object getApp(string appClassName)
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
        Default,Word, PowerPoint
    }
}
