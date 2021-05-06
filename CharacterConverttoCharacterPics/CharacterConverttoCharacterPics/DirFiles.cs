using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using powerPnt = Microsoft.Office.Interop.PowerPoint;

namespace CharacterConverttoCharacterPics
{
    public class DirFiles
    {//以後目錄、路徑均要取得最後的反斜線
        internal static string getDirRoot
        {//https://www.google.com/search?q=c%23+%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&rlz=1C1JRYI_enTW948TW948&oq=%E5%8F%96%E5%BE%97%E5%B0%88%E6%A1%88%E8%B7%AF%E5%BE%91&aqs=chrome.1.69i57j0i5i30l2.7266j0j7&sourceid=chrome&ie=UTF-8
            get =>
                new DirectoryInfo(
                System.AppDomain.CurrentDomain.BaseDirectory)
                .Parent.Parent.Parent.Parent.FullName + "\\";
        }

        internal static FileInfo getCjk_basic_IDS_UCS_Basic_txt()
        {
            DirectoryInfo dirRoot = new DirectoryInfo(getDirRoot);
            IEnumerable<FileInfo> fileList = dirRoot.GetFiles
                ("*.txt", SearchOption.AllDirectories);
            IEnumerable<FileInfo> fileQuery =
                from file in fileList
                where file.Name.IndexOf("cjk-basic-IDS-UCS-Basic.txt") > -1
                select file;
            if (fileQuery.Count() > 0)
                return fileQuery.First();
            else
                return null;
        }

        internal static async void appendFontOkList_txt(List<string> appendTextList,
            List<string> fontoklist)
        {//https://docs.microsoft.com/zh-tw/dotnet/csharp/programming-guide/file-system/how-to-write-to-a-text-file
            using (StreamWriter file = new StreamWriter
                (getFontOkList_txt().FullName, append: true))
            {
                foreach (string item in appendTextList)
                {
                    if (!fontoklist.Contains(item))
                    {
                        await file.WriteLineAsync(item);
                    }
                }
            }
        }

        internal static FileInfo getFontOkList_txt()
        {
            //先求方便了，否則一下要兼顧太多檔案20210426
            return new FileInfo(@"G:\我的雲端硬碟\programming程式設計開發\fontOkList.txt");
            /*
            DirectoryInfo dirRoot = new DirectoryInfo(getDirRoot);
            IEnumerable<FileInfo> fileList = dirRoot.GetFiles
                ("*.txt", SearchOption.AllDirectories);
            IEnumerable<FileInfo> fileQuery =
                from file in fileList
                where file.Name.IndexOf("fontOkList.txt") > -1
                select file;
            if (fileQuery.Count() > 0)
                return fileQuery.First();
            else
                return null;
            */
        }
        internal static string getDir各字型檔相關()
        {
            return getCjk_basic_IDS_UCS_Basic_txt().DirectoryName;
        }

        //internal powerPnt.Presentation get字圖母片pptm()
        internal powerPnt.Presentation get字圖母片pptm()
        {
            try
            {
                App app = new App();
                powerPnt.Application pptApp = app.AppPpt;
                foreach (powerPnt.Presentation ppt in pptApp.Presentations)
                {
                    if (ppt.Name == "字圖母片.pptm")
                    {
                        return ppt;
                    }
                }
                return pptApp.Presentations.Open(
                    getDirRoot + "字圖母片.pptm");
            }
            catch (System.Exception)
            {
                Application.DoEvents();
                //App app = new App()
                //{
                //    AppPpt = null
                //};

                return new App(app.PowerPoint)
                    .AppPpt.Presentations.Open(
                    getDirRoot + "字圖母片.pptm");
            }

        }

        internal static void getPicFolder(string picFolderPath)
        {
            if (Directory.Exists(picFolderPath) == false)
            {
                Directory.CreateDirectory(picFolderPath);
            }
        }
        public static void openFolder(string picDir)
        {
            //Process.Start(picDir);//Shell "explorer " & pth, vbMaximizedFocus;
            //開啟資料夾：https://happyduck1020.pixnet.net/blog/post/34382453-c%23-%E9%96%8B%E5%95%9F%E8%B3%87%E6%96%99%E5%A4%BE
            System.Diagnostics.Process prc = new System.Diagnostics.Process();
            prc.StartInfo.FileName = picDir;
            prc.Start();
            Application.DoEvents();
            warnings.playBeep();
        }

        public static string searchRootDirChange(string dir)
        {
            if (Directory.Exists(dir) == false)
            {
                string newDir;
                DriveInfo[] di = DriveInfo.GetDrives();//https://bit.ly/3mYEqw0
                foreach (DriveInfo item in di)
                {
                    newDir = dir.Replace(Path.GetPathRoot(dir), item.Name);
                    if (Directory.Exists(newDir))
                        return newDir;
                }
                return "";
            }
            return dir;

        }
        //將指定資料夾包成同名壓縮檔zip
        internal static void zipFolderFiles(string dir)
        {
            if (Directory.Exists(dir) == false) return;
            DirectoryInfo di = new DirectoryInfo(dir);
            string fZip = di.Parent.FullName + "\\" + di.Name + ".zip";
            if (File.Exists(fZip)) File.Delete(fZip);
            ZipFile.CreateFromDirectory(dir, fZip,
                CompressionLevel.NoCompression, true
                );
        }
    }
}
