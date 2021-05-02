using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using powerPnt = Microsoft.Office.Interop.PowerPoint;
using WinWord = Microsoft.Office.Interop.Word;

namespace CharacterConverttoCharacterPics
{
    public partial class Form1 : Form
    {
        private bool fontsOK = false;

        //static PowerPnt.Application pptApp; 
        //static  WinWord.Application wdApp;
        //PowerPnt.Presentation ppt;
        //WinWord.Document doc; 
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


            //textBox1.SpecialEffect Access才有此屬性：https://docs.microsoft.com/zh-tw/office/vba/api/access.textbox.specialeffect
            //c# - 如何使RichTextBox外观平整？https://www.coder.work/article/953103
            /*这确实是一种hack，但是您可以做的一件事是将Panel控件拖放到页面上。给它设置一个FixedSingle的BorderStyle(默认情况下为None)。
                将RichTextBox放到面板中，并将BorderStyle设置为none。然后将RichTextBox的Dock属性设置为Fill。
                这将为您提供带有扁平边框的RichTextBox。*/

            //richTextBox1.t.sp

            //if (ppnt==null||wwrd==null)
            //{
            //    MessageBox.Show("請開啟 Word 與 PowerPoint 再繼續");                
            //}
            //else
            //{
            //    pptApp = (PowerPnt.Application)ppnt;
            //    wdApp = (WinWord.Application)wwrd;                
            //    if (wdApp.Documents.Count > 0)
            //    {
            //        doc = wdApp.ActiveDocument;
            //        textBox1.Text = doc.FullName;
            //    }
            //    ppt = pptApp.ActivePresentation;
            //}

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.T://test測試用

                    break;
                case Keys.Escape:
                    this.Close();
                    break;
                case Keys.Enter:
                    goFontsCharsToPics();
                    break;
                default:
                    break;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            goFontsCharsToPics();
            Refresh();
        }

        private void goFontsCharsToPics()
        {
            try
            {
                string chDir = DirFiles.searchRootDirChange(textBox1.Text);
                if (chDir == "")
                {
                    MessageBox.Show("並無此目錄,請確認後再執行！", "",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else if (textBox1.Text != chDir) textBox1.Text = chDir;

                BackColor = Color.Gray;
                this.Enabled = false; button1.Enabled = false;
                string fontname = textBox2.Text;
                WinWord.Document wd = new fontsPics().getFontCharacterset
                    (fontname);
                if (wd != null)
                {
                    BackColor = Color.Red;
                    string picFolder = textBox1.Text;
                    if (picFolder.IndexOf(fontname) == -1)
                    { picFolder += ("\\" + fontname + "\\"); }
                    powerPnt.Presentation ppt =
                        new fontsPics().prepareFontPPT(fontname,
                        float.Parse(textBox3.Text));
                    new fontsPics().addCharsSlidesExportPng(wd, ppt, picFolder,
                        Int32.Parse(textBox4.Text));
                    FileInfo wdf = new FileInfo(wd.FullName);
                    WinWord.Application wdApp = wd.Application;
                    wd.Close();//執行完成，關閉因程式而開啟的Word
                    if (!wdApp.UserControl) wdApp.Quit();   
                    //移動已完成的檔案到已完成資料夾下
                    string destFilename = wdf.Directory.FullName
                                    + "\\done已完成\\" + wdf.Name;
                    if (!File.Exists(destFilename)) wdf.MoveTo(destFilename);
                    if (BackColor != Color.BurlyWood)//若字圖與字型字數無不同，才顯示綠底色
                        BackColor = Color.Green;
                    warnings.playSound();
                }
                this.Enabled = true; button1.Enabled = true;
            }
            catch (Exception e)
            {
                WinWord.Document d = new WinWord.Application().Documents.Add();
                d.Range().Text = e.Message + "\n\r\n\r" +
                    e.Data + "\n\r\n\r" + e.Data +
                    "\n\r\n\r" + e.Source + "\n\r\n\r" +
                    e.HelpLink + "\n\r\n\r" + e.HResult + "\n\r\n\r" +
                    e.InnerException + "\n\r\n\r" +
                    e.StackTrace + "\n\r\n\r" +
                    e.TargetSite + "\n\r\n\r" + e.ToString();
                d.ActiveWindow.Visible = true;
            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.Text = Clipboard.GetText();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (FontsOpsDoc.fontOkList.Contains(textBox2.Text))
            {
                fontsOK = true;
                MessageBox.Show("這個字型已經做過了！或是不必做的", "請檢查！",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                button1.Enabled = false;
                return;
            }
            else
                fontsOK = false;
            button1.Enabled = true;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = Clipboard.GetText();

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string dirPath = textBox1.Text;
            if (Directory.Exists(dirPath))
            {//開啟資料夾：
                Process prc = new Process();
                prc.StartInfo.FileName = dirPath;
                prc.Start();
            }
        }
        static bool close字圖母片 = true;
        internal static bool Close字圖母片 { get => close字圖母片; set => close字圖母片 = value; }
        private void textBox2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                BackColor = Color.DarkOrange;
                Form2SelFont f2 = new Form2SelFont();
                f2.Show();
                WindowState = FormWindowState.Minimized;
                //C# 取消滑鼠事件 handled: https://docs.microsoft.com/zh-tw/dotnet/api/system.windows.forms.handledmouseeventargs?view=net-5.0
                //未成功，再研究20210425
                new HandledMouseEventArgs(
                    e.Button, e.Clicks, e.X, e.Y, e.Delta)
                {
                    Handled = true
                };
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            /*改動態後當不必再用下列方式中止
            //App app = new App();
            //if (app.AppPpt != null && app.PptAppOpenByCode == true)
            //{
            //    try
            //    { app.AppPpt.Quit(); }
            //    catch
            //    { app.AppPpt = null; }
            //}
            */
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            App app = new App();
            powerPnt.Application pptApp = app.AppPpt;
            if (pptApp.Presentations.Count > 0)
            {
                powerPnt.Presentation ppt = pptApp.ActivePresentation;
                if (ppt.Name == "字圖母片.pptm")
                {
                    //若字圖母片檔有打開，即設定好要轉字圖之字型規格，就將其字型名稱與大小載入相關的條件框中
                    textBox2.Text = ppt.Slides[2].Shapes[1].
                        TextFrame.TextRange.Font.NameFarEast;
                    textBox3.Text = ppt.Slides[2].Shapes[1].
                        TextFrame.TextRange.Font.Size.ToString();
                    if (!fontsOK && close字圖母片)
                        ppt.Close();//不帶參數，不會問你存不存檔，直接不存檔就離開。故若有存檔需要，必須先儲存才行
                    else
                        close字圖母片 = true;
                }
                else
                {//如此則可以在一般投影片檢視下，先選擇想要的字型，再自己開啟字圖母片來訂製模板，以供程式參照製作20210423
                    textBox2.Text = ppt.Slides[2].Shapes[1].
                        TextFrame.TextRange.Font.NameFarEast;
                }
            }
            if (app.PptAppOpenByCode)
                pptApp.Quit(); pptApp = null;
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //fontOKList中有而還要做此字型時（如遇當掉時須重作者）
            button1.Enabled = true;
        }
    }
}
