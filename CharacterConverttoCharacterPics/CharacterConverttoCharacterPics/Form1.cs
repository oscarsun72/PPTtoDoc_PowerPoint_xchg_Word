using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPnt= Microsoft.Office.Interop.PowerPoint;
using WinWord=Microsoft.Office.Interop.Word;



namespace CharacterConverttoCharacterPics
{
    public partial class Form1 : Form
    {
        static object ppnt;
        static object wwrd;
        static PowerPnt.Application pptApp; 
        static  WinWord.Application wdApp;
        PowerPnt.Presentation ppt;
        WinWord.Document doc; 
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ppnt = getPowerPnt();
            wwrd = getWinWord();            
            if (ppnt==null||wwrd==null)
            {
                MessageBox.Show("請開啟 Word 與 PowerPoint 再繼續");                
            }
            else
            {
                pptApp = (PowerPnt.Application)ppnt;
                wdApp = (WinWord.Application)wwrd;                
                if (wdApp.Documents.Count > 0)
                {
                    doc = wdApp.ActiveDocument;
                    textBox1.Text = doc.FullName;
                }
                ppt = pptApp.ActivePresentation;
            }
            
        }

        static object getPowerPnt()
        {
            ppnt = System.Runtime.InteropServices.Marshal.GetActiveObject
                ("PowerPoint.Application");
            return ppnt;
        }
        static object getWinWord()
        {
            try
            {
                wwrd = System.Runtime.InteropServices.Marshal.GetActiveObject
                    ("Word.Application");
            }
            catch (Exception)
            {
                //MessageBox.Show("請開啟Word！");
                //throw;
            }
            return wwrd;
        }

        private void Form1_Click(object sender, EventArgs e)
        {
            this.go();
        }
        void go()
        {
            
            if (ppnt == null ||wwrd==null)
                return;
            if (!File.Exists(textBox1.Text))
            {
                MessageBox.Show("來源Word檔路徑有誤，請檢查！");
                return;
            }            
        }
    }
}
