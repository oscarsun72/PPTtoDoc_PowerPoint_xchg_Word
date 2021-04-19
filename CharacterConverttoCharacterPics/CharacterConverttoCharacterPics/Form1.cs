﻿using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using WinWord = Microsoft.Office.Interop.Word;
using powerPnt = Microsoft.Office.Interop.PowerPoint;


namespace CharacterConverttoCharacterPics
{
    public partial class Form1 : Form
    {
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

        private void button1_Click(object sender, EventArgs e)
        {
            string fontname = textBox2.Text;
            WinWord.Document wd = fontsPics.getFontCharacterset
                (fontname);
            if (wd != null)
            {
                BackColor = Color.Red;
                this.Enabled = false;button1.Enabled = false;
                string picFolder = textBox1.Text;
                if (picFolder.IndexOf(fontname) == -1)
                { picFolder += ("\\" + fontname + "\\"); }
                powerPnt.Presentation ppt = 
                    fontsPics.prepareFontPPT(fontname, float.Parse(textBox3.Text));
                fontsPics.addCharsSlidesExportPng(wd,ppt,picFolder);
                BackColor = Color.Green;
                this.Enabled = true; button1.Enabled = true;
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
                MessageBox.Show("這個字型已經做過了！或是不必做的","請檢查！",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                button1.Enabled = false;
                return;
            }
            button1.Enabled = true;
        }

        private void textBox1_Click(object sender, EventArgs e)
        {
            textBox1.Text = Clipboard.GetText();

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
string dirPath = textBox1.Text;
            if (Directory.Exists( dirPath))
            {//開啟資料夾：
                Process prc = new Process();
                prc.StartInfo.FileName = dirPath;
                prc.Start();
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    this.Close();
                    break;
                default:
                    break;
            }
        }
    }
}
