using System.IO;
using System.Windows.Forms;
using winWord = Microsoft.Office.Interop.Word;

namespace Doc_PPt
{
    public class DocOps
    {
        static winWord.Application wdApp;
        public DocOps()
        {
            wdApp = docApp.getDocApp();
        }
        internal static winWord.Document openDoc(string docFullname)
        {
            //string dFullname = getDocFullname();
            if (docFullname == "" || !File.Exists(docFullname))
            {
                MessageBox.Show("請在textBox1文字方塊輸入「字源圖片」的「正確的」全檔名");
                return null;
            }
            return wdApp.Documents.Open(docFullname);
        }

        internal static string getDocFullname()
        {
            TextBox textBox1 = (TextBox)Application.OpenForms[0].Controls["textBox1"];
            if (textBox1.Text.IndexOf("字源圖片") > 1)
            {
                return textBox1.Text.Replace(@"file:///", "").Replace("%20", " ");
            }
            else
            {
                MessageBox.Show("請在textBox1文字方塊輸入「字源圖片」的全檔名"); return "";
            }
        }
    }
}
