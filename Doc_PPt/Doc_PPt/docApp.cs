using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using winWord = Microsoft.Office.Interop.Word;

namespace Doc_PPt
{
    public class docApp
    {
        internal static winWord.Application getDocApp() {
            try
            {
                return (winWord.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch (Exception)
            {
                return new winWord.Application();
                //throw;
            }
            
        }
    }
}
