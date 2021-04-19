using System;
using System.Runtime.InteropServices;
using winWord = Microsoft.Office.Interop.Word;

namespace Doc_PPt
{
    public class docApp
    {
        internal static winWord.Application getDocApp()
        {
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
