using System;
using MSWord = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;

namespace WordDiff
{
    class WordDiff
    {
        [DllImport("user32.dll")]
        private static extern
            bool SetForegroundWindow(IntPtr hWnd);
        
        static void Main(string[] args)
        { 
            //
            //      Check for filenames and validate files exist
            //
            if (args.Length != 2)
            {
                System.Console.WriteLine("To compare two Microsoft Word documents use the following syntax:");
                System.Console.WriteLine("\tWordDiff [Document1] [Document2]");
                return;
            }

            object file1 = args[0];
            object file2 = args[1];

            if (! System.IO.File.Exists((string) file1))
            {
                System.Console.WriteLine("File does not exist: {0}", file1);
                return;
            }
            object fullPath1 = System.IO.Path.GetFullPath((string) file1);

            if (!System.IO.File.Exists((string)file2))
            {
                System.Console.WriteLine("File does not exist: {0}", file2);
                return;
            }
            object fullPath2 = System.IO.Path.GetFullPath((string)file2);

            //
            //      Open Word and compare files
            //
            MSWord.Application msWordApp = new MSWord.Application();
            msWordApp.Visible = false;
            object trueObj = (object) true;
            object falseObj = (object) false;
            object missingObj = Type.Missing;
            MSWord.Document msWordDocument1 = msWordApp.Documents.Open(ref fullPath1, ref missingObj, ref falseObj, ref falseObj, 
                ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, 
                ref trueObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

            MSWord.Document msWordDocument2 = msWordApp.Documents.Open(ref fullPath2, ref missingObj, ref falseObj, ref falseObj, 
                ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj, 
                ref missingObj, ref missingObj, ref missingObj, ref missingObj, ref missingObj);

            MSWord.Document msWordDocumentDiff = msWordApp.CompareDocuments(msWordDocument1, msWordDocument2, 
                MSWord.WdCompareDestination.wdCompareDestinationNew, MSWord.WdGranularity.wdGranularityWordLevel,
                true, true, true, true, true, true, true, true, true, true, "", true);

            msWordDocument1.Close(ref missingObj, ref missingObj, ref missingObj);
            msWordDocument2.Close(ref missingObj, ref missingObj, ref missingObj);

            //
            //      Make sure Word is active and in the foreground
            //
            msWordApp.Visible = true;
            msWordApp.Activate();
            msWordApp.ActiveWindow.SetFocus();
            SetForegroundWindow((System.IntPtr) msWordApp.ActiveWindow.Hwnd);

            return;
        }
    }
}
