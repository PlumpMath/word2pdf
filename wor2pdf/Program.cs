using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace word2pdf
{
    class Program
    {
        private static _Application word;
        private static void Main(string[] args)
        {
            try
            {
                word = new Application();
                word.Visible = false;
            }
            catch (Exception)
            {
                Console.WriteLine("Error while opening ms word");
                Console.Read();
                return;
            }
            
            object oMissing = System.Reflection.Missing.Value;
            object isVisible = true;
            object readOnly = false;
            object oInput = "";
            object oOutput = "";
            object oFormat = WdSaveFormat.wdFormatPDF;

            foreach (var file in Directory.GetFiles(Environment.CurrentDirectory, "*.doc?"))
            {
                oInput = file;
                var index = file.LastIndexOf(@".doc", StringComparison.InvariantCultureIgnoreCase);
                oOutput = file.Substring(0, index);

                var index2 = file.LastIndexOf(@"\", StringComparison.InvariantCulture);
                var fName = file.Substring(index2, file.Length - index2);
                Console.WriteLine("Processing " + fName + " file...");

                Document doc;
                try
                {
                    doc = word.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    doc.Activate();
                }
                catch
                {
                    Console.WriteLine("Error while opening the file");
                    continue;
                }

                try
                {
                    doc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing);

                    doc.Close();
                }
                catch (Exception)
                {
                    Console.WriteLine("Error while saving the file");
                    continue;
                }

                Console.WriteLine("Done");
            }

            word.Quit();
            Console.WriteLine("Press any key to quit");
            Console.Read();
        }
        
    }
}
