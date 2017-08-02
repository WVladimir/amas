using System;
using System.IO;

namespace PatternDoc

{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        static string inputFile = "";
        static string outputFile = "";
        static string substitutionData = "";

        static void Main(string[] argv)
        {

            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
 
            

            ClassLOAccess DoWriter = new ClassLOAccess();

            int i = 0;

            if (argv.Length == 3)
            {
                for (i = 0; i < argv.Length; i++)
                {
                    switch (i)
                    {
                        case 0:
                            inputFile = argv[i];
                            break;
                        case 1:
                            outputFile = argv[i];
                            break;
                        case 2:
                            substitutionData = argv[i];
                            break;
                    }
                }

                DoWriter.ClassLOAcc(inputFile, outputFile, substitutionData);

                try
                {
                    File.Delete(outputFile);
                }
                catch (System.Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                ExtLiibreOffice is_doc = FlExtention(inputFile);

                if (is_doc == ExtLiibreOffice.calc)
                {
                    DoWriter.acalc();
                }
                else if (is_doc == ExtLiibreOffice.write)
                {
                    DoWriter.dwriter();
                }
            }
        }

        enum ExtLiibreOffice { write, calc, zero }

        static ExtLiibreOffice FlExtention(string filename)
        {
            filename = filename.Trim();
            if (
                filename.Substring(filename.Length - 4).ToLower().CompareTo(".odt") == 0 ||
                filename.Substring(filename.Length - 4).ToLower().CompareTo(".doc") == 0  ||
                filename.Substring(filename.Length - 5).ToLower().CompareTo(".docx") == 0
                ) return ExtLiibreOffice.write;
            else if (
                (filename.Substring(filename.Length - 4).ToLower().CompareTo(".ods") == 0) ||
                (filename.Substring(filename.Length - 4).ToLower().CompareTo(".xls")==0) ||
                (filename.Substring(filename.Length - 5).ToLower().CompareTo(".xlsx")==0)
                ) return ExtLiibreOffice.calc;
            else return ExtLiibreOffice.zero;
        }
    }
}




   
