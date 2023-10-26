using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace DLMSMinimal
{
    class Excel
    {
        string path = Path.GetFullPath("prog-feladat-jelsz.xlsx");
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel()
        { 
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
        }

        public string ReadCell(int i,  int j)
        {
            i++;
            j++;
            if (ws.Cells[i, j].Value2 != null) return ws.Cells[i, j].Value2;
            else return null;
        }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            Excel ex = new Excel();
            
            for (int i = 1; ex.ReadCell(i, 1)  != null; i++)
            {
                string text = i + ". " + ex.ReadCell(i, 0);
                string passMessage = ": Eros jelszo!";
                int ucCount = 0;
                int lcCount = 0;
                int nrCount = 0;
                int spCount = 0;

                if (ex.ReadCell(i, 1).Length >= 8)
                {
                    for (int j = 0; j < ex.ReadCell(i, 1).Length; j++)
                    {
                        if (char.IsLetter(ex.ReadCell(i, 1)[j]))
                        {
                            if (char.IsUpper(ex.ReadCell(i, 1)[j])) ucCount++;
                            else lcCount++;
                        }
                        else if (char.IsNumber(ex.ReadCell(i, 1)[j])) nrCount++;
                        else spCount++;
                    }

                    if (ucCount < 2 ||
                        lcCount < 2 ||
                        nrCount < 2 ||
                        spCount < 1)
                        passMessage = ": Gyenge jelszo!";
                }
                else passMessage = ": Gyenge jelszo!";

                text += passMessage;
                
                Console.WriteLine(text);
            }
            Console.ReadKey();
        }
    }
}
