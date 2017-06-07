using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace MailerUtilities
{
    public class messaggio {
        public List<string> Destinatario { get; set; }
        public List<string> Copia { get; set; }
        public List<string> CopiaNascosta { get; set; }
        public List<string> Allegati { get; set; }
        public string Oggetto { get; set; }
        public string Messaggio { get; set; }

    }
    public class XLSX
    {
        public List<messaggio> elaboraExcel(string percorso)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(@percorso, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Range range = xlWorkSheet.UsedRange;
            int righe = range.Rows.Count;
            int colonne = range.Columns.Count;
            int rCnt;
            int cCnt;
            string str;
            for (rCnt = 1; rCnt <= righe; rCnt++)
            {
                for (cCnt = 1; cCnt <= colonne; cCnt++)
                {
                    str = (string)(range.Cells[rCnt, cCnt] as Range).Value2;
                }
            }
            return new List<messaggio>();
        }

    }

}
