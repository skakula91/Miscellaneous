using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.Excel;


namespace ConsoleApp3
{
    public class FindDuplicates
    {
        public static void DuplicateValues(string path,int index)
        {
            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Open(path);
            Worksheet xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(index);
            int totalColumns = xlWorkSheet.UsedRange.Columns.Count;
            int totalRows = xlWorkSheet.UsedRange.Rows.Count;

            int? colNumber = null;

            SortedDictionary<string, List<string>> dict = new SortedDictionary<string, List<string>>();
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>();

            for (int col = colNumber?? 1; col <= (colNumber ?? totalColumns); col++)
            {
                for (int row = 1; row <= totalRows; row++)
                {
                    Range dataRange = (Range)xlWorkSheet.Cells[row, col];
                    string val = dataRange.Value2.ToString();
                    string address = dataRange.Address;
                    if (dict.ContainsKey(val))
                    {
                        if (result.ContainsKey(val))
                        {
                            result[val].Add(address);
                        }
                        else
                        {
                            result.Add(val,dict[val]);
                            result[val].Add(address);
                        }
                    }
                    else
                    {
                        dict.Add(val, new List<string>(){address});
                    }
                }
            }

            if (result.Count > 0)
            {
                var entries = result.Select(d =>
                    string.Format("\"{0}\": [{1}]", d.Key, string.Join(",", d.Value)));
                Console.WriteLine("{" + string.Join(",", entries) + "}");
            }
            Console.WriteLine();

            xlWorkBook.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkBook);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }

        private static int NumberFromExcelColumn(string column)
        {
            int retVal = 0;
            string col = column.ToUpper();
            for (int iChar = col.Length - 1; iChar >= 0; iChar--)
            {
                char colPiece = col[iChar];
                int colNum = colPiece - 64;
                retVal = retVal + colNum * (int)Math.Pow(26, col.Length - (iChar + 1));
            }
            return retVal;
        }
    }
}
