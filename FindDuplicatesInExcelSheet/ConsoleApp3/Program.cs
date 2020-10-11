using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp3
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Path should be of format: D:/Testsheet1.xlsx");
            Console.WriteLine("Enter file path");
            var path = Console.ReadLine();
            Console.WriteLine("Enter sheet number:{ex:2} for the second page");
            var index = Console.ReadLine();
            int.TryParse(index, out int value);
            FindDuplicates.DuplicateValues(path, value);
            Console.ReadLine();
        }
    }
}
