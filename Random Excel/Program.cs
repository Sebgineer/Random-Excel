using IronXL;
using System;
using System.Diagnostics;
using System.Security.Cryptography;

namespace Random_Excel
{
    internal class Program
    {

        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            CryptoRnd(1000);
            CryptoSafeRnd(1000);
        }

        static void CryptoRnd(int length)
        {
            WorkBook book = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet sheet = book.CreateWorkSheet("Random");
            Random rnd = new Random();

            sheet["A1"].Value = "Random Numbers";
            sheet["B1"].Value = "Time Tick";
            Stopwatch sw = new Stopwatch();

            for (int i = 2; i < length + 2; i++)
            {
                sw.Restart();
                int rndNum = rnd.Next();
                sw.Stop();

                sheet[$"A{i}"].Value = rndNum;
                sheet[$"B{i}"].Value = sw.ElapsedTicks;
            }
            decimal sum = sheet[$"B2:B{length + 2}"].Sum();
            sheet[$"B{length + 2}"].Value = sum;

            book.SaveAs(@"C:\Users\seba6474\Documents\RandomExcel\Random_Test.xlsx");
        }

        static void CryptoSafeRnd(int length)
        {
            WorkBook book = WorkBook.Create(ExcelFileFormat.XLSX);
            WorkSheet sheet = book.CreateWorkSheet("Random");
            RNGCryptoServiceProvider rnd = new RNGCryptoServiceProvider();
            byte[] data = new byte[4];

            sheet["A1"].Value = "Random Numbers";
            sheet["B1"].Value = "Time Tick";
            Stopwatch sw = new Stopwatch();

            for (int i = 0; i < length; i++)
            {
                sw.Restart();
                rnd.GetBytes(data);
                int rndNum = BitConverter.ToInt32(data, 0);
                sw.Stop();

                sheet[$"A{i + 2}"].Value = rndNum;
                sheet[$"B{i + 2}"].Value = sw.ElapsedTicks;
            }
            decimal sum = sheet[$"B2:B{length + 2}"].Sum();
            sheet[$"B{length + 2}"].Value = sum;

            book.SaveAs(@"C:\Users\seba6474\Documents\RandomExcel\RandomCrypto_Test.xlsx");
        }
    }
}
