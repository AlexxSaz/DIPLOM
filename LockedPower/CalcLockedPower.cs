using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LockedPowerLibrary;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace LockedPower
{
    class CalcLockedPower
    {
        /// <summary>
        /// Получить массив имен из файла
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        /// <returns>Массив имен</returns>
        private static string[] TextReader(string path)
        {
            path = @"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\" + path;

            var streamReader = new StreamReader(path);

            var str = new string[File.ReadAllLines(path).Length];

            for (int i = 0; i < str.Length; i++)
            {
                str[i] = Convert.ToString(streamReader.ReadLine());
            }
            streamReader.Close();

            return str;
        }

        static void Main(string[] args)
        {
            var section = TextReader("SectionsName.txt");
            var energySystem = TextReader("EnergySystems.txt");
            var parametr = TextReader("NameOfParameters.txt");

            Application reportExcel = new Application();
            Workbook reportWb = reportExcel.Workbooks.Add();
            Worksheet reportWs = reportWb.Worksheets[1];

            reportWs.Name = "Расчет невыпускаемой мощности";

            

            int counter = 0;
            try
            {
                for (int i = 0; i < energySystem.Length; i++)
                {
                    reportWs.Cells[1 + counter+i, 1] = energySystem[i];

                    for (int j = 0; j < parametr.Length; j++)
                    {
                        reportWs.Cells[1 + counter + j, 2] = parametr[j];
                        reportWs.Cells[1 + counter + j, 3] = DataSearch.ParametrsSearcher(@"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\balance10-29-01-2020.xlsx", parametr[j], energySystem[i]);
                        counter++;
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
            }
            reportWb.SaveAs(@"C:\Users\Александр\Desktop\МусорницаОтчетов\2.xlsx");

            reportWb.Close();

            Console.ReadKey();
        }
    }
}
