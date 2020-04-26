using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LockedPowerLibrary;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace LockedPower
{
    /// <summary>
    /// 
    /// </summary>
    class CalcLockedPower
    {
        private static ConsoleHelper.SignalHandler signalHandler { get; set; }

        private static void HandleConsoleSignal(ConsoleSignal consoleSignal)
        {
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            signalHandler += HandleConsoleSignal;
            ConsoleHelper.SetSignalHandler(signalHandler, true);

            Application reportExcel = new Application();
            Workbook reportWb = reportExcel.Workbooks.Open(@"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\Shablon.xlsx");
            Worksheet reportWs = reportWb.Worksheets[1];

            reportWs.Name = "Расчет невыпускаемой мощности";
            reportWs.Cells[1, 1] = "Дата создания отчета: " + DateTime.Now.ToString();

            int counter = 0;
            try
            {
                var valueMDP = DataSearch.MDPSearcher(@"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\ppbr(a)_22012020_1.xls", 2);
                var valueParametr = DataSearch.ParametrsSearcher(@"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\balance10-29-01-2020.xlsx");

                //for (int i = 0; i < valueMDP.Length; i++)
                //{
                //    reportWs.Cells[i + 1, 1] = valueMDP[i];
                //}
                for (int i = 0; i < valueParametr.GetLength(0); i++)
                {
                    for (int j = 0; j < valueParametr.GetLength(1); j++)
                    {
                        reportWs.Cells[counter + 3, 3] = valueParametr[i, j];
                        counter++;
                    }
                }
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
            }

            reportWb.SaveAs(@"C:\Users\Александр\Desktop\МусорницаОтчетов\2.xlsx");
            reportWb.Close(false);
            reportExcel.Quit();
            reportExcel = null;
            reportWb = null;
            reportWs = null;
            GC.Collect();

            Console.ReadKey();
        }
    }
}
