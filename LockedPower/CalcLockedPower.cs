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
    /// Расчет значения невыпускаемой мощности
    /// </summary>
    class CalcLockedPower
    {
        /// <summary>
        /// 
        /// </summary>
        private static ConsoleHelper.SignalHandler signalHandler { get; set; }

        /// <summary>
        /// Килл процесса EXCEL
        /// </summary>
        /// <param name="consoleSignal">Сигнал, поступающий от консоли</param>
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

            int rowCounter = 3;
            int systemCounter = 0;
            try
            {
                var valueMDP = DataSearch.MDPSearcher(@"E:\Programms\С# Progs\DIPLOM\LockedPower\Resources\ppbr(a)_04032020_1.xls", 6);
                var arrayNameMDP = DataSearch.TextReader("sectionsname.txt");
                var valueParametr = DataSearch.ParametrsSearcher("balance10-04-03-2020.xlsx");

                List<string> nameMDP = new List<string>(arrayNameMDP);

                for (int i = 0; i < valueParametr.GetLength(0); i++)
                {
                    ++systemCounter;
                    for (int j = 0; j < valueParametr.GetLength(1); j++)
                    {
                        reportWs.Cells[rowCounter, 3] = valueParametr[i, j];
                        rowCounter++;
                    }
                    reportWs.Cells[rowCounter, 3] =
                        DataCalc.ReserveCalc(i, valueParametr);
                    rowCounter++;

                    foreach (string s in nameMDP)
                    {
                        if (reportWs.Cells[rowCounter, 2].Value2 == s)
                        {
                            reportWs.Cells[rowCounter, 3] =
                                valueMDP[nameMDP.IndexOf(s)];
                            rowCounter++;
                            reportWs.Cells[rowCounter, 3] =
                                DataCalc.LockedPowerCalc(i, valueParametr,
                                valueMDP[nameMDP.IndexOf(s)],
                                nameMDP.IndexOf(s) != 0 ? valueMDP[nameMDP.IndexOf(s) - 1] : 0,
                                systemCounter);
                            rowCounter++;
                            systemCounter = 0;
                        }
                    }
                }
                reportWb.SaveAs(@"C:\Users\Александр\Desktop\МусорницаОтчетов\2.xlsx");
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("Расчет не выполнен!");
            }
            finally
            {
                reportWb.Close(false);
                reportExcel.Quit();
                reportExcel = null;
                reportWb = null;
                reportWs = null;
                GC.Collect();
            }

            Console.ReadKey();
        }
    }
}
