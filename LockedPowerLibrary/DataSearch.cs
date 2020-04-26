using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace LockedPowerLibrary
{
    /// <summary>
    /// Класс для работы с файлами невыпускаемой мощности
    /// </summary>
    public class DataSearch
    {
        /// <summary>
        /// Получить массив имен из файла
        /// </summary>
        /// <param name="path">Путь к файлу</param>
        /// <returns>Массив имен</returns>
        private static string[] TextReader(string path)
        {
            path = @"E:\Programms\С# Progs\DIPLOM\LockedPowerLibrary\Resources\" + path;

            var streamReader = new StreamReader(path);

            var str = new string[File.ReadAllLines(path).Length];

            for (int i = 0; i < str.Length; i++)
            {
                str[i] = Convert.ToString(streamReader.ReadLine());
            }
            streamReader.Close();

            return str;
        }

        /// <summary>
        /// Задание основных параметров файла Excel
        /// </summary>
        /// <param name="path">Путь к рабочему файлу</param>
        /// <param name="nameOfList">Название рабочего листа</param>
        /// <returns>Файл Excel</returns>
        private static Application WorkbookBaseData(string path, string nameOfList)
        {
            Application excelFile = new Application();
            Workbook workbook = excelFile.Workbooks.Open(path);
            Worksheet worksheet = workbook.Worksheets[nameOfList];

            return excelFile;
        }

        /// <summary>
        /// Нахождение необходимого значения МДП
        /// </summary>
        /// <param name="filePath">Путь к Excel файлу для поиска</param>
        /// <param name="hour">Расчетный час</param>
        /// <param name="section">Рассматриваемое сечение</param>
        /// <returns>Искомое значение МДП</returns>
        public static int[] MDPSearcher(string filePath, int hour)
        {
            Application excelFile = WorkbookBaseData(filePath, "8 КС (2)");
            var section = TextReader("SectionsName.txt");
            var valueMDP = new int[section.Length];

            for (int i = 0; i < section.Length; i++)
            {
                Range rangeOfSection = excelFile.Cells.Find(section[i], Type.Missing,
                    XlFindLookIn.xlValues, XlLookAt.xlWhole);
                if (rangeOfSection == null)
                {
                    throw new ArgumentException("Нет такой ячейки!");
                }

                Range rangeOfMDP = excelFile.Cells[rangeOfSection.Row + 1 + hour,
                    rangeOfSection.Column];

                if (int.TryParse(rangeOfMDP.Value2.ToString(), out int value))
                {
                    valueMDP[i] = value;
                }
                else
                {
                    throw new ArgumentException("Для сечения " + section[i] +
                        " значение в ячейке не является числом");
                }
            }

            excelFile.Workbooks.Close();
            excelFile.Quit();
            excelFile = null;
            GC.Collect();

            return valueMDP;
        }

        /// <summary>
        /// Поиск необходимого параметра в Excel файле
        /// </summary>
        /// <param name="filePath">Путь к Excel файлу для поиска</param>
        /// <param name="nameOfParam">Название параметра</param>
        /// <param name="nameOfSystem">Название энергосистемы</param>
        /// <returns>Значение рассматриваемого параметра</returns>
        public static double[,] ParametrsSearcher(string filePath)
        {
            Application excelFile = WorkbookBaseData(filePath, "Баланс мощности");
            var energySystem = TextReader("EnergySystems.txt");
            var parametr = TextReader("NameOfParameters.txt");
            var paramValue = new double[energySystem.Length, parametr.Length];

            for (int i = 0; i < energySystem.Length; i++)
            {
                Range rangeOfSystem = excelFile.Cells.Find(energySystem[i], Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);

                for (int j = 0; j < parametr.Length; j++)
                {
                    Range rangeOfParam = excelFile.Cells.Find(parametr[j], Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);
                    if (rangeOfParam == null || rangeOfSystem == null)
                    {
                        throw new ArgumentException("Нет такой ячейки!");
                    }

                    Range rngOfParamValue = excelFile.Cells[rangeOfSystem.Row,
                        rangeOfParam.Column];

                    if (double.TryParse(rngOfParamValue.Value2.ToString(), out double value))
                    {
                        paramValue[i, j] = Math.Round(value, 2);
                    }
                    else
                    {
                        throw new ArgumentException("Для ЭС " + energySystem[i] +
                            " значение в ячейке параметра " + parametr[j] +
                            " не является числом");
                    }
                }
            }

            excelFile.Workbooks.Close();
            excelFile.Quit();
            excelFile = null;
            GC.Collect();

            return paramValue;
        }

        /// <summary>
        /// Поиск параметра в шаблоне отчета
        /// </summary>
        /// <param name="worksheet">Лист отчета</param>
        /// <param name="counter">Строка, после которой ведется поиск</param>
        /// <param name="nameOfParam">Название искомого параметра</param>
        /// <returns>Значение параметра</returns>
        private static double SearchValueInShablon(Worksheet worksheet,
            int counter, string nameOfParam)
        {
            Range paramRange = worksheet.Cells.Find(nameOfParam,
                worksheet.Cells[counter + 1, 3],
                XlFindLookIn.xlValues, XlLookAt.xlWhole);
            if (paramRange == null)
            {
                throw new ArgumentException("Нет такой ячейки для расчета резерва!");
            }

            double.TryParse(worksheet.Cells[paramRange.Row,
                paramRange.Column + 1].Value2.ToString(),
                out double value);

            return value;
        }

        /// <summary>
        /// Расчет резерва мощности
        /// </summary>
        /// <param name="worksheet">Лист для расчета</param>
        /// <param name="counter">Строка, после которой ведется поиск</param>
        /// <returns>Значение резерва мощности в ЭС</returns>
        public static double ReserveCalc(Worksheet worksheet, int counter)
        {
            double valueWorkPower = SearchValueInShablon(worksheet, counter, "Рабочая мощность");
            double valueLoad = SearchValueInShablon(worksheet, counter, "Потребление");

            return valueWorkPower - valueLoad;
        }
    }
}
