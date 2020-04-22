using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace LockedPowerLibrary
{
    /// <summary>
    /// Класс для работы с файлами невыпускаемой мощности
    /// </summary>
    public class DataSearch
    {
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
        public static int MDPSearcher(string filePath, int hour, string section)
        {
            Application excelFile = WorkbookBaseData(filePath, "8 КС (2)");

            Range rangeOfSection = excelFile.Cells.Find(section, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart);
            if (rangeOfSection == null)
            {
                throw new ArgumentException("Нет такой ячейки!");
            }

            Range rangeOfMDP = excelFile.Cells[rangeOfSection.Row + 1 + hour,
                rangeOfSection.Column];

            if (!int.TryParse(rangeOfMDP.Value2.ToString(), out int valueMDP))
            {
                throw new ArgumentException("Значение в ячейке не явлется числом");
            }

            excelFile.Workbooks.Close();

            return valueMDP;
        }

        /// <summary>
        /// Поиск необходимого параметра в Excel файле
        /// </summary>
        /// <param name="filePath">Путь к Excel файлу для поиска</param>
        /// <param name="nameOfParam">Название параметра</param>
        /// <param name="nameOfSystem">Название энергосистемы</param>
        /// <returns>Значение рассматриваемого параметра</returns>
        public static double ParametrsSearcher(string filePath, string nameOfParam, string nameOfSystem)
        {
            Application excelFile = WorkbookBaseData(filePath, "Баланс мощности");

            Range rangeOfParam = excelFile.Cells.Find(nameOfParam, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart);
            Range rangeOfSystem = excelFile.Cells.Find(nameOfSystem, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart);
            if (rangeOfParam == null || rangeOfSystem == null)
            {
                throw new ArgumentException("Нет такой ячейки!");
            }

            Range rngOfParamValue = excelFile.Cells[rangeOfSystem.Row,
                rangeOfParam.Column];

            if (!double.TryParse(rngOfParamValue.Value2.ToString(), out double paramValue))
            {
                throw new ArgumentException("Значение в ячейке не явлется числом");
            }

            excelFile.Workbooks.Close();

            return Math.Round(paramValue, 2);
        }
    }
}
