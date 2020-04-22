using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LockedPowerLibrary;
using System.IO;

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
            path = "Resources\\" + path;

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

            Console.WriteLine("Для поиска сечения " + section + " нажмите любую кнопку.");
            Console.ReadKey();

            int MDP = 0;
            double paramValue = 0;
            try
            {
                MDP = Locked.MDPSearcher(@"C:\Users\Александр\Desktop\ДИПЛОМ\Исходные данные\ppbr(a)_22012020_1.xls", 1, section);
                paramValue = Locked.ParametrsSearcher(@"C:\Users\Александр\Desktop\ДИПЛОМ\Исходные данные\balance10-29-01-2020.xlsx", "Устан. мощн.", energySystem);
            }
            catch (ArgumentException e)
            {
                Console.WriteLine(e.Message);
            }

            Console.WriteLine("Значение МДП для сечения " + section + " равно: " + MDP.ToString());
            Console.WriteLine("Установленная мощность энергосистемы " + energySystem + " равно: " + paramValue.ToString());
            Console.ReadKey();
        }
    }
}
