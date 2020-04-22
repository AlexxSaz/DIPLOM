using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LockedPowerLibrary;

namespace LockedPower
{
    class CalcLockedPower
    {
        static void Main(string[] args)
        {
            var section = "Кузбасс-Запад";
            string energySystem = "Республики Бурятии";

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
                Console.WriteLine(e);
            }

            Console.WriteLine("Значение МДП для сечения " + section + " равно: " + MDP.ToString());
            Console.WriteLine("Установленная мощность энергосистемы " + energySystem + " равно: " + paramValue.ToString());
            Console.ReadKey();
        }
    }
}
