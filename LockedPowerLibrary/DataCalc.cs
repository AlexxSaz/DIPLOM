using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LockedPowerLibrary
{
    /// <summary>
    /// Расчет параметров ЭС
    /// </summary>
    public class DataCalc
    {
        /// <summary>
        /// Расчет резерва мощности ЭС
        /// </summary>
        /// <param name="numberOfEnergysystem">Номер рассматриваемой ЭС</param>
        /// <param name="parametersOfEsystem">Все параметр всех ЭС</param>
        /// <returns>Значение резерва мощности ЭС</returns>
        public static double ReserveCalc(int numberOfEnergysystem,
            double[,] parametersOfEsystem)
        {
            double valueWorkPower = DataSearch.SearchParamValue(numberOfEnergysystem,
                parametersOfEsystem, "Раб. мощн.");
            double valueLoad = DataSearch.SearchParamValue(numberOfEnergysystem,
                parametersOfEsystem, "Час совм. максимума");

            return valueWorkPower - valueLoad;
        }

        /// <summary>
        /// Расчет невыпускаемого резерва мощности
        /// </summary>
        /// <param name="actualES">Номер рассматриваемой ЭС</param>
        /// <param name="parametersOfEsystem"></param>
        /// <param name="valueMDP">МДП рассматриваемого сечения</param>
        /// <param name="powerFlow">Внешний переток ЭС</param>
        /// <param name="numberOfSystems">Количество ЭС перед сечением</param>
        /// <returns>Величина невыпускаемой мощности</returns>
        public static double LockedPowerCalc(int actualES, double[,] parametersOfEsystem, double valueMDP, double powerFlow,
            int numberOfSystems)
        {
            List<double> valueReserve = new List<double>(numberOfSystems);
            double valueLP = 0;

            for (int i = 0; i < valueReserve.Capacity; i++)
            {
                valueReserve.Add(ReserveCalc(actualES - i, parametersOfEsystem));
                valueLP += valueReserve[i];
            }

            return valueLP + powerFlow - valueMDP;
        }
    }
}
