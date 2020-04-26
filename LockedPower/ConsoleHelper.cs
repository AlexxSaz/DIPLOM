using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace LockedPower
{
    /// <summary>
    /// Вспомогательны класс для консоли
    /// </summary>
    internal class ConsoleHelper
    {
        internal delegate void SignalHandler(ConsoleSignal consoleSignal);

        /// <summary>
        /// Обработка сигнала
        /// </summary>
        /// <param name="handler">Метод обработки сигнала</param>
        /// <param name="add"></param>
        /// <returns></returns>
        [DllImport("Kernel32", EntryPoint = "SetConsoleCtrlHandler")]
        public static extern bool SetSignalHandler(SignalHandler handler, bool add);
    }
}
