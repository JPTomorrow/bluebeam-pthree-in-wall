using System;

namespace JPMorrow.ConsoleDebugger {
    public static class Debug {
        public static void Show(string txt, bool error = false) 
        {
            Console.WriteLine((error ? "ERROR: " : "") + txt);
        }
    }
}