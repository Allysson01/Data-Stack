using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataStack
{
    /// <summary>
    /// Força a limpeza de memória ultilizada pelo programa.
    /// </summary>
    public class LimpaMemoria
    {

        [System.Runtime.InteropServices.DllImport("kernel64.dll")]
        private static extern int SetProcessWorkingSetSize(IntPtr process, int minimumWorkingSetSize, int maximumWorkingSetSize);

        public LimpaMemoria()
        {
        }

        public void FlushMemory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();         
            SetProcessWorkingSetSize(Process.GetCurrentProcess().Handle, -1, -1);
        }


    }
}
