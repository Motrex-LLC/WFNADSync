using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WFNADSync
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());

            ProcessJob processJob = new ProcessJob();
            //processJob.addNewEmpADAccount();
            processJob.syncWFNExistingEmpToAD_ByEmployeeTable();
            //processJob.syncWFNExistingEmpToAD_ByChangeTable();
            processJob.disableTerminatedEmpADAccount();

            //4545
        }
    }
}
