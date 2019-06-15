using NAPS2;
using NAPS2.Automation;
using NAPS2.DI.Modules;
using NAPS2.Worker;
using Ninject;
using Ninject.Parameters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace TwainScanner
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // Initialize Ninject (the DI framework)
            var kernel = new StandardKernel(new CommonModule(), new ConsoleModule());

            Paths.ClearTemp();

            // Parse the command-line arguments (and display help text if appropriate)
            var options = new AutomatedScanningOptions();

            //if (!CommandLine.Parser.Default.ParseArguments(args, options))
            //{
            //    return;
            //}

            // Start a pending worker process
            WorkerManager.Init();

            // Run the scan automation logic
            var scanning = kernel.Get<AutomatedScanning>(new ConstructorArgument("options", options));
            scanning.Execute().Wait();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
