using System;
using System.Windows.Forms;
using WatermarkTool.Forms;

namespace WatermarkTool
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序主入口点
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetHighDpiMode(HighDpiMode.SystemAware);

            Application.Run(new MainForm());
        }
    }
}
