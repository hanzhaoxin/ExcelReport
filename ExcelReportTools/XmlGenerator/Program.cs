using System;
using System.IO;
using System.Windows.Forms;
using Sunisoft.IrisSkin;

namespace XmlGenerator
{
    internal static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        private static void Main()
        {
            var skinEngine = new SkinEngine
            {
                SkinFile = Application.StartupPath + Path.DirectorySeparatorChar + "Skin" + Path.DirectorySeparatorChar +  "skin.ssk"
            };
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}