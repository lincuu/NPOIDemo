using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace NPOIDemo
{
    static class Program
    {
        /// <summary>
        /// 应用程序的主入口点。
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            NPOIDemo2();
            return;

            Application.Run(new FormMain());
        }

        private static void NPOIDemo2()
        {
            //NPOIDemo.CreateExcelColorFile("ExcelColor.xlsx");
            System.Data.DataTable dataSource = NPOIExtension.GetDataTabelFromExcelFile("Test.xlsx", true);




        }


    }
}
