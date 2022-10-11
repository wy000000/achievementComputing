using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace achievementComputing
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Application.Run(new Form1());
            Console.WriteLine("press any key to exit...");
            //Console.ReadKey();
        }
    }
}
