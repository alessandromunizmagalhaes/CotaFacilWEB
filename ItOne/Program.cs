using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ITOne
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            //Instanciando o Addon
            Addon addOn = new Addon();

            //Roda e mantém a aplicação ativa
            System.Windows.Forms.Application.Run();
        }
    }
}
