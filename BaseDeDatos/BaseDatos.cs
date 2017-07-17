using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseDeDatos
{
    class BaseDatos
    {
        private static BaseDatos bd;

        public static BaseDatos Bd
        {
            get
            {
                if (bd == null)
                    bd = new BaseDatos();
                return bd;
            }
        }
        private BaseDatos() { }

       
    }
}
