using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace BaseDeDatos
{
    public static class CadenaDeConecciones
    {
        public enum tiposDeConeccion { paraDominio, poblarScore, crearScore }

        public static string CadenaConeccionDominio {
            get {
                return File.ReadAllLines(@"E:\CadenaDeConecciones.txt").ToList().ElementAt(0);
            }
        }
        public static string CadenaConeccionScore
        {
            get
            {
                return File.ReadAllLines(@"E:\CadenaDeConecciones.txt").ToList().ElementAt(1);
            }
        }
        public static string CadenaConeccionDatosScore
        {
            get
            {
                return File.ReadAllLines(@"E:\CadenaDeConecciones.txt").ToList().ElementAt(2);
            }
        }
        public static string CadenaConeccionDatosPosgre
        {
            get
            {
                return File.ReadAllLines(@"E:\CadenaDeConecciones.txt").ToList().ElementAt(3);
            }
        }
    }
}
