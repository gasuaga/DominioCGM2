using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
namespace Dominio
{
    /**
     * @class   Excel
     *
     * @brief   Se encarga de todas las acciones sobre un archivo Excel
     *
     * @author  WINMACROS
     * @date    14/07/2017
     */

    class Excel
    {
        #region propertys
        public URL Direccion { get; set; }
        #endregion

        #region contructores
        public Excel(HttpPostedFile pDir)
        {
            Direccion = new URL(pDir);
        }
        public Excel(string pNombre, string pDireccion)
        {
            Direccion = new URL(pNombre, pDireccion);
        }
        public Excel() { }
        #endregion

        /**
         * @fn  public string campaniaAsignada()
         *
         * @brief   Busca dentro del excel a que marcador 
         *          se tiene que cargar.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @return  Nombre del marcador donde se carga
         *          String.
         */
        public string campaniaAsignada()
        {
            var column2 = new List<string>();
            using (var rd = new StreamReader(Direccion.ToString()))
            {
                while (!rd.EndOfStream)
                {
                    var splits = rd.ReadLine().Split(';');
                    column2.Add(splits[1]);
                }
                rd.Close();
            }
            return column2.ElementAt(1);
        }

        /**
         * @fn  public List<string[]> leerExcel()
         *
         * @brief   Entra al archivo que pasa el usuario y lee el excel
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @return  Retorna una lista de strimg[] lo cual cada elementeo de la lista
         *          es una linea.
         */
        public List<string[]> leerExcel()
        {
            List<string[]> excel = new List<string[]>();
            String dir = Direccion.ToString();
            string[] aux;
            try
            {
                foreach (var item in File.ReadLines(dir))
                {
                    aux = item.Split(';');
                    if (aux[0] != "")
                        excel.Add(aux);
                }
            }
            catch (Exception){}
            return excel;
        }

        /**
         * @fn  public void hacerScore()
         *
         * @brief   Lee linea por linea y si donde va la id no dice id 
         *          llama a un metodo auxiliar el cual calcula el scor
         *          y lo inserta en la abse.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */

        public void hacerScore()
        {
            String dir = Direccion.ToString();
            string[] aux;
            tolls t = tolls.T;
            foreach (var aux2 in File.ReadLines(dir))
            {
                aux = aux2.Split(';');
                if (aux.ElementAt(0) != "id")
                    t.poblarScore(aux);
            }
        }

        /**
         * @fn  public void cargarContacto(string pContacto)
         *
         * @brief   Incerta una liea en el excel.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pContacto   Linea separada por ; cada campo para insertar en 
         *                      el excel.
         */

        public void cargarContacto(string pContacto)
        {
            tolls t = tolls.T;
            t.matarProceso("EXCEL");
            StreamWriter sR = new StreamWriter(Direccion.ToString(), true);
            if (pContacto.Split(';').Count() == 30) // si tiene mas o menos que 30 seldas esta mal la linea
                sR.WriteLine(pContacto);
            sR.Close();
        }

        /**
         * @fn  public static void crearExcel(List<string[]> pExcel)
         *
         * @brief   Convierte una lista de arrays en un excel (Utilizado  para los sms)
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pExcel  Excel a contruir.
         */

        public static void crearExcel(List<string[]> pExcel)
        {//CRM\" & Day(Date) & Month(Date) & "arch.xls
            StreamWriter sw = new StreamWriter(@"E:\Ivrs\CRM\" + DateTime.Today.ToString("ddMM") + "arch.csv", false);
            foreach (string[] s in pExcel)
            {
                if (validaLinea(s))
                {
                    foreach (string e in s)
                    {
                        sw.Write(e + ";");
                    }
                    sw.WriteLine();
                }
            }
            sw.Close();
            sw.Dispose();
        }

        /**
         * @fn  private static bool validaLinea(string[] pLinea)
         *
         * @brief   Valida la linea del excel antes insertarla.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pLinea  Linea a insertar.
         *
         * @return  True si la linea es correcta, false si no lo es.
         */

        private static bool validaLinea(string[] pLinea)
        {
            foreach (string s in pLinea)
            {
                if (s.Split(';').Count() > 1) //Si la linea contiene un ; esta deficiente 
                    return false;
            }
            return true;
        }

        /**
         * @fn  public bool validarExcel()
         *
         * @brief   Determina si un excel es valido para cargarse a incosert o no.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @return  True si es valido, false si no lo es.
         */

        public bool validarExcel()
        {
            List<string[]> excel = leerExcel();
            string[] cabezal = excel.ElementAt(0);
            int cont = 0;
            if (excel.ElementAt(1)[0] == "") return false;
            while (cabezal.Length > cont)
            {
                string actual = cabezal.ElementAt(cont);
                if (actual == "")
                    return false;
                cont++;
            }
            return true;
        }

    }
}
