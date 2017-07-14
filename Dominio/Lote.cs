using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
{
    /**
     * @class   Lote
     *
     * @brief   Es la clase que se encarga de los
     *          lotes que se van a subir a inconsert 
     *
     * @author  WINMACROS
     * @date    14/07/2017
     */

    class Lote : IEquatable<Lote>, IComparable<Lote>
    {
        #region Variables
        /**
         * @enum    tipoEstado
         *
         * @brief   Son los distintos estados que puede tener un lote
         */
        public enum tipoEstado { Activo, ParaCargar, Eliminado, paraEliminar, Importado, ErrorEliminacion };
        /**
         * @enum    tipoLote
         *
         * @brief   Son los distintos tipos de lote que puede tener  
         *          (errorCampana es cuadno se ingresa mal el motor dentro del excel).
         */
        public enum tipoLote { VB, IVR, Marcador, errorCampana };
        /**
         * @enum    tipoRecurso
         *
         * @brief   Tipo de area a la cual perteneze el lote
         */
        public enum tipoRecurso { marcadorTardia, marcadorTemprana };
        #endregion

        #region propertys
        public string Nombre { get; set; }
        public DateTime Creacion { get; set; }//Fecha de creacion
        public tipoEstado Estado { get; set; }
        public DateTime Eliminado { get; set; } // fecha de eliminacion
        public tipoRecurso UnidadNegocio { get; set; }
        public tipoLote LoteTipo { get; set; }
        public Frecuencia Frec { get; set; }//Frecuencia del lote
        public Marcador Marc { get; set; }//Marcador del lote
        public Excel Exc { get; set; }//Excel ligado al lote
        public bool EliminarLotesAnt { get; set; } //Elimina los lotes de el motor donde se carga 
        #endregion

        #region contructores
        public Lote(string pNombre, DateTime pCreacion, tipoEstado pEstado, tipoLote pTipo, int pFrecuencia, Excel pExc, bool pEliminar)
        {
            Nombre = pNombre;
            Creacion = pCreacion;
            Estado = pEstado;
            LoteTipo = pTipo;
            Exc = pExc;
            EliminarLotesAnt = pEliminar;
            Frec = Frecuencia.frecuenciaXid(pFrecuencia);
        }
        public Lote(string pNombre, DateTime pCreacion, tipoEstado pEstado, tipoLote pTipo, bool pEliminar)
        {
            Nombre = pNombre;
            Creacion = pCreacion;
            Estado = pEstado;
            LoteTipo = pTipo;
            EliminarLotesAnt = pEliminar;
            Frec = Frecuencia.frecuenciaXid(10);
        }
        public Lote(string pNombre)
        {
            Nombre = pNombre;
            Frec = Frecuencia.frecuenciaXid(10);
        }
        #endregion

        /**
         * @fn  public bool asignarMarcador()
         *
         * @brief   Busca dentro de todos los marcadores del sistema
         *          si se encuentra el que tiene escrito en el excel
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @return  True si lo encuentra, False si no existe el marcador
         *          en el sistema
         */
        public bool asignarMarcador()
        {
            Sistema s = Sistema.Sis;
            string campaniaAsignada = Exc.campaniaAsignada();
            Marcador m = new Marcador(campaniaAsignada);
            bool bandera = false;
            int cont = 0;
            while (!bandera && cont < s.Marcadores.Count)
            {
                Marcador marcAcutial = s.Marcadores.ElementAt(cont);
                if (m.Equals(marcAcutial))
                {
                    Marc = marcAcutial;
                    if (Marc.Nombre == "IVR_MORA_TEMPRANA" || Marc.Nombre == "MORA_TEMPRANA_PREDICTIVO")
                        UnidadNegocio = tipoRecurso.marcadorTemprana;
                    else
                        UnidadNegocio = tipoRecurso.marcadorTardia;
                    bandera = true;
                }
                cont++;
            }
            if (!bandera)
                Marc = m;
            return bandera;
        }

        #region estadoLotes

        /**
         * @fn  public static int estadoLoteParaEliminar(string pEstaodLote)
         *
         * @brief   Funciones que se utilizan para determinar estados del lote
         *          al momento de eliminar 
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pEstaodLote Extract de imacros al momento de chequear estado.
         *
         * @return  Estado del lote representado con un int.
         *          1- eliminado
         *          0- completo (Error al eliminar el lote sigue)
         *          -1- Otra opcion
         */

        public static int estadoLoteParaEliminar(string pEstaodLote)
        {
            pEstaodLote = pEstaodLote.Replace(System.Environment.NewLine, "");
            if (pEstaodLote.Trim() == "Eliminando" || pEstaodLote.Trim() == "NODATA" || pEstaodLote.Trim() == "#EANF#")
                return 1;
            else if (pEstaodLote.Trim() == "Completo")
                return 0;
            else
                return -1;
        }

        /**
         * @fn  public static int estadoLoteImportacion(string pEstaodLote)
         *
         * @brief   Funcion que determina el estado del lote al mometno de 
         *          importarlo a inconsert
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pEstaodLote Extract de imacros al momento de chequear estado.
         *
         * @return  Estado del lote representado con un int.
         */

        public static int estadoLoteImportacion(string pEstaodLote) // 1-completo, 0- incompleto,-1 abortado, -2 no importo, 2 cargando
        {
            pEstaodLote = pEstaodLote.Replace(System.Environment.NewLine, "");
            if (pEstaodLote.Trim() == "Completo")
                return 1;
            else if (pEstaodLote.Trim() == "Abortado")
                return -1;
            else if (pEstaodLote.Trim() == "NODATA")
                return -2;
            else if (pEstaodLote.Trim() == "En ejecución")
                return 2;
            else
                return 0;
        }
        #endregion

        /**
         * @fn  public void cargarScore()
         *
         * @brief   Funcion utilizada para insertarle el score a un lote
         *          antes de ser cargado a inconsert.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */

        public void cargarScore()
        {
            tolls t = tolls.T;
            t.matarProceso("EXCEL");
            string nuevaUrl = @"\\CGMSERVER\Archivos\DOCUMENTOS C+G\CGM & ASOC\AREA ESTRATEGIA Y COBRANZAS\Allyson\Cargas inConcert\Automaticos\lotesConScore\";
            StreamWriter file = new StreamWriter(nuevaUrl + Exc.Direccion.Nombre + Exc.Direccion.Extencion, false);
            List<string[]> excel = Exc.leerExcel();
            if (excel.ElementAt(0)[8] == "Score")
            {
                baseDatos bd = baseDatos.Bd;
                int cantCamp = 0;
                int id = 0;
                OrdenarBigFish(excel);
                foreach (string[] f in excel)
                {
                    cantCamp = f.Count();
                    int.TryParse(f[0].Substring(3), out id);
                    for (int i = 0; i < cantCamp; i++)
                    {
                        if (i == 8 && id != 0)
                            file.Write(bd.buscarScore(id) + ";");
                        else
                            file.Write(f[i] + ";");
                    }

                    file.WriteLine("");
                }
                file.Close();

                Exc = new Excel { Direccion = new URL(Exc.Direccion.Nombre + Exc.Direccion.Extencion, nuevaUrl) };
            }
        }

        /**
         * @fn  public void OrdenarBigFish(List<string[]> arrDesordenado)
         *
         * @brief   Ordena un excel desordenado
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   arrDesordenado  Excel a ordenar.
         */

        public void OrdenarBigFish(List<string[]> arrDesordenado)
        {
            bool hubo_cambio;
            double monto, montoSigiente;
            int k = 0;
            string[] CamposSigientes, Campos, temp;
            int pos = arrDesordenado.Count - 1; //Guarda en pos la longitud del array
            do
            {
                hubo_cambio = false;
                for (int i = 0; i < pos; i++)
                {
                    Campos = arrDesordenado[i];
                    CamposSigientes = arrDesordenado[i + 1];
                    if (Campos.Count() == 29 && CamposSigientes.Count() == 29)
                    {
                        if (double.TryParse(Campos[7], out monto) && double.TryParse(CamposSigientes[7], out montoSigiente))
                        {
                            if (monto < montoSigiente)
                            {
                                temp = arrDesordenado[i];
                                arrDesordenado[i] = arrDesordenado[i + 1];
                                arrDesordenado[i + 1] = temp;
                                hubo_cambio = true;
                                k = i;
                            }
                        }
                    }
                }
                pos = k;
            }
            while (hubo_cambio);
        }

        #region overrides
        public override bool Equals(object obj)
        {
            Lote l = obj as Lote;
            if (l == null)
                return false;
            return l.Nombre.Equals(this.Nombre);
        }
        public override string ToString()
        {
            if (Marc != null && Frec != null)
                return "Nombre: " + Nombre + "Estado: " + Estado.ToString() + "Frecunacia: " + Frec.ToString() + "Eliminar anterires: " + EliminarLotesAnt + " Marcador: " + Marc.Nombre;
            if (Marc != null)
                return "Nombre: " + Nombre + "Estado: " + Estado.ToString() + "Eliminar anterires: " + EliminarLotesAnt + " Marcador: " + Marc.Nombre;
            if (Frec != null)
                return "Nombre: " + Nombre + "Estado: " + Estado.ToString() + "Frecunacia: " + Frec.ToString() + "Eliminar anterires: " + EliminarLotesAnt;
            return "Nombre: " + Nombre + "Estado: " + Estado.ToString() + "Eliminar anterires: " + EliminarLotesAnt;
        }
        public bool Equals(Lote other)
        {
            return other.Nombre.Equals(this.Nombre);
        }

        public int CompareTo(Lote other)
        {
            return this.Nombre.CompareTo(other.Nombre);
        }
        #endregion
    }
}
