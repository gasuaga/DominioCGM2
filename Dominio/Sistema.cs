using iMacros;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
{
    /**
     * @class   Sistema
     *
     * @brief   La clase sistema se va a encargar
     *          de las cosas generales y lugar donde
     *          se guarden todas las listas de los 
     *          objetos.
     *          Implementa un algoritmo de singleton.
     *
     * @author  WINMACROS
     * @date    17/07/2017
     */

    class Sistema
    {
        #region variables
        /** @brief   Instancia estatica y unica del sistema */
        private static Sistema sis;
        /**
         * @enum    resultAgregarLote
         *
         * @brief   Los resultados posibles al momento de agregar
         *          un lote al sitema.
         */
        public enum resultAgregarLote
        {
            agrego, elLoteYaExiste, nombreMalEscrito, noTieneCampana,
            errorEnExtencion, noExisteCampanaSistema, excelDeficiente,
            nombreDistintoCampana, nombreLargo
        };

        /**
         * @enum    navegadores
         *
         * @brief   Los tipos de navegadores que se puede
         *          ejecutar imacros.
         *          
         */

        public enum navegadores
        {
            ///< silent emulador de imacros
            ///< cr google chrome
            ///< fx Firefox
            ///< ie Internet explorer
            fx, ie, cr, silent
        };

        #region direcciones
        public string direccion =
                     "E:/Macros/CartarLoteInconsert/";
        public string urlDataSourcer =
                     @"E:\lotesCargadosConMacro\";
        public string direccionBusquedaLote =
                     "E:/Macros/BuscarRecursosParaEnviar/";
        public string dirMacroBats =
                     "E:/Macros/Bat's/";
        #endregion

        #endregion

        #region property
        public List<Lote> Lotes { get; set; }
        public List<Frecuencia> Frec { get; set; }
        public List<Marcador> Marcadores { get; set; }

        /**
         * @property    public static Sistema Sis
         *
         * @brief   Si el sistema no fue llamado nunca
         *          desde que inicio el programa iguala 
         *          sistema al contructor.
         *
         * @return  La unica intancia de sistema.
         */

        public static Sistema Sis
        {
            get
            {
                if (sis == null)
                    sis = new Sistema();
                return sis;
            }
        }
        public App m_app { get; set; }
        public Status S { get; set; }
        #endregion

        private Sistema()
        {
            tolls t = tolls.T;
            Lotes = new List<Lote>();
            Frec = new List<Frecuencia>();
            Marcadores = new List<Marcador>();
            m_app = new App();
            precarga();
        }

        /**
         * @fn  private void precarga()
         *
         * @brief   Precarga las listas del sistema.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         */

        private void precarga()
        {
            baseDatos bd = baseDatos.Bd;
            Automatizaciones a = Automatizaciones.A;
            // bd.Prueba();
            //a.activosSms();
            List<Marcador> mar = bd.cargarMarcadores();
            Marcadores.AddRange(mar);

            List<Lote> lotess = bd.cargaLotesUnaSemana();
            Lotes.AddRange(lotess);
            accionesCodigo("Precarga la fecha del Score", bd.precargaFecha().ToString());
            tolls.FechaScore = bd.precargaFecha();

        }

        #region iniciarYlotear

        /**
         * @fn  public void iniciarYLogear(string pNavegador)
         *
         * @brief   Son los metodos utilizados para ejecutar
         *          imacros con cualquien navegador e iniciar
         *          secion en inconsert.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         * @param   pNavegador  The navegador.
         */

        public void iniciarYLogear(string pNavegador)
        {
            ejecutarMacroIniciar(m_app, "-" + pNavegador);
            m_app.iimOpen();
            ejecutarMacro(m_app, direccion + "Logearse.iim", "Logearse", true);
        }
        public void iniciarYLogearAux(string pNavegador)
        {
            ejecutarMacroIniciar(m_app, "-" + pNavegador);
            m_app.iimOpen();
            ejecutarMacro(m_app, direccion + "Logearse.iim", "Logearse", true);
        }
        #endregion

        /**
         * @fn  public void iniciarSubida()
         *
         * @brief   Es el metodo que toma todos los lotes
         *          que ahi en el sistema para cargar.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         */

        public void iniciarSubida()
        {
            m_app = new App();
            tolls t = tolls.T;

            List<Lote> lotesParaCargar = t.lotesPara(Lote.tipoEstado.ParaCargar);
            List<Marcador> campanasDistintasParaEliminar = t.campanasParaEliminarDistintas(lotesParaCargar);
            List<Marcador> campanasDistintas = t.campanasDistintaslotes(lotesParaCargar);

            iniciarSubidaAux(campanasDistintas, campanasDistintasParaEliminar, lotesParaCargar);
        }

        private void iniciarSubidaAux(List<Marcador> pCampanasDisitntas, List<Marcador> pCampanasParaEliminar, List<Lote> pLotesParaCargar)
        {
            accionesCodigo("Iniciar subida");
            accionesCodigo("------------------");
            accionesCodigo("------------------");
            accionesCodigo("------------------");
            /* try {
                 foreach (Lote lo in pLotesParaCargar)
                     lo.cargarScore();
             }
             catch (Exception) { }*/

            tolls t = tolls.T;
            baseDatos bd = baseDatos.Bd;

            Marcador.cargarLoteAMotor();

            foreach (Marcador m in pCampanasDisitntas)
                m.asignarFrec();//Asigna las frecuencias separando todos los lotes por campaña que existan

            foreach (Lote l in pLotesParaCargar)
                bd.agregarLote(l);//Agrega el lote a la base de datos

            iniciarYLogear(Sistema.navegadores.fx.ToString());

            foreach (Marcador m in pCampanasParaEliminar)
                m.eliminarLotes();

            List<Lote> lotesParaEliminar = t.lotesPara(Lote.tipoEstado.paraEliminar);
            foreach (Lote l in lotesParaEliminar)
                eliminarImportacion(l, 0);

            m_app.iimClose();

            iniciarYLogear(Sistema.navegadores.silent.ToString());
            foreach (Lote l in pLotesParaCargar)
                crearNuevaImportacion(l);

            m_app.iimClose();

            iniciarYLogear(Sistema.navegadores.fx.ToString());
            foreach (Marcador m in pCampanasDisitntas)
                m.cargarLotesAlMarcador();

            m_app.iimClose();

            accionesCodigo("------------------");
            accionesCodigo("------------------");
            accionesCodigo("Finalizo la carga de los lotes solicitados");
            accionesCodigo("------------------");
            accionesCodigo("------------------");
        }

        /**
         * @fn  public Marcador crearNuevoMotor(Lote pLote)
         *
         * @brief   Crear nuevo marcador para el lote
         *          que intento cargar.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         * @param   pLote   Lote que solicito el
         *                  motor.
         *
         * @return  Marcador para el pLote.
         */
        public Marcador crearNuevoMotor(Lote pLote)
        {
            return new Marcador(pLote.Exc.campaniaAsignada(), tolls.T.tipoMarcador(pLote.Exc.Direccion));

        }

        /**
         * @fn  public resultAgregarLote agregarLote(Lote pLote)
         *
         * @brief   Validador al momento de agregar un lote al sistema.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         * @param   pLote   Lote a validar.
         *
         * @return  Resultado de validacion.
         */

        public resultAgregarLote agregarLote(Lote pLote)
        {

            tolls t = tolls.T;
            resultAgregarLote ret = resultAgregarLote.agrego;
            Lote.tipoLote tipoLoteXNom = t.tipoLoteXNombreArchivo(pLote.Exc.Direccion);
            bool excel = pLote.Exc.validarExcel();
            if (!excel)
                ret = resultAgregarLote.excelDeficiente;
            else if (pLote.Exc.Direccion.Extencion != ".csv")
                ret = resultAgregarLote.errorEnExtencion;
            else if (tipoLoteXNom == Lote.tipoLote.errorCampana)
                ret = resultAgregarLote.nombreMalEscrito;
            else if (tipoLoteXNom != pLote.LoteTipo)
                ret = resultAgregarLote.nombreDistintoCampana;
            else if (pLote.Nombre.Length > 32)
                ret = resultAgregarLote.nombreLargo;
            else if (!pLote.asignarMarcador())
                ret = resultAgregarLote.noExisteCampanaSistema;
            else if (existeLote(pLote))
                ret = resultAgregarLote.elLoteYaExiste;
            else
                Lotes.Add(pLote);
            accionesCodigo("Se intenta cargar el lote:" + pLote.ToString() + " Al sistema. ", ret.ToString());
            return ret;
        }

        /**
         * @fn  public bool existeLote(Lote pLote)
         *
         * @brief   Si existe el lote en el sistema.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         * @param   pLote   Lote.
         *
         * @return  True si existe, false si no.
         */

        public bool existeLote(Lote pLote)
        {
            bool ret = false;
            int cont = 0;
            while (cont < Lotes.Count && !ret)
            {
                Lote loteActual = Lotes.ElementAt(cont);
                if (pLote.Equals(loteActual))
                    ret = true;
                cont++;
            }
            return ret;
        }

        /**
         *
         * @brief   Todas las funciones relacionadas con dejar logs.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         */
        #region  logs

        public Status ejecutarMacroIniciar(App pApp, string pNavegados)
        {
            Status stat = pApp.iimInit(pNavegados, true, "", "", "", 300);
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Macros\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - Se a iniciado el navegador: " + pNavegados + " Que dio como resultado: " + stat.ToString());
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
            return stat;
        }

        public Status ejecutarMacro(App pApp, string pEjecutar, string pAccion, bool pTipoAccion)
        {
            Status stat;
            if (pTipoAccion)
                stat = pApp.iimPlay(pEjecutar, 60);
            else
                stat = pApp.iimPlayCode(pEjecutar);
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Macros\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + pAccion + "Que dio como resultado: " + stat.ToString());
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
            return stat;
        }
        public void ejecutarMacro(App pApp, string pVariable, string pValor)
        {
            try
            {
                Status stat = pApp.iimSet(pVariable, pValor);
                StreamWriter sW = new StreamWriter(@"E:\Logs\Macros\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + " La variable " + pVariable + " Dentro de imacros tomo el valor de: " + pValor + " Y dio como resultado: " + stat.ToString());
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
        }
        public string ejecutarMacroExtract(App pApp, int pPos, string pReferencia)
        {
            string stat = pApp.iimGetExtract(pPos).Split('[').ElementAt(0);
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Macros\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + " Se genero un extract que dio como resultado:  " + stat + " -  Para la accion: " + pReferencia);
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
            return stat;
        }
        public void accionoBaseDatos(string pAccion, string pResultado)
        {
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Bd\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + pAccion + " Que dio como resultado: " + pResultado);
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
        }

        public void accionesCodigo(string pAccion, string pResultado)
        {
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Codigo\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + pAccion + " Que dio como resultado: " + pResultado);
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
        }
        public void accionesCodigo(string pAccion)
        {
            try
            {
                StreamWriter sW = new StreamWriter(@"E:\Logs\Codigo\" + DateTime.Today.ToString("yyyyMMdd") + ".txt", true);
                sW.WriteLine(DateTime.Now.ToLongTimeString() + " - " + pAccion);
                sW.Close();
                sW.Dispose();
            }
            catch (Exception) { }
        }

        #endregion
    }
}
