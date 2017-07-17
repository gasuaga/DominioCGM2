using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iMacros; // usign del controlador de imacros
namespace Dominio
{
    /**
     * @class   Marcador
     *
     * @brief   Es la clase relacionada con todas las tareas de los 
     *          marcadores dentro de inconsert, tanto eliminar insertar modificar
     *          lotes sobre uno o mas marcadores (Motores de marcacion)
     *
     * @author  WINMACROS
     * @date    13/07/2017
     */

    class Marcador
    {
        #region Variables
        /**
         * @enum    tipoMarcador
         *
         * @brief   Representa que tipo de marcador es de recurso (vb, mvp) o marcador.
         */
        public enum tipoMarcador { vb, marcador, mvp };

        /**
         * @enum    estadoMotor
         *
         * @brief   Indica el estado del motor en el momento actual.
         *          Activo = si el motor se encuentra llamano.
         *          paraLimpiar= si el motor se encuentra
         *          para que se eliminen los lotes que tiene.
         *          detenido = el motor se encuentra parado
         */
        public enum estadoMotor { activo, denenido, paraLimpiar };
        #endregion

        #region Propertys
        public string Nombre { get; set; }
        public tipoMarcador Tipo { get; set; }
        public estadoMotor Estado { get; set; }

        /**
         * @property    public List<loteMarcador> HistorialLotes
         *
         * @brief   Obtiene o setea los lotes que estubieron en ese marcador
         *
         * @return  The historial lotes.
         */
        public List<LoteMarcador> HistorialLotes { get; set; }
        /**
         * @property    public List<loteMarcador> LotesActivos
         *
         * @brief   Obtiene o setea los lotes que se encuentran
         *          andando en el motor
         *
         * @return  The lotes activos.
         */
        public List<LoteMarcador> LotesActivos { get; set; }
        /**
         * @property    public List<loteMarcador> LotesParaCargar
         *
         * @brief   Obtiene o setea los lotes que se 
         *          van a cargar al motor
         *
         * @return  The lotes para cargar.
         */
        public List<LoteMarcador> LotesParaCargar { get; set; }
        /**
         * @property    private List<string> LotesParaSacarReporte
         *
         * @brief   Si el motor es de recurso, antes de eliminar
         *          los lotes guarda los nombres para luego sacar
         *          un reporte de contactavilidad
         *
         * @return  The lotes para sacar reporte.
         */
        private List<string> LotesParaSacarReporte { get; set; }
        #endregion

        #region contructores
        public Marcador(string pNombre, tipoMarcador pTipo)
        {
            Nombre = pNombre;
            Tipo = pTipo;
            HistorialLotes = new List<loteMarcador>();
            LotesActivos = new List<loteMarcador>();
            LotesParaCargar = new List<loteMarcador>();
            LotesParaSacarReporte = new List<string>();
        }
        public Marcador(string pNombre)
        {
            Nombre = pNombre;
            HistorialLotes = new List<loteMarcador>();
            LotesActivos = new List<loteMarcador>();
            LotesParaCargar = new List<loteMarcador>();
            LotesParaSacarReporte = new List<string>();
        }
        #endregion
        
        #region irAlMotor

        /**
         * @fn  public int irAlMortor()
         *
         * @brief   Va al motor y extrae la pocicion en 
         *          la que se encuentra dentro de inconsert
         *
         * @author  WINMACROS
         * @date    13/07/2017
         *
         * @return  La pocicion del motor.
         */
        public int irAlMortor()
        {
            Sistema s = Sistema.Sis;
            bool bandera = true;
            int posPosible = 2;//arranca de la pocicion dos (la 1 son los cabezales de la tabla)
            while (bandera) //minetras no encuentre al motor que tiene el mismo nombre que this sigue buscando
            {
                s.ejecutarMacro(s.m_app, "nombreMotorC", Nombre);
                s.ejecutarMacro(s.m_app, "posC", posPosible.ToString());
                s.ejecutarMacro(s.m_app, s.direccion + "irAlMotor.iim", "Ir al motor", true);//se pociciona sobre el motor
                string extract = s.ejecutarMacroExtract(s.m_app, 2, "Nombre del motor antes de entrar");//extrae ek binvre dek nitir
                if (extract == this.Nombre)
                    bandera = false;
                else
                    posPosible++;
                if (s.S != iMacros.Status.sOk)
                    posPosible = 2;
            }
            return posPosible;
        }
        /**
         * @fn  public void irAlMortor(int pPos)
         *
         * @brief   Va al motor ubicado en una pocicion
         *
         * @author  WINMACROS
         * @date    13/07/2017
         *
         * @param   pPos    La pocicion del motor.
         */
        public void irAlMortor(int pPos)
        {
            Sistema s = Sistema.Sis;
            s.ejecutarMacro(s.m_app, "nombreMotorC", Nombre);
            s.ejecutarMacro(s.m_app, "posC", pPos.ToString());
            s.ejecutarMacro(s.m_app, s.direccion + "irAlMotor.iim", "Ir al motor", true);
        }
        /**
         * @fn  public int irAlMotorYDetener()
         *
         * @brief   Va al motor y lo detiene si esta 
         *          en funcionamiento
         *
         * @author  WINMACROS
         * @date    13/07/2017
         *
         * @return  La pocicion del motor que detiene.
         */
        public int irAlMotorYDetener()
        {
            Sistema s = Sistema.Sis;
            int pos = irAlMortor();
            String estadoMotor = s.m_app.iimGetExtract(1);
            detenerMotor(estadoMotor, pos);
            return pos;
        }
        /**
         * @fn  private void entrarAlMotor(int pPos)
         *
         * @brief   Entra al motor de una pocicion
         *
         * @author  WINMACROS
         * @date    13/07/2017
         *
         * @param   pPos    Pocicion del motor
         */
        private void entrarAlMotor(int pPos)
        {
            Sistema s = Sistema.Sis;
            s.ejecutarMacro(s.m_app, "TAG POS=" + pPos + " TYPE=td ATTR=IDX:10", "Entrar al motor", false);
            esperarQueElMotorCarge();

        }
        #endregion

        /**
         * @fn  private void esperarQueElMotorCarge()
         *
         * @brief   Una vez dentor del motor esperar a que el motor carge todos los lotes 
         *          o tambien cuando se importa un lote al motor para esperear que carge
         *          los contactos
         *
         * @author  WINMACROS
         * @date    13/07/2017
         */
        private void esperarQueElMotorCarge()
        {
            bool bandera = false;
            Sistema s = Sistema.Sis;
            do
            {
                s.ejecutarMacro(s.m_app, "TAG POS=1 TYPE=div ATTR=class:formWaitingMessage EXTRACT=HTM", "Fijarse en el div que tiene los mensajes si tiene algo", false);
                string ext = s.ejecutarMacroExtract(s.m_app, 0, "Mensajes del div de mensajes");
                if (ext.Contains("none") || ext.Contains("#EANF#") || ext.Contains("NODATA"))
                    bandera = false;
                else
                {
                    bandera = true;
                    s.ejecutarMacro(s.m_app, "WAIT SECONDS=2", "Esteperar 2 segundos", false);
                }
            } while (bandera);
        }

        #region eliminarLotes

        /**
         * @fn  public void eliminarLotes()
         *
         * @brief   Elimina todos los lotes dentro del
         *          marcador.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */
        public void eliminarLotes()
        {
            Sistema s = Sistema.Sis;
            int pos = irAlMotorYDetener();
            s.ejecutarMacro(s.m_app, "TAG POS=" + pos + " TYPE=TD ATTR=idx:4 EXTRACT=TXT", "Extrae campaña asociada al proceso", false);
            string campana = s.ejecutarMacroExtract(s.m_app, 0, "Campaña del motor");
            s.m_app.iimClose();// cierro el firefox
            s.iniciarYLogear(Sistema.navegadores.silent.ToString()); // abre el imacros con el navegador nativo
            irAlMortor(pos);
            entrarAlMotor(pos);
            int lotes = lotesParaEliminar(); //cantidad de lotes dentor del motor
            if (LotesParaSacarReporte.Count >= 1)// si la lista de reportes no esta vacia saca los reportes
                tolls.T.obtenerReporteLote(LotesParaSacarReporte, campana);          // que se encuentran dentor de la lista
            string rutaReportes = @"E:\reportes\"; //Direccion donde se guarda los reportes
            string[] archivos = System.IO.Directory.GetFiles(rutaReportes, "*.bak"); // se borran todos los archivos .bak 
            foreach (string arch in archivos)                                        //son descargados junto a los pdf
                System.IO.File.Delete(arch);

            s.m_app.iimClose(); //se cierra el imacros
            s.iniciarYLogear(Sistema.navegadores.fx.ToString()); //Se abre el firefox
            irAlMortor(pos);
            entrarAlMotor(pos);
            if (lotes > 0)
                for (int i = 0; i < lotes; i++)
                    eliminarLoteMotor(); //elimina la x cantidad de lotes que recolecto anteriormente

        }

        /**
         * @fn  private int lotesParaEliminar()
         *
         * @brief   Cuenta cuantos lotes estan incluidos en el motor.
         *          Si el motor es de recursos guarda los nombre de los 
         *          lotes en una lista para sacar reporte
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @return  Cantidad de lotes dentro del motor (Int).
         */

        private int lotesParaEliminar()
        {
            Sistema s = Sistema.Sis;
            baseDatos bd = baseDatos.Bd;
            bool bandera = false;

            int pos = 2; // Arranca de la pocicion 2 ya que la primera es el cabezal de la tabla

            do
            {
                Status sta = s.ejecutarMacro(s.m_app, "TAG SELECTOR=HTML>BODY>DIV:nth-of-type(4)>DIV>DIV:nth-of-type(2)>DIV:nth-of-type(2)>DIV:nth-of-type(2)>DIV>FORM>DIV:nth-of-type(2)>FIELDSET:nth-of-type(3)>DIV:nth-of-type(2)>DIV>TABLE>TBODY>TR:nth-of-type(2)>TD>DIV>DIV>TABLE>TBODY>TR:nth-of-type(" + pos.ToString() + ")>TD:nth-of-type(2) EXTRACT=TXT", "Reccore todos los lotes del motor", false);
                if (sta != iMacros.Status.sOk)
                    bandera = false; //si el statis de imacros no es ok es porque llego al final de los lotes del motor
                else
                {
                    string extract = s.ejecutarMacroExtract(s.m_app, 0, "Nombre del lote dentro del motor");
                    if (!extract.Contains("#EANF#"))//si no contiene vasura el extract 
                    {
                        bandera = true;
                        if (this.Tipo == tipoMarcador.mvp || this.Tipo == tipoMarcador.vb)
                            LotesParaSacarReporte.Add(extract);// agrega a lotes para sacar reporte
                        if (!eliminarLoteAnterior(extract)) 
                            s.Lotes.Add(elLoteNoExiste(extract));
                        //si no encuentra el lote dentro de la lista de lotes
                        //activos crea un nevo lote solo con el nombre y el estado
                        //para eliminar
                        pos++;
                    }
                    else
                        return pos - 2;
                }
            } while (bandera);
            return pos - 2;
        }

        /**
         * @fn  private Lote elLoteNoExiste(string pNombre)
         *
         * @brief   Si el lote no existe en el sistema crea un lote
         *          con solo el nombre y el estado para eliminar.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   Nombre del lote.
         *
         * @return  Un lote nuevo el cual se va a eliminar.
         */

        private Lote elLoteNoExiste(string pNombre)
        {
            baseDatos bd = baseDatos.Bd;
            Lote l = new Lote(pNombre);
            l.Marc = this;
            bd.eliminarLoteMotor(l);
            l.Estado = Lote.tipoEstado.paraEliminar;
            return l;
        }

        /**
         * @fn  private bool eliminarLoteAnterior(string pNombre)
         *
         * @brief   Elimina el lote dentro del sistema y la base de datos.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pNombre Nombre del lote.
         *
         * @return  Ture si existe el lote en el sitema false de loc ontrario.
         */

        private bool eliminarLoteAnterior(string pNombre)
        {
            bool bandera = false;
            int cont = 0;
            while (!bandera && this.LotesActivos.Count > cont)
            {
                loteMarcador lActual = LotesActivos[cont];
                if (pNombre == lActual.Lot.Nombre)
                {
                    baseDatos bd = baseDatos.Bd;
                    LotesActivos.RemoveAt(cont); //saco el lote de los lotes activos del marcador
                    HistorialLotes.Add(lActual);
                    lActual.Hasta = DateTime.Today;//cambio el hasta porque se elimino hoy 
                    bd.eliminarLoteMotor(lActual); // marco en la base cuando se elimino 
                    lActual.Lot.Estado = Lote.tipoEstado.paraEliminar;
                    bandera = true;
                }
                cont++;
            }
            return bandera;
        }

        #region eliminar lote del motor 

        /**
         * @fn  private void eliminarLoteMotor()
         *
         * @brief   Es el encargado de eliminar el lote que se 
         *          encuentra primero dentro del motor.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */

        private void eliminarLoteMotor()
        {
            esperarQueElMotorCarge();
            Sistema s = Sistema.Sis;
            s.ejecutarMacro(s.m_app, s.direccion + "eliminar lote/VencerLotes.iim", "Vencer el lote del motor", true);
            esperarQueElMotorCarge();
            esperarQueElMotorCarge();
            s.ejecutarMacro(s.m_app, s.direccion + "eliminar lote/EliminarLoteMotor.iim", "Eliminar el lote del motor", true);
            esperarQueElMotorCarge();
            esperarQueElMotorCarge();
        }
        #endregion


        #endregion

        /**
         * @fn  public void asignarFrec()
         *
         * @brief   Es el encargado de agarrar todos los lotes que se van a cargar
         *          y crear en base a la frecuencia de cada lote su base de 
         *          contactavilidad y la frecuencia.
         *          Busca el minimo comuin multiplo entre los n frecuencias distintas que agrege
         *          y en base a eso calgula la frecuencia 
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */

        public void asignarFrec()
        {
            baseDatos bd = baseDatos.Bd;
            if (LotesParaCargar.Count > 1)
            {
                Sistema s = Sistema.Sis;
                s.accionesCodigo("Asignar frecuencia");

                int[] numeros = new int[LotesParaCargar.Count];

                for (int i = 0; i < LotesParaCargar.Count; i++)
                {
                    numeros[i] = LotesParaCargar.ElementAt(i).Lot.Frec.Id;
                }

                int numMultiplo = minimoComunMultiplo(numeros);
                s.accionesCodigo("El minimo comun multiplo de las frecuencias ingresadas es=" + numMultiplo);

                int numMax = 50000 * numMultiplo;

                int aux = 0;

                s.accionesCodigo("Numero maximo", numMax.ToString());

                foreach (loteMarcador lM in LotesParaCargar)
                {

                    aux = numMultiplo / lM.Lot.Frec.Id;
                    s.accionesCodigo("Prioridad de lote = " + aux);
                    lM.Lot.Frec.PrioridadLote = aux;

                    lM.Lot.Frec.BaseContactacion = (numMax / aux) - lM.Lot.Exc.leerExcel().Count;

                    s.accionesCodigo("Base de contactacion = " + lM.Lot.Frec.BaseContactacion);

                    s.accionesCodigo("Asignar frecuencia al lote = " + lM.Lot.Nombre, lM.Lot.Frec.ToString());

                }
            }
            else
                LotesParaCargar.ElementAt(0).Lot.Frec = new Frecuencia(100, 10, 1);
        }

        /**
         * @fn  public int minimoComunMultiplo(int[] pNumeros)
         *
         * @brief   Calcula el minimo comun multiplo dentro de un 
         *          array desordenado de int's.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pNumeros    Los numeros cuales se tiene que hayar el minimo comun multiploi.
         *
         * @return  El numero cual se puede dividir por todos los pNumeros.
         */

        public int minimoComunMultiplo(int[] pNumeros)
        {

            int mayor = numeroMayor(pNumeros); // el nuimero si o si tiene que ser igual o mayor al maximo
            bool bandera = true;
            int multiplo = mayor;
            do
            {
                bandera = true;
                foreach (int i in pNumeros)
                {
                    if (multiplo % i != 0)
                        bandera = false;
                }
                if (!bandera)
                    multiplo++;
            } while (!bandera);
            return multiplo;
        }

        /**
         * @fn  public int numeroMayor(int[] pNumeros)
         *
         * @brief   Busca el numero mayor del array 
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   pNumeros    Los numeros donde tiene que buscar.
         *
         * @return  El numero mayor de los pNumeros.
         */

        public int numeroMayor(int[] pNumeros)
        {
            int mayor = int.MinValue;
            foreach (int i in pNumeros)
            {
                if (i > mayor)
                    mayor = i;
            }
            return mayor;
        }

        
        /**
         * @fn  private void detenerMotor(string pEstado, int pos)
         *
         * @brief   Detiene el motor en inconsert
         *
         * @author  WINMACROS
         * @date    17/07/2017
         *
         * @param   pEstado Estado actual del motor.
         * @param   pos     Pos del motor en la tabla.
         */

        private void detenerMotor(string pEstado, int pos)
        {
            baseDatos bd = baseDatos.Bd;
            Sistema s = Sistema.Sis;
            if (pEstado == "RUNNING")//si el estado acutal del motor es en andando se lo frena
            {
                s.ejecutarMacro(s.m_app, "TAG POS=" + pos + " TYPE=td ATTR=IDX:9", "Detener el motor si esta en running", false);
                irAlMortor();
            }
            Estado = estadoMotor.denenido;
            bd.cambiarEstadoMotor(this);
        }

        /**
         * @fn  public void cargarLotesAlMarcador()
         *
         * @brief   Crea los lotes dentro del motor.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         */

        public void cargarLotesAlMarcador()
        {
            Sistema s = Sistema.Sis;
            baseDatos bd = baseDatos.Bd;
            int pos = irAlMotorYDetener();
            entrarAlMotor(pos);
            LotesParaCargar.Sort();
            foreach (LoteMarcador lM in LotesParaCargar)
            {
                lM.cargarVariables();
                s.ejecutarMacro(s.m_app, s.direccion + "Cargar lote/CargarLoteAMotor.iim", "Cargar el lote: " + lM.Lot.Nombre + " Al motor", true);
                esperarQueElMotorCarge();
                lM.cargarVariables();
                s.ejecutarMacro(s.m_app, s.direccion + "Cargar lote/CrearLoteDentroMotor.iim", "Crear un nuevo lote dentro del motor", true);
                esperarQueElMotorCarge();
                s.ejecutarMacro(s.m_app, "TAG POS=1 TYPE=BUTTON:BUTTON ATTR=ID:btnSave CONTENT=Guardar", "Guardar cambios en el motor", false);
                esperarQueElMotorCarge();
                lM.Lot.Estado = Lote.tipoEstado.Activo;
                LotesActivos.Add(lM);
                bd.cambiarEstadoLote(lM.Lot);
                s.accionesCodigo("Cargar el lote: " + lM.Lot.Nombre + " En el motor", "Completo");
                bd.agregarLoteEnMarcador(lM);
            }
            LotesParaCargar = LotesParaCargar.Where(l => !LotesActivos.Contains(l)).ToList();
            darlePLayMotor();
        }

        /**
         * @fn  public void darlePLayMotor()
         *
         * @brief   Da play al motor.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         */

        public void darlePLayMotor()
        {
            Sistema s = Sistema.Sis;
            baseDatos bd = baseDatos.Bd;
            s.ejecutarMacro(s.m_app, "Datos", Nombre);
            s.ejecutarMacro(s.m_app, s.direccion + "Cargar lote/PlayLote.iim", "Darle play al motor", true);
            Estado = estadoMotor.activo;
            bd.cambiarEstadoMotor(this);
        }

        /**
         * @fn  public static void cargarLoteAMotor()
         *
         * @brief   Carga los lotes que hay en sistema
         *          y los que son del motor se los asigna
         *          a la lista loter para cargar.
         *
         * @author  WINMACROS
         * @date    17/07/2017
         */

        public static void cargarLoteAMotor()
        {
            Sistema s = Sistema.Sis;
            tolls t = tolls.T;
            s.accionesCodigo("Carga de lotes a los motores");
            List<Lote> lot = t.lotesPara(Lote.tipoEstado.ParaCargar);
            DateTime loteDesde = DateTime.Today;
            foreach (Marcador m in t.campanasDistintaslotes(lot))
            {
                s.accionesCodigo("Para la campaña: " + m.Nombre + " Estan los lotes: ");
                m.LotesParaCargar = new List<LoteMarcador>();//cuidadod aca reseteo lista cada ves que cargo nuevos lotes por si las deudas
                foreach (Lote l in lot)
                {
                    if (l.Marc.Equals(m))
                    {
                        s.accionesCodigo(l.ToString());
                        LoteMarcador lotMarc = new LoteMarcador(l, loteDesde, LoteMarcador.fechaHastaLote(loteDesde, l));
                        m.LotesParaCargar.Add(lotMarc);
                    }
                }
            }
        }
        

        #region overides
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            Marcador marc = (Marcador)obj;
            if (marc == null)
                return false;
            return this.Nombre.Equals(marc.Nombre);
        }
        public override int GetHashCode()
        {
            return Nombre.GetHashCode();
        }
        #endregion
    }
}
