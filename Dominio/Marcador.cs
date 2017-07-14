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
        public List<loteMarcador> HistorialLotes { get; set; }
        /**
         * @property    public List<loteMarcador> LotesActivos
         *
         * @brief   Obtiene o setea los lotes que se encuentran
         *          andando en el motor
         *
         * @return  The lotes activos.
         */
        public List<loteMarcador> LotesActivos { get; set; }
        /**
         * @property    public List<loteMarcador> LotesParaCargar
         *
         * @brief   Obtiene o setea los lotes que se 
         *          van a cargar al motor
         *
         * @return  The lotes para cargar.
         */
        public List<loteMarcador> LotesParaCargar { get; set; }
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
    }
}
