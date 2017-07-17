using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
{
    /**
     * @class   LoteMarcador
     *
     * @brief   Se encarga de guardar el lote con la fecha de vigencia
     *
     * @author  WINMACROS
     * @date    14/07/2017
     */

    public class LoteMarcador : IComparable<LoteMarcador>
    {
        #region propertys
        public Lote Lot { get; set; }
        public DateTime Desde { get; set; }
        public DateTime Hasta { get; set; }
        #endregion

        #region constructores
        public LoteMarcador(Lote pLote, DateTime pDesde, DateTime pHasta)
        {
            Lot = pLote;
            Desde = pDesde;
            Hasta = pHasta;
        }
        public LoteMarcador(Lote pLote)
        {
            Lot = pLote;
        }
        #endregion

        /**
         * @fn  public void cargarVariables()
         *
         * @brief   Carga las variables dentor de imacros con 
         *          los datos necesarios para cargarlo 
         *          al motor
         *
         * @author  WINMACROS
         * @date    14/07/2017
         */

        public void cargarVariables()
        {
            Sistema s = Sistema.Sis;
            s.ejecutarMacro(s.m_app, "nombreMotorC", Lot.Marc.Nombre);
            s.ejecutarMacro(s.m_app, "nombreLoteC", Lot.Nombre);
            s.ejecutarMacro(s.m_app, "fechaIniC", Desde.ToString("yyyy-MM-dd"));
            s.ejecutarMacro(s.m_app, "fechaFinC", Hasta.ToString("yyyy-MM-dd"));
            s.ejecutarMacro(s.m_app, "baseContC", Lot.Frec.BaseContactacion.ToString());
            s.ejecutarMacro(s.m_app, "prioLoteC", Lot.Frec.PrioridadLote.ToString());
            s.ejecutarMacro(s.m_app, "tiempoEsperaC", hallarTiempo());
        }

        private string hallarTiempo()
        {
            tolls t = tolls.T;
            int i = Lot.Exc.leerExcel().Count / 1000;
            return i.ToString();
        }

        /**
         * @fn  public static DateTime fechaHastaLote(DateTime fechaDesde, Lote pLote)
         *
         * @brief   Calcula la fecha de venciomiento del lote.
         *
         * @author  WINMACROS
         * @date    14/07/2017
         *
         * @param   fechaDesde  Fecha de inicio del lote.
         * @param   pLote       Lote.
         *
         * @return  Fecha de vencimiento datetime.
         */

        public static DateTime fechaHastaLote(DateTime fechaDesde, Lote pLote)
        {
            DateTime fecha = new DateTime();
            if (pLote.LoteTipo == Lote.tipoLote.IVR && pLote.UnidadNegocio == Lote.tipoRecurso.marcadorTemprana)
                fecha = fechaDesde;
            else
                fecha = fechaDesde.AddDays(15);
            return fecha;
        }

        #region overrides
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            LoteMarcador l = (LoteMarcador)obj;
            if (l == null)
                return false;
            return l.Lot.Equals(Lot);
        }
        public override int GetHashCode()
        {
            return Lot.GetHashCode();
        }
        public int CompareTo(LoteMarcador other)
        {
            return this.Lot.CompareTo(other.Lot);
        }
        #endregion
    }
}
