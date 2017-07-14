using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
{
    /**
     * @class   Frecuencia
     *
     * @brief   En esta clase se guarda todo lo 
     *          relacionado con la frecuencia del lote
     *          la frecuencia indica como va a llamar 
     *          cuando ahi mas de un lote (llama 1 un lote
     *          2 otro lote) 
     *
     * @author  WINMACROS
     * @date    14/07/2017
     */
    class Frecuencia
    {
        #region property
        /**
         * @property    public int Id
         *
         * @brief   Obtiene o setea el id que representara
         *          la cantidad de llamados que tine.
         *
         * @return  The identifier.
         */
        public int Id { get; set; }
        public int BaseContactacion { get; set; }
        public int PrioridadLote { get; set; }
        #endregion

        #region contructores
        public Frecuencia(int pBaseContactacion, int pPrioridadLote, int pId)
        {
            BaseContactacion = pBaseContactacion;
            PrioridadLote = pPrioridadLote;
            Id = pId;
        }
        public Frecuencia(int pId)
        {
            Id = pId;
        }
        #endregion

        #region overrides
        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            Frecuencia frec = (Frecuencia)obj;
            if (frec == null)
                return false;
            return this.Id.Equals(frec.Id);
        }
        public override int GetHashCode()
        {
            return Id.GetHashCode();
        }
        public override string ToString()
        {
            return "Base: " + BaseContactacion + "Prioridad: " + PrioridadLote;
        }
        #endregion
    }
}
