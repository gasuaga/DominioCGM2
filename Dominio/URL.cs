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
     * @class   URL
     *
     * @brief   Es quien contiene toda la 
     *          informacion sobre la 
     *          url del excel
     *
     * @author  WINMACROS
     * @date    14/07/2017
     */

    public class URL
    {
        #region propery
        public string Direccion { get; set; }
        public string Nombre { get; set; }
        public string Extencion { get; set; }
        #endregion

        #region constructores
        public URL(HttpPostedFile pDireccion)
        {
            Sistema sis = Sistema.Sis;
            string nombreTemp = Path.GetFileName(pDireccion.FileName);
            Nombre = nombreTemp.Substring(0, nombreTemp.Length - 4);
            Extencion = Path.GetExtension(nombreTemp);
            Direccion = sis.urlDataSourcer;
            try
            {
                if (File.Exists(Direccion + Nombre + Extencion))
                    File.Delete(Direccion + Nombre + Extencion);
                pDireccion.SaveAs(Direccion + Nombre + Extencion);
            }
            catch (Exception)
            {
                new URL(Nombre + 1, pDireccion);
            }
        }
        public URL(string pNombre, string pDireccion)
        {
            Nombre = pNombre.Substring(0, pNombre.Length - 4);
            Direccion = pDireccion;
            Extencion = pNombre.Substring(pNombre.Length - 4);
        }
        public URL(string pNombre, HttpPostedFile pDireccion)
        {
            Sistema sis = Sistema.Sis;
            string nombreTemp = Path.GetFileName(pDireccion.FileName);
            Nombre = pNombre;
            Extencion = Path.GetExtension(nombreTemp);
            Direccion = sis.urlDataSourcer;
            try
            {
                if (File.Exists(Direccion + Nombre + Extencion))
                    File.Delete(Direccion + Nombre + Extencion);
                pDireccion.SaveAs(Direccion + Nombre + Extencion);
            }
            catch (Exception)
            {
                new URL(Nombre + 1, pDireccion);
            }
        }
        #endregion

        #region  overrides
        public override string ToString()
        {
            return @Direccion + @Nombre + Extencion;
        }
        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            URL u = obj as URL;
            if (u == null) return false;
            return u.ToString().Equals(this.ToString());
        }

        public bool Equals(URL other)
        {
            return other.ToString().Equals(this.ToString());
        }
        #endregion
    }
}
