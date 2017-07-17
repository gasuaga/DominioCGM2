using BaseDeDatos;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
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

        public Lote buscarLote(string pNombre)
        {
            tolls t = tolls.T;
            Sistema s = Sistema.Sis;
            Lote lot = null;
            s.accionoBaseDatos("Busqueda de lote con el nombre: " + pNombre);
            SqlConnection cn = Coneccion.CrearConeccionSql
                            (CadenaDeConecciones.tiposDeConeccion.paraDominio);
            SqlCommand cmd = new SqlCommand(Query_s.buscarLote, cn);
            try
            {
                Coneccion.AbrirConeccion(cn);
                cmd.Parameters.Add(new SqlParameter("@pNombre", pNombre));
                s.accionoBaseDatos("Se ejecuto la query:" + cmd.CommandText);
                s.accionoBaseDatos("@pNombre paso a valer:" + pNombre);
                SqlDataReader datos = cmd.ExecuteReader();
                lot = crearLote(datos).ElementAt(0);
                s.accionoBaseDatos("Buscar Lote", "OK");
            }
            catch (SqlException e)
            {
                s.accionoBaseDatos("Buscar Lote", "Error: " + e.Message);
            }
            finally
            {
                Coneccion.CerrarConeccion(cn);
            }
            return lot;
        }

        private List<Lote> crearLote(SqlDataReader dr)
        {
            List<Lote> lot = new List<Lote>();
            tolls t = tolls.T;
            Marcador m;
            string nombre;
            DateTime creacion;
            Lote.tipoLote tipoLot;
            Lote.tipoRecurso tipoUni;
            Lote.tipoEstado estadoLot;
            bool eliminado;
            Lote l;
            while (dr.Read())
            {
                m = t.hallarMarcador(dr["nombreMarc"].ToString());
                nombre = dr["nombre"].ToString();
                DateTime.TryParse(dr["creacion"].ToString(), out creacion);
                tipoLot = t.tipoLote(dr["tipoLote"].ToString());
                tipoUni = t.tipoUnidadLote(dr["tipoUnidad"].ToString());
                estadoLot = t.estadoLote(dr["estadoLote"].ToString());
                eliminado = t.elLoteEstaEliminado(estadoLot);
                l = new Lote(nombre, creacion, estadoLot, tipoLot, eliminado);
                l.Marc = m;
                lot.Add(l);
            }
            return lot;
        }

        public List<Marcador> cargarMarcadores()
        {
            tolls t = tolls.T;
            Sistema s = Sistema.Sis;
            List<Marcador> marc = new List<Marcador>();
            s.accionoBaseDatos("Se buscan todos los marcadores");
            SqlConnection cn = Coneccion.CrearConeccionSql
                       (CadenaDeConecciones.tiposDeConeccion.paraDominio);
            SqlCommand cmd = new SqlCommand(Query_s.todoMarcadores, cn);
            Marcador.tipoMarcador tipo;
            try
            {
                Coneccion.AbrirConeccion(cn);
                s.accionoBaseDatos("Se ejecuta la query: " + cmd.CommandText);
                SqlDataReader datos = cmd.ExecuteReader();
                while (datos.Read())
                {
                    tipo = t.tipoMarcador(datos["tipoMarcador"].ToString());
                    marc.Add(new Marcador(datos["nombre"].ToString(), tipo));
                }
                s.accionoBaseDatos("Buscar marcadores", "OK");
            }
            catch (SqlException e)
            {
                s.accionoBaseDatos("Buscar marcadores", "Error: " + e.Message);
            }
            finally
            {
                Coneccion.CerrarConeccion(cn);
            }
            return marc;
        }

        public List<Lote> cargaLotesUnaSemana()
        {
            Sistema s = Sistema.Sis;
            tolls t = tolls.T;
            s.accionoBaseDatos("Se ejecuta cargar lotes una semana");
            List<Lote> lot = new List<Lote>();
            SqlConnection cn = Coneccion.CrearConeccionSql
                             (CadenaDeConecciones.tiposDeConeccion.paraDominio);
            SqlCommand cmd = new SqlCommand(Query_s.cargarLotes1Semana, cn);
            try
            {
                Coneccion.AbrirConeccion(cn);
                cmd.Parameters.Add(new SqlParameter("@fecha", DateTime.Today.AddDays(-8)));
                cmd.Parameters.Add(new SqlParameter("@estado", Lote.tipoEstado.ParaCargar.ToString()));
                s.accionoBaseDatos("Se ejecuta la query: " + cmd.CommandText);
                SqlDataReader datos = cmd.ExecuteReader();
                lot = crearLote(datos);
                s.accionoBaseDatos("Lote una semana", "OK");
            }
            catch (SqlException e)
            {
                s.accionoBaseDatos("Lotes una semana", "Error: " + e.Message);
            }
            finally
            {
                Coneccion.CerrarConeccion(cn);
            }
            return lot;
        }

        public void agregarLote(Lote pLote)
        {
            Sistema s = Sistema.Sis;
            s.accionoBaseDatos("Se ingresa el lote: " + pLote.Nombre + " a la base de datos");
            SqlConnection cn = Coneccion.CrearConeccionSql
                        (CadenaDeConecciones.tiposDeConeccion.paraDominio);
            SqlCommand cmd = new SqlCommand(Query_s.insertFrecuencia, cn);
            Frecuencia f = pLote.Frec;
            SqlTransaction trs = null;
            int id = 86;
            try
            {
                trs = Coneccion.CrearTransaccion(cn);
                cmd.Transaction = trs;
                if (f.BaseContactacion != 100)
                {
                    Coneccion.agregarParametro(cmd, "@frec", f.Id);
                    Coneccion.agregarParametro(cmd, "@base", f.BaseContactacion);
                    Coneccion.agregarParametro(cmd, "@prioridad", f.PrioridadLote);
                    int.TryParse(cmd.ExecuteScalar().ToString(), out id);
                    cmd.Parameters.Clear();
                }
                Coneccion.agregarParametro(cmd, "@nombre", pLote.Nombre);
                Coneccion.agregarParametro(cmd, "@creacion", pLote.Creacion);
                Coneccion.agregarParametro(cmd, "@tipoLote", pLote.LoteTipo.ToString());
                Coneccion.agregarParametro(cmd, "@tipoUnidad", pLote.UnidadNegocio.ToString());
                Coneccion.agregarParametro(cmd, "@frec_id", id);
                Coneccion.agregarParametro(cmd, "@estadoLote", pLote.Estado.ToString());
                Coneccion.agregarParametro(cmd, "@cantLote", pLote.Exc.leerExcel().Count - 1);
                cmd.CommandText = Query_s.insertLote;
                cmd.ExecuteNonQuery();
                trs.Commit();
            }
            catch (Exception e) {
                if (trs != null) trs.Rollback();
                s.accionoBaseDatos("Crear lote : " + pLote.Nombre, "Error: " + e.Message);
            }
            finally
            {
                Coneccion.CerrarConeccion(cn);
            }
        }

    }
}
