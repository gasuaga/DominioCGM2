using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Npgsql;
using System.Data.SqlClient;
namespace BaseDeDatos
{
    /**
     * @class   Coneccion
     *
     * @brief   Clase encargada de la coneccion
     *
     * @author  WINMACROS
     * @date    17/07/2017
     */

    public static class Coneccion
    {
        public static void AbrirConeccion (SqlConnection pCn)
        {
            if (pCn.State == System.Data.ConnectionState.Closed)
                pCn.Open();
        }
        public static void AbrirConeccion(NpgsqlConnection pCn)
        {
            if (pCn.State == System.Data.ConnectionState.Closed)
                pCn.Open();
        }
        public static void CerrarConeccion(SqlConnection pCn)
        {
            if (pCn.State != System.Data.ConnectionState.Closed)
            {
                pCn.Close();
                pCn.Dispose();
            }                
        }
        public static void CerrarConeccion(NpgsqlConnection pCn)
        {
            if (pCn.State != System.Data.ConnectionState.Closed)
            {
                pCn.Close();
                pCn.Dispose();
            }
        }
        public static SqlConnection CrearConeccionSql(CadenaDeConecciones.tiposDeConeccion pTipo)
        {
            switch (pTipo)
            {
                case CadenaDeConecciones.tiposDeConeccion.paraDominio:
                    return new SqlConnection(CadenaDeConecciones.CadenaConeccionDominio);
                case CadenaDeConecciones.tiposDeConeccion.crearScore:
                    return new SqlConnection(CadenaDeConecciones.CadenaConeccionScore);
                case CadenaDeConecciones.tiposDeConeccion.poblarScore:
                    return new SqlConnection(CadenaDeConecciones.CadenaConeccionDatosScore);
            }
            return null;
        }
        public static NpgsqlConnection CrearConeccionPostgre()
        {
            return new NpgsqlConnection(CadenaDeConecciones.CadenaConeccionDatosPosgre);
        }
        public static SqlTransaction CrearTransaccion(SqlConnection pCn)
        {
            if (pCn.State != System.Data.ConnectionState.Open)
                AbrirConeccion(pCn);
            return pCn.BeginTransaction();
        }
        public static void agregarParametro (SqlCommand pCmd, string pNomParametro, object pValor)
        {
            pCmd.Parameters.AddWithValue(pNomParametro, pValor);
        }

    }
}
