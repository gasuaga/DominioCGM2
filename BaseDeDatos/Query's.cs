using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaseDeDatos
{
    /**
     * @class   Query_s
     *
     * @brief   Es donde se guardan todas las querys utilizadas.
     *
     * @author  WINMACROS
     * @date    17/07/2017
     */

    public static class Query_s
    {
        public static string buscarLote = @"select l.*
                           , M.nombre as nombreMarc
                           , M.estado as estadoMarc
                          from Lotes l
                            inner join  lotesMarcador lM on lM.loteNombre = l.nombre
                            inner join Marcadores M on m.nombre = lM.marcadorNombre
                          WHERE 
                            l.nombre = @pNombre";
        public static string todoMarcadores = @"select * from Marcadores";
        public static string cargarLotes1Semana = @"select l.*
                           , M.nombre as nombreMarc
                           , M.estado as estadoMarc
                          from Lotes l
                            inner join  lotesMarcador lM
                            inner join Marcadores M 
                          WHERE 
                            creacion > @fecha and estado != @estado ";
        public static string insertFrecuencia = @"INSERT INTO Frecuencias 
                                (frec, baseContactacion, prioridadLote) 
                                VALUES(@frec, @base, @prioridad); 
                                select CAST(scope_identity() AS INT); ";
        public static string insertLote = @"INSERT INTO Lotes 
                                VALUES (@nombre, @creacion,  @tipoLote, 
                                @tipoUnidad, @frec_id, @estadoLote, @cantLote) ";

    }
}
