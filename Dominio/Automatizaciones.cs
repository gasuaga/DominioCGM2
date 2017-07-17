using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dominio
{
    /**
    * @class   Automatizaciones
    *
    * @brief   Son funciones que se utilizan para crear
    *          coasas (los activos, ivr, sms).
    *
    * @author  WINMACROS
    * @date    28/06/2017
    */

    public class Automatizaciones
    {
        tolls t = tolls.T;
        private static Automatizaciones a;

        public static Automatizaciones A
        {
            get
            {
                if (a == null)
                    a = new Automatizaciones();
                return a;
            }
        }

        private Automatizaciones() { }

        /**
         * @fn  public List<URL> ejecutaraMacroiVR(string pMotor)
         *
         * @brief   Ejecuta la macro para los ivr segun el motor.
         *
         * @author  WINMACROS
         * @date    28/06/2017
         *
         * @param   pMotor  Equipo de negocio.
         *
         * @return  Lista de urls que son los excel's creados;
         */

        public List<URL> ejecutaraMacroiVR(string pMotor)
        {
            Sistema s = Sistema.Sis;
            List<URL> ret = new List<URL>();
            switch (pMotor)
            {
                case "POWER":
                    ret = archivosPowerMvp();
                    break;
                case "RANGER":
                    ret = archivosRangerMvp();
                    break;
                case "MERCURIUS":
                    ret = archivosMercuriusMvp();
                    break;
            }
            tolls.T.mandarMail(pMotor);
            return ret;
        }
        

        private List<URL> archivosMercuriusMvp()
        {
            List<URL> ur = new List<URL>();
            t.EjecutarMacro("mvpMPromotora", 0);
            ur.Add(t.copiarArchivo("mercuriusPromot", "ivr"));

            t.EjecutarMacro("mvpMValorII", 0);
            ur.Add(t.copiarArchivo("mercuriusValorII", "ivr"));

            t.EjecutarMacro("mvpMAsi", 0);
            ur.Add(t.copiarArchivo("mercuriusAsi", "ivr"));

            t.EjecutarMacro("mvpMFinan", 0);
            ur.Add(t.copiarArchivo("mercuriusFinan", "ivr"));

            t.EjecutarMacro("mvpMSantIII", 0);
            ur.Add(t.copiarArchivo("mercuriusSantIII", "ivr"));

            t.EjecutarMacro("mvpMCiclo", 0);
            ur.Add(t.copiarArchivo("mercuriusCicloc", "ivr"));

            t.EjecutarMacro("mvpMProV", 0);
            ur.Add(t.copiarArchivo("mercuriusPromoV", "ivr"));

            t.EjecutarMacro("mvpMBbva", 0);
            ur.Add(t.copiarArchivo("mercuriusBbva", "ivr"));

            t.EjecutarMacro("mvpMCicloIII", 0);
            ur.Add(t.copiarArchivo("mercuriusCiclocIII", "ivr"));

            t.EjecutarMacro("mvpMCicloII", 0);
            ur.Add(t.copiarArchivo("mercuriusCiclocII", "ivr"));

            t.EjecutarMacro("mvpMCredII", 0);
            ur.Add(t.copiarArchivo("mercuriusCredII", "ivr"));

            t.EjecutarMacro("mvpMCicloIV", 0);
            ur.Add(t.copiarArchivo("mercuriusCiclocIV", "ivr"));

            t.EjecutarMacro("mvpMValor", 0);
            ur.Add(t.copiarArchivo("mercuriusValor", "ivr"));

            t.EjecutarMacro("mvpMItalII", 0);
            ur.Add(t.copiarArchivo("mercuriusItalcII", "ivr"));

            t.EjecutarMacro("mvpMAsaC", 0);
            ur.Add(t.copiarArchivo("mercuriusAstexCit", "ivr"));

            t.EjecutarMacro("mvpMMicroIII", 0);
            ur.Add(t.copiarArchivo("mercuriusMicoIII", "ivr"));

            t.EjecutarMacro("mvpMMicro", 0);
            ur.Add(t.copiarArchivo("mercuriusMico", "ivr"));

            t.EjecutarMacro("mvpMCitib", 0);
            ur.Add(t.copiarArchivo("mercuriusCiti", "ivr"));

            t.EjecutarMacro("mvpMPromII", 0);
            ur.Add(t.copiarArchivo("mercuriusPromoII", "ivr"));

            t.EjecutarMacro("mvpMCredip", 0);
            ur.Add(t.copiarArchivo("mercuriusCreditplus", "ivr"));

            t.EjecutarMacro("mvpMMicII", 0);
            ur.Add(t.copiarArchivo("mercuriusMicoII", "ivr"));

            t.EjecutarMacro("mvpMItalV", 0);
            ur.Add(t.copiarArchivo("mercuriusItalcV", "ivr"));

            t.EjecutarMacro("mvpMItalc", 0);
            ur.Add(t.copiarArchivo("mercuriusItalc", "ivr"));

            t.EjecutarMacro("mvpMBbvIII", 0);
            ur.Add(t.copiarArchivo("mercuriusBbvaIII", "ivr"));

            t.EjecutarMacro("mvpMFinanIII", 0);
            ur.Add(t.copiarArchivo("mercuriusFinanIII", "ivr"));

            t.EjecutarMacro("mvpMMila", 0);
            ur.Add(t.copiarArchivo("mercuriusMila", "ivr"));

            t.EjecutarMacro("mvpMSantII", 0);
            ur.Add(t.copiarArchivo("mercuriusSantII", "ivr"));

            t.EjecutarMacro("mvpMFinanII", 0);
            ur.Add(t.copiarArchivo("mercuriusFinanII", "ivr"));

            t.EjecutarMacro("mvpMAsiII", 0);
            ur.Add(t.copiarArchivo("mercuriusAsiII", "ivr"));

            t.EjecutarMacro("mvpMCred", 0);
            ur.Add(t.copiarArchivo("mercuriusCreditel", "ivr"));

            t.EjecutarMacro("mvpMAsiIII", 0);
            ur.Add(t.copiarArchivo("mercuriusAsiIII", "ivr"));

            t.EjecutarMacro("mvpMPront", 0);
            ur.Add(t.copiarArchivo("mercuriusPronto", "ivr"));

            t.EjecutarMacro("mvpMItalIV", 0);
            ur.Add(t.copiarArchivo("mercuriusItalcIV", "ivr"));

            t.EjecutarMacro("mvpMItalIII", 0);
            ur.Add(t.copiarArchivo("mercuriusItalcIII", "ivr"));

            t.EjecutarMacro("mvpMSant", 0);
            ur.Add(t.copiarArchivo("mercuriusSant", "ivr"));

            t.EjecutarMacro("mvpMBbvII", 0);
            ur.Add(t.copiarArchivo("mercuriusBbvaII", "ivr"));

            t.EjecutarMacro("mvpMValoIII", 0);
            ur.Add(t.copiarArchivo("mercuriusValorIII", "ivr"));

            t.EjecutarMacro("mvpMPromoIII", 0);
            ur.Add(t.copiarArchivo("mercuriusPromoIII", "ivr"));

            t.EjecutarMacro("mvpMCash", 0);
            ur.Add(t.copiarArchivo("mercuriusCash", "ivr"));

            t.EjecutarMacro("mvpMClaro", 0);
            ur.Add(t.copiarArchivo("mercuriusClaro", "ivr"));

            t.EjecutarMacro("mvpMPronII", 0);
            ur.Add(t.copiarArchivo("mercuriusProntoII", "ivr"));

            List<Lote> list = new List<Lote>();
            List<URL> ret = new List<URL>();
            int frec = 1;
            foreach (URL u in ur)
            {
                if (u.Nombre.Contains("mercuriusCred"))
                    frec = 2;
                try
                {
                    if (u != null)
                    {
                        Excel e = new Excel { Direccion = u };
                        if (e.validarExcel())
                        {
                            Sistema s = Sistema.Sis;
                            Lote l = new Lote(u.Nombre, DateTime.Today, Lote.tipoEstado.ParaCargar, t.tipoLoteXCampana(u), frec, e, true);
                            if (s.agregarLote(l) == Sistema.resultAgregarLote.agrego)
                            {
                                list.Add(l);
                                ret.Add(u);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }

            if (ret.Count <= (ur.Count / 1.70))
                ret = archivosMercuriusMvp();
            return ret;
        }

        private List<URL> archivosRangerMvp()
        {
            List<URL> ur = new List<URL>();
            t.EjecutarMacro("mvpAnda", 0);
            ur.Add(t.copiarArchivo("Anda", "ivr"));

            t.EjecutarMacro("mvpCash", 0);
            ur.Add(t.copiarArchivo("Cash", "ivr"));

            t.EjecutarMacro("mvpBarracaNort", 0);
            ur.Add(t.copiarArchivo("BarracNot", "ivr"));

            t.EjecutarMacro("mvpBBVA", 0);
            ur.Add(t.copiarArchivo("bbva", "ivr"));

            t.EjecutarMacro("mvpBCBS", 0);
            ur.Add(t.copiarArchivo("blueCros", "ivr"));

            t.EjecutarMacro("mvpCabal", 0);
            ur.Add(t.copiarArchivo("Cabal", "ivr"));

            t.EjecutarMacro("mvpCacson", 0);
            ur.Add(t.copiarArchivo("Cacson", "ivr"));

            t.EjecutarMacro("mvpCanal10", 0);
            ur.Add(t.copiarArchivo("Canal10", "ivr"));

            t.EjecutarMacro("mvpCarUp", 0);
            ur.Add(t.copiarArchivo("CarUp", "ivr"));

            t.EjecutarMacro("mvpCasMan", 0);
            ur.Add(t.copiarArchivo("CasMan", "ivr"));

            t.EjecutarMacro("mvpCay", 0);
            ur.Add(t.copiarArchivo("CAYC", "ivr"));

            t.EjecutarMacro("mvpCgmAgr", 0);
            ur.Add(t.copiarArchivo("CgmAgr", "ivr"));

            t.EjecutarMacro("mvpCGSA", 0);
            ur.Add(t.copiarArchivo("Cgsa", "ivr"));

            t.EjecutarMacro("mvpCYVV", 0);
            ur.Add(t.copiarArchivo("CYVV", "ivr"));

            t.EjecutarMacro("mvpCint", 0);
            ur.Add(t.copiarArchivo("Cint", "ivr"));

            t.EjecutarMacro("mvpCodac", 0);
            ur.Add(t.copiarArchivo("Codac", "ivr"));

            t.EjecutarMacro("mvpComsa", 0);
            ur.Add(t.copiarArchivo("Comsa", "ivr"));

            t.EjecutarMacro("mvpComayc", 0);
            ur.Add(t.copiarArchivo("Comayc", "ivr"));

            t.EjecutarMacro("mvpComef", 0);
            ur.Add(t.copiarArchivo("Comef", "ivr"));

            t.EjecutarMacro("mvpCoopace", 0);
            ur.Add(t.copiarArchivo("Coopace", "ivr"));

            t.EjecutarMacro("mvpCoopBanc", 0);
            ur.Add(t.copiarArchivo("CoopBanc", "ivr"));

            t.EjecutarMacro("mvpCopac", 0);
            ur.Add(t.copiarArchivo("Copac", "ivr"));

            t.EjecutarMacro("mvpCopagran", 0);
            ur.Add(t.copiarArchivo("Copagran", "ivr"));

            t.EjecutarMacro("mvpCredic", 0);
            ur.Add(t.copiarArchivo("Credic", "ivr"));

            t.EjecutarMacro("mvpCredif", 0);
            ur.Add(t.copiarArchivo("Credif", "ivr"));

            t.EjecutarMacro("mvpCredipl", 0);
            ur.Add(t.copiarArchivo("Credipl", "ivr"));

            t.EjecutarMacro("mvpCredipunta", 0);
            ur.Add(t.copiarArchivo("Credipu", "ivr"));

            t.EjecutarMacro("mvpCredirap", 0);
            ur.Add(t.copiarArchivo("Credirap", "ivr"));

            t.EjecutarMacro("mvpCredisur", 0);
            ur.Add(t.copiarArchivo("Credisur", "ivr"));

            t.EjecutarMacro("mvpCreditoYa", 0);
            ur.Add(t.copiarArchivo("CrYa", "ivr"));

            t.EjecutarMacro("mvpCredClubEste", 0);
            ur.Add(t.copiarArchivo("CredClub", "ivr"));

            t.EjecutarMacro("mvpCreditosDirectos", 0);
            ur.Add(t.copiarArchivo("CredDir", "ivr"));

            t.EjecutarMacro("mvpCredS", 0);
            ur.Add(t.copiarArchivo("CredS", "ivr"));

            t.EjecutarMacro("mvpDiLuss", 0);
            ur.Add(t.copiarArchivo("DiLuss", "ivr"));

            t.EjecutarMacro("mvpEcoc", 0);
            ur.Add(t.copiarArchivo("Ecoc", "ivr"));

            t.EjecutarMacro("mvpElDorado", 0);
            ur.Add(t.copiarArchivo("Dorado", "ivr"));

            t.EjecutarMacro("mvpElPais", 0);
            ur.Add(t.copiarArchivo("ElPais", "ivr"));

            t.EjecutarMacro("mvpValor", 0);
            ur.Add(t.copiarArchivo("Valor", "ivr"));

            t.EjecutarMacro("mvpFarSegu", 0);
            ur.Add(t.copiarArchivo("FarSeg", "ivr"));

            /*  t.EjecutarMacro("mvpFastCred", 0);
              ur.Add(t.copiarArchivo("FasCre", "ivr"));*/

            t.EjecutarMacro("mvpFUCAC", 0);
            ur.Add(t.copiarArchivo("FUCAC", "ivr"));

            t.EjecutarMacro("mvpFUCAC2", 0);
            ur.Add(t.copiarArchivo("FUCACM", "ivr"));

            t.EjecutarMacro("mvpFUCAC3", 0);
            ur.Add(t.copiarArchivo("FUCACP", "ivr"));

            t.EjecutarMacro("mvpFucerep", 0);
            ur.Add(t.copiarArchivo("Fucerep", "ivr"));

            t.EjecutarMacro("mvpFundasol", 0);
            ur.Add(t.copiarArchivo("Fundsol", "ivr"));

            t.EjecutarMacro("mvpGrupoGama", 0);
            ur.Add(t.copiarArchivo("GrupGama", "ivr"));

            t.EjecutarMacro("mvpGuil", 0);
            ur.Add(t.copiarArchivo("Guil", "ivr"));

            t.EjecutarMacro("mvpHdc", 0);
            ur.Add(t.copiarArchivo("Hdc", "ivr"));

            t.EjecutarMacro("mvpLigaSan", 0);
            ur.Add(t.copiarArchivo("LigaSan", "ivr"));

            t.EjecutarMacro("mvpMarc", 0);
            ur.Add(t.copiarArchivo("Marc", "ivr"));

            t.EjecutarMacro("mvpMilen", 0);
            ur.Add(t.copiarArchivo("Milen", "ivr"));

            t.EjecutarMacro("mvpMonte", 0);
            ur.Add(t.copiarArchivo("Monte", "ivr"));

            t.EjecutarMacro("mvpMontE", 0);
            ur.Add(t.copiarArchivo("MonteE", "ivr"));

            t.EjecutarMacro("mvpMontRef", 0);
            ur.Add(t.copiarArchivo("MontR", "ivr"));

            t.EjecutarMacro("mvpNelsonSobr", 0);
            ur.Add(t.copiarArchivo("NelSobr", "ivr"));

            t.EjecutarMacro("mvpNestorCa", 0);
            ur.Add(t.copiarArchivo("NesCa", "ivr"));

            t.EjecutarMacro("mvpNuevoSig", 0);
            ur.Add(t.copiarArchivo("NueSig", "ivr"));

            t.EjecutarMacro("mvpNuevo", 0);
            ur.Add(t.copiarArchivo("Nuevo", "ivr"));

            t.EjecutarMacro("mvpPass", 0);
            ur.Add(t.copiarArchivo("Pass", "ivr"));

            t.EjecutarMacro("mvpPrestacel", 0);
            ur.Add(t.copiarArchivo("Prestacel", "ivr"));

            t.EjecutarMacro("mvpProsegur", 0);
            ur.Add(t.copiarArchivo("Prosegur", "ivr"));

            t.EjecutarMacro("mvpRapid", 0);
            ur.Add(t.copiarArchivo("Rapid", "ivr"));

            t.EjecutarMacro("mvpRepMic", 0);
            ur.Add(t.copiarArchivo("RepMic", "ivr"));

            t.EjecutarMacro("mvpSaint", 0);
            ur.Add(t.copiarArchivo("Saint", "ivr"));

            t.EjecutarMacro("mvpVolt", 0);
            ur.Add(t.copiarArchivo("Volt", "ivr"));

            t.EjecutarMacro("mvpWur", 0);
            ur.Add(t.copiarArchivo("Wur", "ivr"));

            t.EjecutarMacro("mvpYTa", 0);
            ur.Add(t.copiarArchivo("YTa", "ivr"));

            List<Lote> list = new List<Lote>();
            List<URL> ret = new List<URL>();
            foreach (URL u in ur)
            {
                try
                {
                    if (u != null)
                    {
                        Excel e = new Excel { Direccion = u };
                        if (e.validarExcel())
                        {
                            Sistema s = Sistema.Sis;
                            Lote l = new Lote(u.Nombre, DateTime.Today, Lote.tipoEstado.ParaCargar, t.tipoLoteXCampana(u), 1, e, true);
                            if (s.agregarLote(l) == Sistema.resultAgregarLote.agrego)
                            {
                                list.Add(l);
                                ret.Add(u);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }

            return ret;
        }

        private List<URL> archivosPowerMvp()
        {
            List<URL> ur = new List<URL>();
            t.EjecutarMacro("mvpCdcEspecial", 0);
            ur.Add(t.copiarArchivo("cdcEspecial", "ivr"));

            t.EjecutarMacro("mvpRetopEspecual", 0);
            ur.Add(t.copiarArchivo("retopEspecial", "ivr"));

            t.EjecutarMacro("mvpPronto", 0);
            ur.Add(t.copiarArchivo("pronto", "ivr"));

            t.EjecutarMacro("mvpCreciditaND", 0);
            ur.Add(t.copiarArchivo("CreditiaND", "ivr"));

            t.EjecutarMacro("mvpCreciditaP", 0);
            ur.Add(t.copiarArchivo("CreditiaP", "ivr"));

            t.EjecutarMacro("mvpCreciditaS", 0);
            ur.Add(t.copiarArchivo("CreditiaS", "ivr"));

            t.EjecutarMacro("mvpCreciditaB", 0);
            ur.Add(t.copiarArchivo("CreditiaB", "ivr"));

            t.EjecutarMacro("mvpCreciditaT", 0);
            ur.Add(t.copiarArchivo("CreditiaT", "ivr"));

            t.EjecutarMacro("mvpCreciditaNP", 0);
            ur.Add(t.copiarArchivo("CreditiaNP", "ivr"));

            t.EjecutarMacro("mvpRetop", 0);
            ur.Add(t.copiarArchivo("Retop", "ivr"));

            t.EjecutarMacro("mvpCreciditaP2", 0);
            ur.Add(t.copiarArchivo("CreditiaP2", "ivr"));
            List<Lote> list = new List<Lote>();
            List<URL> ret = new List<URL>();
            foreach (URL u in ur)
            {
                try
                {
                    if (u != null)
                    {
                        Excel e = new Excel { Direccion = u };
                        if (e.validarExcel())
                        {
                            Sistema s = Sistema.Sis;
                            Lote l = new Lote(u.Nombre, DateTime.Today, Lote.tipoEstado.ParaCargar, t.tipoLoteXCampana(u), 1, e, true);
                            if (s.agregarLote(l) == Sistema.resultAgregarLote.agrego)
                            {
                                l.cargarScore();
                                list.Add(l);
                                ret.Add(u);
                            }
                        }
                    }
                }
                catch (Exception) { }
            }
            if (ret.Count <= (ur.Count / 1.70))
                ret = archivosPowerMvp();
            return ret;
        }

        private List<URL> archivosMercurius()
        {
            List<URL> ur = new List<URL>();
            t.EjecutarMacro("mvpMercuriusF1", 0);
            ur.Add(t.copiarArchivo("mercuriusf1", "ivr"));
            t.EjecutarMacro("mvpMercuriusF2", 0);
            ur.Add(t.copiarArchivo("mercuriusf2", "ivr"));
            t.EjecutarMacro("mvpMercuriusCc", 0);
            ur.Add(t.copiarArchivo("mercuriuscc", "ivr"));
            t.EjecutarMacro("mvpMercuriusCreditel", 0);
            ur.Add(t.copiarArchivo("mercuriuscreditel", "ivr"));
            return ur;
        }

        public List<URL> archivosMarcador(string pEquipo)
        {
            List<URL> ur = new List<URL>();
            switch (pEquipo)
            {
                case "PRONTO":
                    t.EjecutarMacro("mdProntoACont", 0);
                    ur.Add(t.copiarArchivo("prontoACont", "md"));
                    t.EjecutarMacro("mdProntoCont", 0);
                    ur.Add(t.copiarArchivo("prontoCont", "md"));
                    break;
                case "WHITE":
                    t.EjecutarMacro("mdWhiteACont", 0);
                    ur.Add(t.copiarArchivo("whiteACont", "md"));
                    t.EjecutarMacro("mdWhiteCont", 0);
                    ur.Add(t.copiarArchivo("whiteCont", "md"));
                    break;
                case "NEGRO":
                    t.EjecutarMacro("mdNegroACont", 0);
                    ur.Add(t.copiarArchivo("negroACont", "md"));
                    t.EjecutarMacro("mdNegroCont", 0);
                    ur.Add(t.copiarArchivo("negroCont", "md"));
                    break;
                case "ORANGE":
                    t.EjecutarMacro("mdOrangeACont", 0);
                    ur.Add(t.copiarArchivo("orangeACont", "md"));
                    t.EjecutarMacro("mdOrangeCont", 0);
                    ur.Add(t.copiarArchivo("orangeCont", "md"));
                    break;
                case "AMBAR":
                    t.EjecutarMacro("mdAmbarACont", 0);
                    ur.Add(t.copiarArchivo("ambarACont", "md"));
                    t.EjecutarMacro("mdAmbarCont", 0);
                    ur.Add(t.copiarArchivo("ambarCont", "md"));
                    break;
                case "OCA":
                    t.EjecutarMacro("mdOcaACont", 0);
                    ur.Add(t.copiarArchivo("ocaACont", "md"));
                    t.EjecutarMacro("mdOcaCont", 0);
                    ur.Add(t.copiarArchivo("ocaCont", "md"));
                    break;
            }
            return ur;
        }

        public void activosSms()
        {
            baseDatos bd = baseDatos.Bd;
            List<string[]> lote = bd.smsDelBigFish();
            Excel.crearExcel(lote);
            tolls.T.EjecutarMacro(0);
        }


    }
}
