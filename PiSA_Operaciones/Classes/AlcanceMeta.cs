using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PiSA_Operaciones.Classes
{
    internal class AlcanceMeta
    {
        internal static List<AlcanceMetaRecord> Leer22N(string _ruta)
        {
            var alcanceMetaList = new List<AlcanceMetaRecord>();
            var listaEntregas = new List<string>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    string entrega = fila.Cell(28).GetString();
                    bool entregaUnica = (listaEntregas.IndexOf(entrega) == -1 && entrega != string.Empty) ? true : false;

                    if (listaEntregas.IndexOf(entrega) == -1 && entrega != string.Empty) listaEntregas.Add(entrega);

                    var alcance = new AlcanceMetaRecord
                    {
                        A = fila.Cell(1).GetString(),
                        B = fila.Cell(2).IsEmpty() ? 0 : fila.Cell(2).GetValue<int>(),
                        C = fila.Cell(3).IsEmpty() ? DateTime.MinValue : fila.Cell(3).GetDateTime(),
                        D = fila.Cell(4).GetString(),
                        E = fila.Cell(5).GetString(),
                        F = fila.Cell(6).GetString(),
                        G = fila.Cell(7).GetString(),
                        H = fila.Cell(8).GetString(),
                        I = fila.Cell(9).GetString(),
                        J = fila.Cell(10).GetString(),
                        K = fila.Cell(11).GetString(),
                        L = fila.Cell(12).GetString(),
                        M = fila.Cell(13).GetString(),
                        N = fila.Cell(14).IsEmpty() ? DateTime.MinValue : fila.Cell(14).GetDateTime(),
                        O = fila.Cell(15).GetString(),
                        P = fila.Cell(16).GetString(),
                        Q = fila.Cell(17).GetString(),
                        R = fila.Cell(18).GetString(),
                        S = fila.Cell(19).IsEmpty() ? 0 : fila.Cell(19).GetValue<int>(),
                        T = fila.Cell(20).GetString(),
                        U = fila.Cell(21).IsEmpty() ? 0 : fila.Cell(21).GetValue<int>(),
                        V = fila.Cell(22).IsEmpty() ? 0 : fila.Cell(22).GetValue<int>(),
                        W = fila.Cell(23).GetString(),
                        X = fila.Cell(24).IsEmpty() ? 0 : fila.Cell(24).GetValue<float>(),
                        Y = fila.Cell(25).IsEmpty() ? 0 : fila.Cell(25).GetValue<float>(),
                        Z = fila.Cell(26).GetString(),
                        AA = fila.Cell(27).GetString(),
                        AB = entrega,
                        AC = fila.Cell(29).IsEmpty() ? 0 : fila.Cell(29).GetValue<int>(),
                        AD = fila.Cell(30).GetString(),
                        AE = fila.Cell(31).IsEmpty() ? DateTime.MinValue : fila.Cell(31).GetDateTime(),
                        AF = fila.Cell(32).IsEmpty() ? 0 : fila.Cell(32).GetValue<int>(),
                        AG = fila.Cell(33).IsEmpty() ? DateTime.MinValue : fila.Cell(33).GetDateTime(),
                        AH = fila.Cell(34).GetString(),
                        AI = fila.Cell(35).GetString(),
                        AJ = fila.Cell(36).IsEmpty() ? DateTime.MinValue : fila.Cell(36).GetDateTime(),
                        AK = fila.Cell(37).GetString(),
                        AL = fila.Cell(38).IsEmpty() ? DateTime.MinValue : fila.Cell(38).GetDateTime(),
                        AM = fila.Cell(39).IsEmpty() ? DateTime.MinValue : fila.Cell(39).GetDateTime(),
                        AN = fila.Cell(40).IsEmpty() ? DateTime.MinValue : fila.Cell(40).GetDateTime(),
                        AO = fila.Cell(41).GetString(),
                        AP = fila.Cell(42).GetString(),
                        AQ = fila.Cell(43).GetString(),
                        AR = fila.Cell(44).IsEmpty() ? DateTime.MinValue : fila.Cell(44).GetDateTime(),
                        AS = fila.Cell(45).IsEmpty() ? DateTime.MinValue : fila.Cell(45).GetDateTime(),
                        AT = fila.Cell(46).GetString(),
                        AU = fila.Cell(47).GetString(),
                        AV = fila.Cell(48).GetString(),
                        AW = fila.Cell(49).GetString(),
                        AX = fila.Cell(50).GetString(),
                        AY = fila.Cell(51).GetString(),
                        AZ = fila.Cell(52).IsEmpty() ? 0 : fila.Cell(52).GetValue<int>(),
                        BA = fila.Cell(53).GetString(),
                        BB = fila.Cell(54).GetString(),
                        BC = entregaUnica ? 1 : 0,
                        BD = fila.Cell(56).IsEmpty() ? 0 : fila.Cell(56).GetValue<float>(),
                        BE = fila.Cell(57).IsEmpty() ? 0 : fila.Cell(57).GetValue<int>(),
                        BF = fila.Cell(58).GetString(),
                        BG = fila.Cell(59).GetString(),
                        BH = fila.Cell(60).IsEmpty() ? 0 : fila.Cell(60).GetValue<float>(),
                        BI = fila.Cell(61).IsEmpty() ? 0 : fila.Cell(61).GetValue<int>(),
                        BJ = fila.Cell(62).IsEmpty() ? DateTime.MinValue : fila.Cell(62).GetDateTime(),
                        BK = fila.Cell(63).GetString(),
                        BL = fila.Cell(64).GetString(),
                        BM = fila.Cell(65).GetString(),
                        BN = fila.Cell(66).GetString(),
                        BO = fila.Cell(67).GetString(),
                        BP = fila.Cell(68).GetString(),
                    };
                    alcance.AZ = alcance.S - alcance.AF;
                    alcance.BA = alcance.D == "ZKE" ? "Consigna" : "Pedido";
                    alcance.BD = alcance.Y * alcance.AF;
                    alcance.BE = alcance.BD > 0 ? 1 : 0;
                    alcance.BH = alcance.AZ * alcance.Y;
                    alcance.BK = alcance.AH == "" ? "No Facturado" : "Facturado";

                    alcanceMetaList.Add(alcance);
                }
            }

            return alcanceMetaList;
        }

        internal static List<AlcanceMetaZLO10> LeerZLO10(string _ruta)
        {
            var zlo10List = new List<AlcanceMetaZLO10>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var zlo10 = new AlcanceMetaZLO10
                    {
                        Entrega = fila.Cell(1).GetString(),
                        DT = fila.Cell(2).GetString(),
                    };
                    zlo10List.Add(zlo10);
                }
            }
            zlo10List.Reverse();

            return zlo10List;
        }

        internal static List<AlcanceMetaVLO6F> LeerVL06F(string _ruta)
        {
            var vl06fList = new List<AlcanceMetaVLO6F>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var vl06f = new AlcanceMetaVLO6F
                    {
                        Entrega = fila.Cell(1).GetString(),
                        LugarDestino = fila.Cell(5).GetString(),
                        StatusMovimiento = fila.Cell(8).GetString(),
                    };
                    vl06fList.Add(vl06f);
                }
            }

            return vl06fList;
        }

        internal static List<AlcanceMetaZSD137> LeerZSD137(string _ruta)
        {
            var zsd137List = new List<AlcanceMetaZSD137>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var zsd137 = new AlcanceMetaZSD137
                    {
                        DocumentoComercial = fila.Cell(1).GetString(),
                        Fecha = fila.Cell(49).IsEmpty() ? DateTime.MinValue : fila.Cell(49).GetDateTime(),
                    };
                    var res = zsd137List.Find(x => x.DocumentoComercial == zsd137.DocumentoComercial);
                    if (res == null && zsd137.Fecha != DateTime.MinValue) zsd137List.Add(zsd137);
                }
            }

            return zsd137List;
        }

        internal static void CrearExcel(List<AlcanceMetaRecord> _listaAlcanceMeta,int _noMes)
        {
            using (var libro = new XLWorkbook())
            {
                var hoy = DateTime.Now;
                string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string rutaCompleta = Path.Combine(rutaEscritorio, $"Alcance a la meta - {(Meses)_noMes} {hoy.Day}_{hoy.Hour}{hoy.Minute}.xlsx");

                // Hojas de Excel
                var hojaDatos = libro.Worksheets.Add("Datos");
                var hojaBloqueo = libro.Worksheets.Add("BLOQUEO");
                var hoja = libro.Worksheets.Add("Alcance a la meta");

                // Hoja Datos
                hojaDatos.Cell(1, 1).Value = "CANAL";
                hojaDatos.Cell(1, 2).Value = "ORG";
                hojaDatos.Cell(1, 3).Value = "PRIORIDAD";
                hojaDatos.Cell(2, 1).Value = "25";
                hojaDatos.Cell(2, 2).Value = "POTC";
                hojaDatos.Cell(2, 3).Value = "OTC";
                hojaDatos.Cell(3, 1).Value = "26";
                hojaDatos.Cell(3, 2).Value = "POTC";
                hojaDatos.Cell(3, 3).Value = "OTC";
                hojaDatos.Cell(4, 1).Value = "27";
                hojaDatos.Cell(4, 2).Value = "POTC";
                hojaDatos.Cell(4, 3).Value = "OTC";
                hojaDatos.Cell(5, 1).Value = "28";
                hojaDatos.Cell(5, 2).Value = "POTC";
                hojaDatos.Cell(5, 3).Value = "OTC";
                hojaDatos.Cell(6, 1).Value = "25";
                hojaDatos.Cell(6, 2).Value = "FRIS";
                hojaDatos.Cell(6, 3).Value = "FRISO";
                hojaDatos.Cell(7, 1).Value = "26";
                hojaDatos.Cell(7, 2).Value = "FRIS";
                hojaDatos.Cell(7, 3).Value = "FRISO";
                hojaDatos.Cell(8, 1).Value = "27";
                hojaDatos.Cell(8, 2).Value = "FRIS";
                hojaDatos.Cell(8, 3).Value = "FRISO";
                hojaDatos.Cell(9, 1).Value = "35";
                hojaDatos.Cell(9, 2).Value = "PIPR";
                hojaDatos.Cell(9, 3).Value = "MP";
                hojaDatos.Cell(10, 1).Value = "87";
                hojaDatos.Cell(10, 2).Value = "PIPR";
                hojaDatos.Cell(10, 3).Value = "MP";
                hojaDatos.Cell(11, 1).Value = "25";
                hojaDatos.Cell(11, 2).Value = "PIPR";
                hojaDatos.Cell(11, 3).Value = "VM";
                hojaDatos.Cell(12, 1).Value = "26";
                hojaDatos.Cell(12, 2).Value = "PIPR";
                hojaDatos.Cell(12, 3).Value = "VM";
                hojaDatos.Cell(13, 1).Value = "27";
                hojaDatos.Cell(13, 2).Value = "PIPR";
                hojaDatos.Cell(13, 3).Value = "VM";
                hojaDatos.Cell(14, 1).Value = "28";
                hojaDatos.Cell(14, 2).Value = "PIPR";
                hojaDatos.Cell(14, 3).Value = "VM";

                hojaDatos.Range("A1:C1").Style.Font.Bold = true;
                hojaDatos.Columns().AdjustToContents();

                // Hoja Bloqueo
                hojaBloqueo.Cell(1, 1).Value = "Pedido";
                hojaBloqueo.Cell(1, 2).Value = "Bloqueo de factura";
                hojaBloqueo.Cell(1, 3).Value = "CTE";
                hojaBloqueo.Cell(1, 4).Value = "Comentarios";

                hojaBloqueo.Range("A1:D1").Style.Font.Bold = true;
                hojaBloqueo.Columns().AdjustToContents();

                // Encabezados Excel principal
                hoja.Cell(1, 1).Value = "Documento comercial";
                hoja.Cell(1, 2).Value = "Posición (SD)";
                hoja.Cell(1, 3).Value = "Fecha de pedido";
                hoja.Cell(1, 4).Value = "Clase doc.ventas";
                hoja.Cell(1, 5).Value = "Cliente";
                hoja.Cell(1, 6).Value = "Solicitante";
                hoja.Cell(1, 7).Value = "Destino";
                hoja.Cell(1, 8).Value = "Destinatario de Mercancias";
                hoja.Cell(1, 9).Value = "Organización ventas";
                hoja.Cell(1, 10).Value = "Canal distribución";
                hoja.Cell(1, 11).Value = "Oficina de ventas";
                hoja.Cell(1, 12).Value = "Grupo de vendedores";
                hoja.Cell(1, 13).Value = "Zona de ventas";
                hoja.Cell(1, 14).Value = "Fecha pref.entrega";
                hoja.Cell(1, 15).Value = "N° pedido cliente";
                hoja.Cell(1, 16).Value = "Material";
                hoja.Cell(1, 17).Value = "Texto breve de material";
                hoja.Cell(1, 18).Value = "Grupo de artículos";
                hoja.Cell(1, 19).Value = "Cantidad de pedido";
                hoja.Cell(1, 20).Value = "Un.medida venta";
                hoja.Cell(1, 21).Value = "Cantidad-acum-confir";
                hoja.Cell(1, 22).Value = "Cantidad pedida - cantidad confirmada";
                hoja.Cell(1, 23).Value = "Centro";
                hoja.Cell(1, 24).Value = "Valor neto";
                hoja.Cell(1, 25).Value = "Precio neto";
                hoja.Cell(1, 26).Value = "Motivo de rechazo";
                hoja.Cell(1, 27).Value = "Status total crédito";
                hoja.Cell(1, 28).Value = "Entrega";
                hoja.Cell(1, 29).Value = "Posición de entrega";
                hoja.Cell(1, 30).Value = "Tipo de Documento";
                hoja.Cell(1, 31).Value = "Fecha de Entrega";
                hoja.Cell(1, 32).Value = "Cantidad entrega";
                hoja.Cell(1, 33).Value = "Fe.mov.mcía.real";
                hoja.Cell(1, 34).Value = "Factura";
                hoja.Cell(1, 35).Value = "Tipo de documento2";
                hoja.Cell(1, 36).Value = "Fecha de Factura";
                hoja.Cell(1, 37).Value = "Programa para control";
                hoja.Cell(1, 38).Value = "Fe.act.desp.expd.";
                hoja.Cell(1, 39).Value = "Inic.actual transp.";
                hoja.Cell(1, 40).Value = "Inicio en UTC";
                hoja.Cell(1, 41).Value = "Tipo de contratista";
                hoja.Cell(1, 42).Value = "Clase de transporte";
                hoja.Cell(1, 43).Value = "Agente servicios";
                hoja.Cell(1, 44).Value = "Creado el";
                hoja.Cell(1, 45).Value = "Fecha documento";
                hoja.Cell(1, 46).Value = "Tp.doc.subsiguiente";
                hoja.Cell(1, 47).Value = "Tp.doc.subsiguiente3";
                hoja.Cell(1, 48).Value = "Carácter 1";
                hoja.Cell(1, 49).Value = "Motivo pedido";
                hoja.Cell(1, 50).Value = "Denominación";
                hoja.Cell(1, 51).Value = "Pedido Bloqueo";
                hoja.Cell(1, 52).Value = "Piezas negadas";
                hoja.Cell(1, 53).Value = "Tipo Pedido";
                hoja.Cell(1, 54).Value = "Mercados";
                hoja.Cell(1, 55).Value = "Suma de Entrega";
                hoja.Cell(1, 56).Value = "Venta real";
                hoja.Cell(1, 57).Value = "Aplica";
                hoja.Cell(1, 58).Value = "Transporte";
                hoja.Cell(1, 59).Value = "Para facturar";
                hoja.Cell(1, 60).Value = "Venta sin confirmar";
                hoja.Cell(1, 61).Value = "SEMANA";
                hoja.Cell(1, 62).Value = "DATOS B";
                hoja.Cell(1, 63).Value = "Estatus Factura";
                hoja.Cell(1, 64).Value = "MES creación";
                hoja.Cell(1, 65).Value = "Num Pedido";
                hoja.Cell(1, 66).Value = "MES Factura";
                hoja.Cell(1, 67).Value = "MES de Entrega";
                hoja.Cell(1, 68).Value = "Lugar Destino";

                hoja.Range("A1:BP1").Style.Font.Bold = true;

                int filaActual = 2;
                foreach (var row in _listaAlcanceMeta)
                {
                    hoja.Cell(filaActual, 1).Value = row.A;
                    hoja.Cell(filaActual, 2).Value = row.B;
                    //hoja.Cell(filaActual, 3).Value = row.C == DateTime.MinValue ? null : row.C;
                    hoja.Cell(filaActual, 3).Value = row.C;
                    hoja.Cell(filaActual, 4).Value = row.D;
                    hoja.Cell(filaActual, 5).Value = row.E;
                    hoja.Cell(filaActual, 6).Value = row.F;
                    hoja.Cell(filaActual, 7).Value = row.G;
                    hoja.Cell(filaActual, 8).Value = row.H;
                    hoja.Cell(filaActual, 9).Value = row.I;
                    hoja.Cell(filaActual, 10).Value = row.J;
                    hoja.Cell(filaActual, 11).Value = row.K;
                    hoja.Cell(filaActual, 12).Value = row.L;
                    hoja.Cell(filaActual, 13).Value = row.M;
                    //hoja.Cell(filaActual, 14).Value = row.N == DateTime.MinValue ? null : row.N;
                    hoja.Cell(filaActual, 14).Value = row.N;
                    hoja.Cell(filaActual, 15).Value = row.O;
                    hoja.Cell(filaActual, 16).Value = row.P;
                    hoja.Cell(filaActual, 17).Value = row.Q;
                    hoja.Cell(filaActual, 18).Value = row.R;
                    hoja.Cell(filaActual, 19).Value = row.S;
                    hoja.Cell(filaActual, 20).Value = row.T;
                    hoja.Cell(filaActual, 21).Value = row.U;
                    hoja.Cell(filaActual, 22).Value = row.V;
                    hoja.Cell(filaActual, 23).Value = row.W;
                    hoja.Cell(filaActual, 24).Value = row.X;
                    hoja.Cell(filaActual, 25).Value = row.Y;
                    hoja.Cell(filaActual, 26).Value = row.Z;
                    hoja.Cell(filaActual, 27).Value = row.AA;
                    hoja.Cell(filaActual, 28).Value = row.AB;
                    hoja.Cell(filaActual, 29).Value = row.AC;
                    hoja.Cell(filaActual, 30).Value = row.AD;
                    //hoja.Cell(filaActual, 31).Value = row.AE == DateTime.MinValue ? null : row.AE;
                    hoja.Cell(filaActual, 31).Value = row.AE;
                    hoja.Cell(filaActual, 32).Value = row.AF;
                    //hoja.Cell(filaActual, 33).Value = row.AG == DateTime.MinValue ? null : row.AG;
                    hoja.Cell(filaActual, 33).Value = row.AG;
                    hoja.Cell(filaActual, 34).Value = row.AH;
                    hoja.Cell(filaActual, 35).Value = row.AI;
                    //hoja.Cell(filaActual, 36).Value = row.AJ == DateTime.MinValue ? null : row.AJ;
                    hoja.Cell(filaActual, 36).Value = row.AJ;
                    hoja.Cell(filaActual, 37).Value = row.AK;
                    //hoja.Cell(filaActual, 38).Value = row.AL == DateTime.MinValue ? null : row.AL;
                    //hoja.Cell(filaActual, 39).Value = row.AM == DateTime.MinValue ? null : row.AM;
                    //hoja.Cell(filaActual, 40).Value = row.AN == DateTime.MinValue ? null : row.AN;
                    hoja.Cell(filaActual, 38).Value = row.AL;
                    hoja.Cell(filaActual, 39).Value = row.AM;
                    hoja.Cell(filaActual, 40).Value = row.AN;
                    hoja.Cell(filaActual, 41).Value = row.AO;
                    hoja.Cell(filaActual, 42).Value = row.AP;
                    hoja.Cell(filaActual, 43).Value = row.AQ;
                    //hoja.Cell(filaActual, 44).Value = row.AR == DateTime.MinValue ? null : row.AR;
                    //hoja.Cell(filaActual, 45).Value = row.AS == DateTime.MinValue ? null : row.AS;
                    hoja.Cell(filaActual, 44).Value = row.AR;
                    hoja.Cell(filaActual, 45).Value = row.AS;
                    hoja.Cell(filaActual, 46).Value = row.AT;
                    hoja.Cell(filaActual, 47).Value = row.AU;
                    hoja.Cell(filaActual, 48).Value = row.AV;
                    hoja.Cell(filaActual, 49).Value = row.AW;
                    hoja.Cell(filaActual, 50).Value = row.AX;
                    hoja.Cell(filaActual, 51).FormulaA1 = $"=IFERROR(VLOOKUP(A{filaActual},BLOQUEO!A:B,2,0),\"\")"; // AY
                    hoja.Cell(filaActual, 52).Value = row.AZ;
                    hoja.Cell(filaActual, 53).Value = row.BA;
                    hoja.Cell(filaActual, 54).FormulaA1 = $"=IF(I{filaActual}=\"POTC\",\"OTC\",IF(I{filaActual}=\"FRIS\",\"FRISO\",VLOOKUP(J{filaActual},Datos!$A$9:$C$14,3,0)))"; // BB
                    hoja.Cell(filaActual, 55).Value = row.BC;
                    hoja.Cell(filaActual, 56).Value = row.BD;
                    hoja.Cell(filaActual, 57).Value = row.BE;
                    hoja.Cell(filaActual, 58).Value = row.BF; // BF
                    hoja.Cell(filaActual, 59).Value = row.BG; // BG
                    hoja.Cell(filaActual, 60).Value = row.BH;
                    hoja.Cell(filaActual, 61).FormulaA1 = $"=WEEKNUM(AR{filaActual})"; // BI
                    hoja.Cell(filaActual, 62).Value = row.BJ; // BJ
                    hoja.Cell(filaActual, 63).Value = row.BK;
                    hoja.Cell(filaActual, 64).FormulaA1 = $"=TEXT(AR{filaActual}, \"MMM\")"; // BL
                    hoja.Cell(filaActual, 65).FormulaA1 = $"=IF(MONTH(AR{filaActual})={_noMes + 1},IF(A{filaActual}=A{filaActual + 1},0,1),0)"; // BM
                    hoja.Cell(filaActual, 66).FormulaA1 = $"=IF(AJ{filaActual}=\"\",\"\",TEXT(AJ{filaActual},\"MMM\"))"; // BN
                    hoja.Cell(filaActual, 67).FormulaA1 = $"=IF(BJ{filaActual}=\"\",\"\",TEXT(BJ{filaActual},\"MMM\"))"; // BO
                    hoja.Cell(filaActual, 68).Value = row.BP; // BP
                    filaActual++;
                }

                // Ajuste de columnas y personalizacion
                hoja.RangeUsed().SetAutoFilter();

                // Guardar el archivo
                libro.SaveAs(rutaCompleta);
            }
        }
    }
}
