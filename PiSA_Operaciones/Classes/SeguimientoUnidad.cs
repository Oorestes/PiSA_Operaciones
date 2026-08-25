using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PiSA_Operaciones.Classes
{
    internal class SeguimientoUnidad
    {
        internal static List<SeguimientoUnidadRecord> LeerZLO22N(string _ruta)
        {
            var _seguimientoUnidades = new List<SeguimientoUnidadRecord>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var seguimientoUnidad = new SeguimientoUnidadRecord{
                        A = fila.Cell(6).GetString(),
                        B = fila.Cell(15).GetString(),
                        C = fila.Cell(1).GetString(),
                        D = $"{fila.Cell(1).GetString()}-{fila.Cell(16).GetString()}",
                        E = fila.Cell(28).GetString(),
                        F = string.Empty,
                        G = null,
                        H = null,
                        I = null,
                        J = fila.Cell(23).GetString(),
                        K = fila.Cell(32).GetValue<int>(),
                        L = 0,
                        M = fila.Cell(9).GetString(),
                        N = fila.Cell(10).GetString(),
                        O = null,
                        P = 0,
                        Q = string.Empty,
                        R = string.Empty,
                        S = string.Empty,
                        T = string.Empty,
                        Material = fila.Cell(16).GetString(),
                        Factura = fila.Cell(34).GetString()
                    };

                    _seguimientoUnidades.Add(seguimientoUnidad);
                }
            }

            return _seguimientoUnidades;
        }

        internal static List<SeguimientoUnidadZSD137> LeerZSD137(string _ruta)
        {
            var _listaZSD137 = new List<SeguimientoUnidadZSD137>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach(var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var seguimientoZSD137 = new SeguimientoUnidadZSD137
                    {
                        Pedido = fila.Cell(1).GetString(),
                        FechaEntrega = fila.Cell(14).IsEmpty() ? DateTime.MinValue : fila.Cell(14).GetValue<DateTime>(),
                        Material = fila.Cell(22).GetString(),
                        PrecioNeto = fila.Cell(29).IsEmpty() ? 0 : fila.Cell(29).GetValue<float>(),
                        FechaCita = fila.Cell(49).IsEmpty() ? DateTime.MinValue : fila.Cell(49).GetValue<DateTime>(),
                        HoraCita = fila.Cell(50).IsEmpty() ? DateTime.MinValue : fila.Cell(50).GetValue<DateTime>()
                    };
                    seguimientoZSD137.PrecioNeto = (float)Math.Round(seguimientoZSD137.PrecioNeto, 2, MidpointRounding.AwayFromZero);

                    _listaZSD137.Add(seguimientoZSD137);
                }
            }
            _listaZSD137.Reverse();

            return _listaZSD137;
        }

        internal static List<SeguimientoUnidadZLO10> LeerZLO10(string _ruta)
        {
            var _listaZLO10 = new List<SeguimientoUnidadZLO10>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var seguimientoZLO10 = new SeguimientoUnidadZLO10
                    {
                        Entrega = fila.Cell(1).GetString(),
                        DT = fila.Cell(2).GetString()
                    };

                    _listaZLO10.Add(seguimientoZLO10);
                }
            }
            _listaZLO10.Reverse();

            return _listaZLO10;
        }

        internal static List<SeguimientoUnidadVLO6F> LeerVLO6F(string _ruta)
        {
            var _listaVLO6F = new List<SeguimientoUnidadVLO6F>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var seguimientoVLO6F = new SeguimientoUnidadVLO6F
                    {
                        Entrega = fila.Cell(1).GetString(),
                        Destinatario = fila.Cell(5).GetString()
                    };

                    _listaVLO6F.Add(seguimientoVLO6F);
                }
            }

            return _listaVLO6F;
        }

        internal static void CrearExcel(List<SeguimientoUnidadRecord> _listaSeguimientoUnidades, bool _generaCC)
        {
            using (var libro = new XLWorkbook())
            {
                var hoy = DateTime.Now;
                string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string rutaCompleta = Path.Combine(rutaEscritorio, $"Seguimiento de unidades DT {hoy.Hour}{hoy.Minute}{hoy.Second}.xlsx");

                var hoja = libro.Worksheets.Add("Seguimiento de unidades DT");

                hoja.Cell(1, 1).Value = "Cliente";
                hoja.Cell(1, 2).Value = "Orden de compra";
                hoja.Cell(1, 3).Value = "Pedido";
                hoja.Cell(1, 4).Value = "Pedido&Item";
                hoja.Cell(1, 5).Value = "Entrega";
                hoja.Cell(1, 6).Value = "DT";
                hoja.Cell(1, 7).Value = "Fecha Cita";
                hoja.Cell(1, 8).Value = "Hora Cita";
                hoja.Cell(1, 9).Value = "Factura";
                hoja.Cell(1, 10).Value = "Centro";
                hoja.Cell(1, 11).Value = "Cantidad Entregada";
                hoja.Cell(1, 12).Value = "Precio Neto";
                hoja.Cell(1, 13).Value = "Organización";
                hoja.Cell(1, 14).Value = "Canal";
                hoja.Cell(1, 15).Value = "Fecha Pref Entrega";
                hoja.Cell(1, 16).Value = "Monto Real";
                hoja.Cell(1, 17).Value = "Incidencias";
                hoja.Cell(1, 18).Value = "Tipo de Incidencias";
                hoja.Cell(1, 19).Value = "Fecha Reprogramación";
                hoja.Cell(1, 20).Value = "Destinatario";

                int _filaactual = 2;
                foreach (var fila in _listaSeguimientoUnidades)
                {
                    hoja.Cell(_filaactual, 1).Value = fila.A;
                    hoja.Cell(_filaactual, 2).Value = fila.B;
                    hoja.Cell(_filaactual, 3).Value = fila.C;
                    hoja.Cell(_filaactual, 4).Value = fila.D;
                    hoja.Cell(_filaactual, 5).Value = fila.E;
                    hoja.Cell(_filaactual, 6).Value = fila.F;
                    hoja.Cell(_filaactual, 7).Value = fila.G == DateTime.MinValue ? null : fila.G;
                    hoja.Cell(_filaactual, 8).Value = fila.H;
                    hoja.Cell(_filaactual, 9).Value = fila.Factura;
                    hoja.Cell(_filaactual, 10).Value = fila.J;
                    hoja.Cell(_filaactual, 11).Value = fila.K;
                    hoja.Cell(_filaactual, 12).Value = fila.L;
                    hoja.Cell(_filaactual, 13).Value = fila.M;
                    hoja.Cell(_filaactual, 14).Value = fila.N;
                    hoja.Cell(_filaactual, 15).Value = fila.O;
                    hoja.Cell(_filaactual, 16).Value = fila.P;
                    hoja.Cell(_filaactual, 17).Value = fila.Q;
                    hoja.Cell(_filaactual, 18).Value = fila.R;
                    hoja.Cell(_filaactual, 19).Value = fila.S;
                    hoja.Cell(_filaactual, 20).Value = fila.T;
                    _filaactual++;
                }

                if (_generaCC)
                {
                    var hojaCC = libro.Worksheets.Add("Centro de control");

                    hojaCC.Cell(1, 1).Value = "Cliente";
                    hojaCC.Cell(1, 2).Value = "DT";
                    hojaCC.Cell(1, 3).Value = "Centro";
                    hojaCC.Cell(1, 4).Value = "Fecha cita";
                    hojaCC.Cell(1, 5).Value = "Hora cita";

                    hojaCC.Range("A1:E1").Style.Font.Bold = true;
                    hojaCC.Columns().AdjustToContents();

                    int _filaCC = 2;
                    foreach (var fila in _listaSeguimientoUnidades)
                    {
                        if (fila.F == string.Empty) continue;
                        if (fila.G == DateTime.MinValue) continue;
                        if (hoy.Month != fila.G.Value.Month) continue;

                        hojaCC.Cell(_filaCC, 1).Value = fila.A;
                        hojaCC.Cell(_filaCC, 2).Value = fila.F;
                        hojaCC.Cell(_filaCC, 3).Value = fila.J;
                        hojaCC.Cell(_filaCC, 4).Value = fila.G == DateTime.MinValue ? null : fila.G;
                        hojaCC.Cell(_filaCC, 5).Value = fila.H;
                        _filaCC++;
                    }

                    hojaCC.Range("A1:E1").Style.Font.Bold = true;
                    hojaCC.Column(5).Style.NumberFormat.Format = "h:mm:ss AM/PM";
                    hojaCC.Columns().AdjustToContents();
                    hojaCC.RangeUsed().SetAutoFilter();
                }

                hoja.Range("A1:T1").Style.Font.Bold = true;
                hoja.Column(8).Style.NumberFormat.Format = "h:mm:ss AM/PM";
                hoja.Column(12).Style.NumberFormat.Format = "$ #,##0.00";
                hoja.Column(16).Style.NumberFormat.Format = "$ #,##0.00";
                hoja.Columns().AdjustToContents();
                hoja.RangeUsed().SetAutoFilter();

                // Guardar el archivo
                libro.SaveAs(rutaCompleta);
            }
        }
    }
}
