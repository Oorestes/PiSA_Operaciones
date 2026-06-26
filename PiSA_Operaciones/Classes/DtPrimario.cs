using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PiSA_Operaciones.Classes
{
    internal class DtPrimario
    {
        internal static List<DtPrimarioRecord> LeerZLO10(string _ruta)
        {
            var listaRecordsDtPrimarios = new List<DtPrimarioRecord>();
            var listaEntregas = new List<string>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    string entrega = fila.Cell(1).GetString();
                    string dt = fila.Cell(2).GetString();
                    bool entregaExiste = listaEntregas.IndexOf(entrega) != -1 ? true : false;

                    if (!entregaExiste)
                    {
                        listaEntregas.Add(entrega);

                        var record = new DtPrimarioRecord
                        {
                            A = fila.Cell(15).GetString(), // ZLO10 - O
                            B = fila.Cell(11).GetString(), // ZLO10 - K
                            C = fila.Cell(12).GetString(), // ZLO10 - L
                            D = string.Empty,
                            E = string.Empty,
                            F = entrega, // ZLO10 - A
                            G = string.Empty,
                            H = dt.StartsWith("9") ? string.Empty : dt, // ZLO10 - B
                            I = string.Empty,
                            J = dt.StartsWith("9") ? dt : string.Empty,
                            K = DateTime.MinValue,
                            L = DateTime.MinValue,
                            M = string.Empty,
                            N = 0,
                            O = string.Empty,
                        };
                        listaRecordsDtPrimarios.Add(record);
                        continue;
                    }

                    var _objetoExistente = listaRecordsDtPrimarios.FirstOrDefault(x => x.F == entrega);
                    if (_objetoExistente != null)
                    {
                        if (dt.StartsWith("9"))
                        {
                            _objetoExistente.J = _objetoExistente.J == string.Empty ? dt : $"{_objetoExistente.J}, {dt}";
                            continue;
                        }

                        if (_objetoExistente.H == string.Empty)
                        {
                            _objetoExistente.H = dt;
                            continue;
                        }

                        _objetoExistente.I = _objetoExistente.I == string.Empty ? dt : $"{_objetoExistente.I}, {dt}";
                    }
                }
            }
            return listaRecordsDtPrimarios;
        }

        internal static List<DtPrimarioVLO6F> LeerVL06F(string _ruta)
        {
            var vl06fList = new List<DtPrimarioVLO6F>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    if (fila.Cell(1).IsEmpty()) break;

                    var vl06f = new DtPrimarioVLO6F
                    {
                        Entrega = fila.Cell(1).GetString(),
                        LugarDestinatario = fila.Cell(5).GetString()
                    };
                    vl06fList.Add(vl06f);
                }
            }

            return vl06fList;
        }

        internal static List<DtPrimarioZLO22N> LeerZLO22N(string _ruta)
        {
            var zlo22nList = new List<DtPrimarioZLO22N>();
            var entregasRegistradas = new List<string>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    var entrega = fila.Cell(28).GetString();

                    if (entrega == string.Empty) continue;
                    if (entregasRegistradas.IndexOf(entrega) != -1) continue;

                    entregasRegistradas.Add(entrega);

                    var zlo22n = new DtPrimarioZLO22N
                    {
                        Entrega = entrega,
                        DocumentoComercial = fila.Cell(1).GetString()
                    };
                    zlo22nList.Add(zlo22n);
                }
            }

            return zlo22nList;
        }

        internal static List<DtPrimarioZSD137> LeerZSD137(string _ruta)
        {
            var zsd137List = new List<DtPrimarioZSD137>();
            var documentosExistentes = new List<string>();

            using (var libro = new XLWorkbook(_ruta))
            {
                var hoja = libro.Worksheet(1);
                var filas = hoja.RowsUsed().Skip(1);

                foreach (var fila in filas)
                {
                    var documento = fila.Cell(1).GetString();

                    if (documento == string.Empty) break;

                    if (documentosExistentes.IndexOf(documento) != -1)
                    {
                        var objetoExistente = zsd137List.FirstOrDefault(x => x.DocumentoComercial == documento);
                        objetoExistente.Monto += fila.Cell(52).GetValue<float>() * fila.Cell(29).GetValue<float>();
                        continue;
                    }

                    documentosExistentes.Add(documento);

                    var zsd137 = new DtPrimarioZSD137
                    {
                        DocumentoComercial = documento,
                        OC = fila.Cell(15).GetString(),
                        FechaCita = fila.Cell(49).IsEmpty() ? DateTime.MinValue : fila.Cell(49).GetDateTime(),
                        HoraCita = fila.Cell(50).IsEmpty() ? DateTime.MinValue : fila.Cell(50).GetValue<DateTime>(),
                        Centro = fila.Cell(30).GetString(),
                        Monto = fila.Cell(52).GetValue<float>() * fila.Cell(29).GetValue<float>()
                    };
                    zsd137List.Add(zsd137);
                }
            }

            return zsd137List;
        }

        internal static void CrearExcel(List<DtPrimarioRecord> _listaDtPrimarios)
        {
            using (var libro = new XLWorkbook())
            {
                var hoy = DateTime.Now;
                string rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                string rutaCompleta = Path.Combine(rutaEscritorio, $"DT Primarios {hoy.Hour}{hoy.Minute}{hoy.Second}.xlsx");

                // Hojas de Excel
                var hojaDtPrimarios = libro.Worksheets.Add("DT Primarios");

                // Encabezados
                hojaDtPrimarios.Cell(1, 1).Value = "CLIENTE";
                hojaDtPrimarios.Cell(1, 2).Value = "ORG.";
                hojaDtPrimarios.Cell(1, 3).Value = "CANAL";
                hojaDtPrimarios.Cell(1, 4).Value = "OC";
                hojaDtPrimarios.Cell(1, 5).Value = "PEDIDO";
                hojaDtPrimarios.Cell(1, 6).Value = "ENTREGA";
                hojaDtPrimarios.Cell(1, 7).Value = "CENTRO";
                hojaDtPrimarios.Cell(1, 8).Value = "DT PRIMARIO";
                hojaDtPrimarios.Cell(1, 9).Value = "DT PRIMARIO 2";
                hojaDtPrimarios.Cell(1, 10).Value = "DT SECUNDARIO";
                hojaDtPrimarios.Cell(1, 11).Value = "FECHA CITA";
                hojaDtPrimarios.Cell(1, 12).Value = "HORA CITA";
                hojaDtPrimarios.Cell(1, 13).Value = "LUGAR DE ENTREGA";
                hojaDtPrimarios.Cell(1, 14).Value = "MONTO";
                hojaDtPrimarios.Cell(1, 15).Value = "ESTATUS";

                hojaDtPrimarios.Range("A1:O1").Style.Font.Bold = true;
                hojaDtPrimarios.Columns().AdjustToContents();

                // Llenado de datos
                int filaActual = 2;

                foreach (var row in _listaDtPrimarios)
                {
                    hojaDtPrimarios.Cell(filaActual, 1).Value = row.A;
                    hojaDtPrimarios.Cell(filaActual, 2).Value = row.B;
                    hojaDtPrimarios.Cell(filaActual, 3).Value = row.C;
                    hojaDtPrimarios.Cell(filaActual, 4).Value = row.D;
                    hojaDtPrimarios.Cell(filaActual, 5).Value = row.E;
                    hojaDtPrimarios.Cell(filaActual, 6).Value = row.F;
                    hojaDtPrimarios.Cell(filaActual, 7).Value = row.G;
                    hojaDtPrimarios.Cell(filaActual, 8).Value = row.H;
                    hojaDtPrimarios.Cell(filaActual, 9).Value = row.I;
                    hojaDtPrimarios.Cell(filaActual, 10).Value = row.J;
                    hojaDtPrimarios.Cell(filaActual, 11).Value = row.K;
                    hojaDtPrimarios.Cell(filaActual, 12).Value = row.L;
                    hojaDtPrimarios.Cell(filaActual, 13).Value = row.M;
                    hojaDtPrimarios.Cell(filaActual, 14).Value = row.N;
                    hojaDtPrimarios.Cell(filaActual, 15).Value = row.O;
                    filaActual++;
                }

                // Ajuste de columnas y personalizacion
                hojaDtPrimarios.Column(12).Style.NumberFormat.Format = "h:mm:ss AM/PM";
                hojaDtPrimarios.Columns().AdjustToContents();
                hojaDtPrimarios.RangeUsed().SetAutoFilter();

                // Guardar el archivo
                libro.SaveAs(rutaCompleta);
            }
        }
    }
}
