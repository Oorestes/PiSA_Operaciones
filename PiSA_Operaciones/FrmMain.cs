using PiSA_Operaciones.Classes;
using Syncfusion.WinForms.DataGrid.Enums;
using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PiSA_Operaciones
{
    enum Meses
    {
        Enero,
        Febrero,
        Marzo,
        Abril,
        Mayo,
        Junio,
        Julio,
        Agosto,
        Septiembre,
        Octubre,
        Noviembre,
        Diciembre
    }

    public partial class FrmMain : Form
    {
        private static List<DtPrimarioRecord> _listaDtPrimarios = new List<DtPrimarioRecord>();
        private static List<AlcanceMetaRecord> _listaAlcanceMeta = new List<AlcanceMetaRecord>();
        private static List<SeguimientoUnidadRecord> _listaSeguimientoUnidades = new List<SeguimientoUnidadRecord>();
        private string _lastExecutedReport = string.Empty;
        public FrmMain()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-MX");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("es-MX");

            InitializeComponent();

            BtnExportar.Visible = false;
            ComboMeses.Items.AddRange(Enum.GetNames(typeof(Meses)));
        }

        private void BtnRutaZLO10_Click(object sender, System.EventArgs e) => TbxRutaZLO10.Text = GetFilePath();

        private void BtnRutaZSD137_Click(object sender, System.EventArgs e) => TbxRutaZSD137.Text = GetFilePath();

        private void BtnRutaVLO6F_Click(object sender, System.EventArgs e) => TbxRutaVLO6F.Text = GetFilePath();

        private void BtnRutaZLO22N_Click(object sender, System.EventArgs e) => TbxRutaZLO22N.Text = GetFilePath();

        private string GetFilePath()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls";
                ofd.Title = "Seleccionar archivo de Excel";

                if (ofd.ShowDialog() == DialogResult.OK) return ofd.FileName;
            }
            return string.Empty;
        }

        private bool ValidarRutasSeleccionadas()
        {
            if (TbxRutaZLO10.Text == string.Empty
                || TbxRutaZSD137.Text == string.Empty
                || TbxRutaVLO6F.Text == string.Empty
                || TbxRutaZLO22N.Text == string.Empty)
            {
                MessageBox.Show("Se debe seleccionar la ruta de todos los archivos!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private void LimpiarDataGrid()
        {
            sfDataGridMain.Columns.Clear();
            sfDataGridMain.DataSource = null;
            _lastExecutedReport = string.Empty;
            BtnExportar.Visible = false;
        }

        private readonly Dictionary<string, ConfiguracionColumna> mapaColumnasDtPrimario = new Dictionary<string, ConfiguracionColumna>
        {
            { "A", new ConfiguracionColumna { Titulo = "CLIENTE", Formato = TipoFormato.Texto } },
            { "B", new ConfiguracionColumna { Titulo = "ORG.", Formato = TipoFormato.Texto } },
            { "C", new ConfiguracionColumna { Titulo = "CANAL", Formato = TipoFormato.Texto } },
            { "D", new ConfiguracionColumna { Titulo = "OC", Formato = TipoFormato.Texto } },
            { "E", new ConfiguracionColumna { Titulo = "PEDIDO", Formato = TipoFormato.Texto } },
            { "F", new ConfiguracionColumna { Titulo = "ENTREGA", Formato = TipoFormato.Texto } },
            { "G", new ConfiguracionColumna { Titulo = "CENTRO", Formato = TipoFormato.Texto } },
            { "H", new ConfiguracionColumna { Titulo = "DT PRIMARIO", Formato = TipoFormato.Texto } },
            { "I", new ConfiguracionColumna { Titulo = "DT PRIMARIO 2", Formato = TipoFormato.Texto } },
            { "J", new ConfiguracionColumna { Titulo = "DT SECUNDARIO", Formato = TipoFormato.Texto } },
            { "K", new ConfiguracionColumna { Titulo = "FECHA CITA", Formato = TipoFormato.Fecha } },
            { "L", new ConfiguracionColumna { Titulo = "HORA CITA", Formato = TipoFormato.Hora } },
            { "M", new ConfiguracionColumna { Titulo = "LUGAR DE ENTREGA", Formato = TipoFormato.Texto } },
            { "N", new ConfiguracionColumna { Titulo = "MONTO", Formato = TipoFormato.Moneda } },
            { "O", new ConfiguracionColumna { Titulo = "ESTATUS", Formato = TipoFormato.Texto } }
        };

        private void BtnDtPrimarios_Click(object sender, System.EventArgs e)
        {
            if (!ValidarRutasSeleccionadas()) return;

            try
            {
                // Limpiar y vaciar el DataGrid antes de procesar los datos
                LimpiarDataGrid();
                _lastExecutedReport = "DtPrimarios";

                // Leer archivos
                _listaDtPrimarios = DtPrimario.LeerZLO10(TbxRutaZLO10.Text);
                List<DtPrimarioVLO6F> _listaVL06F = DtPrimario.LeerVL06F(TbxRutaVLO6F.Text);
                List<DtPrimarioZLO22N> _listaZLO22N = DtPrimario.LeerZLO22N(TbxRutaZLO22N.Text);
                List<DtPrimarioZSD137> _listaZSD137 = DtPrimario.LeerZSD137(TbxRutaZSD137.Text);

                // Procesado de datos
                var listDocumentos = new List<string>();
                for (int i = 0; i < _listaDtPrimarios.Count; i++)
                {
                    var vl06f = _listaVL06F.Find(x => x.Entrega == _listaDtPrimarios[i].F);
                    var zlo22n = _listaZLO22N.Find(x => x.Entrega == _listaDtPrimarios[i].F);
                    var zsd137 = _listaZSD137.Find(x => x.DocumentoComercial == zlo22n.DocumentoComercial);

                    _listaDtPrimarios[i].D = zsd137 != null ? zsd137.OC : string.Empty;
                    _listaDtPrimarios[i].E = zlo22n != null ? zlo22n.DocumentoComercial : string.Empty;
                    _listaDtPrimarios[i].G = zsd137 != null ? zsd137.Centro : string.Empty;
                    _listaDtPrimarios[i].K = zsd137 != null ? zsd137.FechaCita : null;
                    _listaDtPrimarios[i].L = zsd137 != null ? zsd137.HoraCita : null;
                    _listaDtPrimarios[i].M = vl06f != null ? vl06f.LugarDestinatario : string.Empty;
                    _listaDtPrimarios[i].N = zsd137 != null ? zsd137.Monto : 0;

                    listDocumentos.Add(_listaDtPrimarios[i].E);
                }

                for (int i = 0; i < _listaDtPrimarios.Count; i++)
                {
                    if (listDocumentos.LastIndexOf(_listaDtPrimarios[i].E) != i) _listaDtPrimarios[i].N = 0;
                }

                // Llenar Datagrid con la información procesada
                sfDataGridMain.DataSource = _listaDtPrimarios;
                sfDataGridMain.AutoGenerateColumns = true;
                sfDataGridMain.AllowFiltering = true;
                sfDataGridMain.FilterRowPosition = RowPosition.Top;
                sfDataGridMain.AutoSizeController.AutoSizeCalculationMode = AutoSizeCalculationMode.SmartFit;
                sfDataGridMain.AutoSizeColumnsMode = AutoSizeColumnsMode.AllCells;

                // Habilitar botón para exportar a Excel
                BtnExportar.Visible = true;
            }
            catch
            {
                MessageBox.Show("Ocurrió un error al procesar los archivos!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void sfDataGridMain_AutoGeneratingColumn(object sender, Syncfusion.WinForms.DataGrid.Events.AutoGeneratingColumnArgs e)
        {
            ConfiguracionColumna _configuracionColumna = new ConfiguracionColumna();

            if(_lastExecutedReport == "DtPrimarios" && mapaColumnasDtPrimario.TryGetValue(e.Column.MappingName, out ConfiguracionColumna DtPrimariosConfig))
                _configuracionColumna = DtPrimariosConfig;

            if (_lastExecutedReport == "AlcanceMeta" && mapaColumnasAlcanceMeta.TryGetValue(e.Column.MappingName, out ConfiguracionColumna AlcanceMetaConfig))
                _configuracionColumna = AlcanceMetaConfig;

            if (_lastExecutedReport == "SeguimientoUnidades" && mapaColumnasSeguimientoUnidades.TryGetValue(e.Column.MappingName, out ConfiguracionColumna SeguimientoUnidadesConfig))
                _configuracionColumna = SeguimientoUnidadesConfig;

            if (_configuracionColumna.Titulo == null || _configuracionColumna.Titulo == string.Empty) e.Cancel = true;

            e.Column.HeaderText = _configuracionColumna.Titulo;

            switch (_configuracionColumna.Formato)
            {
                case TipoFormato.Texto:
                    e.Column.CellStyle.HorizontalAlignment = HorizontalAlignment.Left;
                    break;

                case TipoFormato.Moneda:
                    e.Column.Format = "C2";
                    e.Column.CellStyle.HorizontalAlignment = HorizontalAlignment.Right;
                    break;

                case TipoFormato.Fecha:
                    e.Column.Format = "d";
                    e.Column.CellStyle.HorizontalAlignment = HorizontalAlignment.Center;
                    break;

                case TipoFormato.Hora:
                    e.Column.Format = "h:mm:ss AM/PM";
                    e.Column.CellStyle.HorizontalAlignment = HorizontalAlignment.Center;
                    break;
            }
        }

        private void BtnAlcanceMeta_Click(object sender, System.EventArgs e)
        {
            if (!ValidarRutasSeleccionadas()) return;
            if(ComboMeses.SelectedIndex == -1)
            {
                MessageBox.Show("Se debe seleccionar un mes para poder continuar", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                // Limpiar y vaciar el DataGrid antes de procesar los datos
                LimpiarDataGrid();
                _lastExecutedReport = "AlcanceMeta";

                // Leer archivos
                List<AlcanceMetaZLO10> listaZLO10 = AlcanceMeta.LeerZLO10(TbxRutaZLO10.Text);
                _listaAlcanceMeta = AlcanceMeta.Leer22N(TbxRutaZLO22N.Text);
                List<AlcanceMetaVLO6F> listaVL06F = AlcanceMeta.LeerVL06F(TbxRutaVLO6F.Text);
                List<AlcanceMetaZSD137> listaZSD137 = AlcanceMeta.LeerZSD137(TbxRutaZSD137.Text);

                // Procesado de datos
                for (int i = 0; i < _listaAlcanceMeta.Count; i++)
                {
                    var dt = listaZLO10.Find(x => x.Entrega == _listaAlcanceMeta[i].AB);
                    var vl06f = listaVL06F.Find(x => x.Entrega == _listaAlcanceMeta[i].AB);
                    var zsd137 = listaZSD137.Find(x => x.DocumentoComercial == _listaAlcanceMeta[i].A);

                    _listaAlcanceMeta[i].C = _listaAlcanceMeta[i].C == DateTime.MinValue ? null : _listaAlcanceMeta[i].C;
                    _listaAlcanceMeta[i].N = _listaAlcanceMeta[i].N == DateTime.MinValue ? null : _listaAlcanceMeta[i].N;
                    _listaAlcanceMeta[i].AE = _listaAlcanceMeta[i].AE == DateTime.MinValue ? null : _listaAlcanceMeta[i].AE;
                    _listaAlcanceMeta[i].AG = _listaAlcanceMeta[i].AG == DateTime.MinValue ? null : _listaAlcanceMeta[i].AG;
                    _listaAlcanceMeta[i].AJ = _listaAlcanceMeta[i].AJ == DateTime.MinValue ? null : _listaAlcanceMeta[i].AJ;
                    _listaAlcanceMeta[i].AL = _listaAlcanceMeta[i].AL == DateTime.MinValue ? null : _listaAlcanceMeta[i].AL;
                    _listaAlcanceMeta[i].AM = _listaAlcanceMeta[i].AM == DateTime.MinValue ? null : _listaAlcanceMeta[i].AM;
                    _listaAlcanceMeta[i].AN = _listaAlcanceMeta[i].AN == DateTime.MinValue ? null : _listaAlcanceMeta[i].AN;
                    _listaAlcanceMeta[i].AR = _listaAlcanceMeta[i].AR == DateTime.MinValue ? null : _listaAlcanceMeta[i].AR;
                    _listaAlcanceMeta[i].AS = _listaAlcanceMeta[i].AS == DateTime.MinValue ? null : _listaAlcanceMeta[i].AS;
                    _listaAlcanceMeta[i].BF = dt != null ? dt.DT : string.Empty;
                    _listaAlcanceMeta[i].BG = vl06f != null ? vl06f.StatusMovimiento : string.Empty;
                    _listaAlcanceMeta[i].BJ = zsd137 != null ? zsd137.Fecha : null;
                    _listaAlcanceMeta[i].BP = vl06f != null ? vl06f.LugarDestino : string.Empty;
                }

                // Llenar Datagrid con la información procesada
                sfDataGridMain.DataSource = _listaAlcanceMeta;
                sfDataGridMain.AutoGenerateColumns = true;
                sfDataGridMain.AllowFiltering = true;
                sfDataGridMain.FilterRowPosition = RowPosition.Top;
                sfDataGridMain.AutoSizeController.AutoSizeCalculationMode = AutoSizeCalculationMode.SmartFit;
                sfDataGridMain.AutoSizeColumnsMode = AutoSizeColumnsMode.AllCells;

                // Habilitar botón para exportar a Excel
                BtnExportar.Visible = true;
            }
            catch
            {
                MessageBox.Show("Ocurrió un error al procesar los archivos!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnExportar_Click(object sender, System.EventArgs e)
        {
            if (_lastExecutedReport == "DtPrimarios") DtPrimario.CrearExcel(_listaDtPrimarios);
            if (_lastExecutedReport == "AlcanceMeta") AlcanceMeta.CrearExcel(_listaAlcanceMeta, ComboMeses.SelectedIndex);
            if (_lastExecutedReport == "SeguimientoUnidades") SeguimientoUnidad.CrearExcel(_listaSeguimientoUnidades, CheckGenerarCC.Checked);
            
            MessageBox.Show("Proceso finalizado!\nEl archivo generado se encuentra en el escritorio", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private readonly Dictionary<string, ConfiguracionColumna> mapaColumnasAlcanceMeta = new Dictionary<string, ConfiguracionColumna>
        {
            { "A", new ConfiguracionColumna { Titulo = "Documento comercial", Formato = TipoFormato.Texto } },
            { "B", new ConfiguracionColumna { Titulo = "Posición (SD)", Formato = TipoFormato.Texto } },
            { "C", new ConfiguracionColumna { Titulo = "Fecha de pedido", Formato = TipoFormato.Texto } },
            { "D", new ConfiguracionColumna { Titulo = "Clase doc.ventas", Formato = TipoFormato.Texto } },
            { "E", new ConfiguracionColumna { Titulo = "Cliente", Formato = TipoFormato.Texto } },
            { "F", new ConfiguracionColumna { Titulo = "Solicitante", Formato = TipoFormato.Texto } },
            { "G", new ConfiguracionColumna { Titulo = "Destino", Formato = TipoFormato.Texto } },
            { "H", new ConfiguracionColumna { Titulo = "Destinatario de Mercancias", Formato = TipoFormato.Texto } },
            { "I", new ConfiguracionColumna { Titulo = "Organización ventas", Formato = TipoFormato.Texto } },
            { "J", new ConfiguracionColumna { Titulo = "Canal distribución", Formato = TipoFormato.Texto } },
            { "K", new ConfiguracionColumna { Titulo = "Oficina de ventas", Formato = TipoFormato.Texto } },
            { "L", new ConfiguracionColumna { Titulo = "Grupo de vendedores", Formato = TipoFormato.Texto } },
            { "M", new ConfiguracionColumna { Titulo = "Zona de ventas", Formato = TipoFormato.Texto } },
            { "N", new ConfiguracionColumna { Titulo = "Fecha pref.entrega", Formato = TipoFormato.Texto } },
            { "O", new ConfiguracionColumna { Titulo = "N° pedido cliente", Formato = TipoFormato.Texto } },
            { "P", new ConfiguracionColumna { Titulo = "Material", Formato = TipoFormato.Texto } },
            { "Q", new ConfiguracionColumna { Titulo = "Texto breve de material", Formato = TipoFormato.Texto } },
            { "R", new ConfiguracionColumna { Titulo = "Grupo de artículos", Formato = TipoFormato.Texto } },
            { "S", new ConfiguracionColumna { Titulo = "Cantidad de pedido", Formato = TipoFormato.Texto } },
            { "T", new ConfiguracionColumna { Titulo = "Un.medida venta", Formato = TipoFormato.Texto } },
            { "U", new ConfiguracionColumna { Titulo = "Cantidad-acum-confir", Formato = TipoFormato.Texto } },
            { "V", new ConfiguracionColumna { Titulo = "Cantidad pedida - cantidad confirmada", Formato = TipoFormato.Texto } },
            { "W", new ConfiguracionColumna { Titulo = "Centro", Formato = TipoFormato.Texto } },
            { "X", new ConfiguracionColumna { Titulo = "Valor neto", Formato = TipoFormato.Texto } },
            { "Y", new ConfiguracionColumna { Titulo = "Precio neto", Formato = TipoFormato.Texto } },
            { "Z", new ConfiguracionColumna { Titulo = "Motivo de rechazo", Formato = TipoFormato.Texto } },
            { "AA", new ConfiguracionColumna { Titulo = "Status total crédito", Formato = TipoFormato.Texto } },
            { "AB", new ConfiguracionColumna { Titulo = "Entrega", Formato = TipoFormato.Texto } },
            { "AC", new ConfiguracionColumna { Titulo = "Posición de entrega", Formato = TipoFormato.Texto } },
            { "AD", new ConfiguracionColumna { Titulo = "Tipo de Documento", Formato = TipoFormato.Texto } },
            { "AE", new ConfiguracionColumna { Titulo = "Fecha de Entrega", Formato = TipoFormato.Texto } },
            { "AF", new ConfiguracionColumna { Titulo = "Cantidad entrega", Formato = TipoFormato.Texto } },
            { "AG", new ConfiguracionColumna { Titulo = "Fe.mov.mcía.real", Formato = TipoFormato.Texto } },
            { "AH", new ConfiguracionColumna { Titulo = "Factura", Formato = TipoFormato.Texto } },
            { "AI", new ConfiguracionColumna { Titulo = "Tipo de documento2", Formato = TipoFormato.Texto } },
            { "AJ", new ConfiguracionColumna { Titulo = "Fecha de Factura", Formato = TipoFormato.Texto } },
            { "AK", new ConfiguracionColumna { Titulo = "Programa para control", Formato = TipoFormato.Texto } },
            { "AL", new ConfiguracionColumna { Titulo = "Fe.act.desp.expd.", Formato = TipoFormato.Texto } },
            { "AM", new ConfiguracionColumna { Titulo = "Inic.actual transp.", Formato = TipoFormato.Texto } },
            { "AN", new ConfiguracionColumna { Titulo = "Inicio en UTC", Formato = TipoFormato.Texto } },
            { "AO", new ConfiguracionColumna { Titulo = "Tipo de contratista", Formato = TipoFormato.Texto } },
            { "AP", new ConfiguracionColumna { Titulo = "Clase de transporte", Formato = TipoFormato.Texto } },
            { "AQ", new ConfiguracionColumna { Titulo = "Agente servicios", Formato = TipoFormato.Texto } },
            { "AR", new ConfiguracionColumna { Titulo = "Creado el", Formato = TipoFormato.Texto } },
            { "AS", new ConfiguracionColumna { Titulo = "Fecha documento", Formato = TipoFormato.Texto } },
            { "AT", new ConfiguracionColumna { Titulo = "Tp.doc.subsiguiente", Formato = TipoFormato.Texto } },
            { "AU", new ConfiguracionColumna { Titulo = "Tp.doc.subsiguiente3", Formato = TipoFormato.Texto } },
            { "AV", new ConfiguracionColumna { Titulo = "Carácter 1", Formato = TipoFormato.Texto } },
            { "AW", new ConfiguracionColumna { Titulo = "Motivo pedido", Formato = TipoFormato.Texto } },
            { "AX", new ConfiguracionColumna { Titulo = "Denominación", Formato = TipoFormato.Texto } },
            { "AY", new ConfiguracionColumna { Titulo = "Pedido Bloqueo", Formato = TipoFormato.Texto } },
            { "AZ", new ConfiguracionColumna { Titulo = "Piezas negadas", Formato = TipoFormato.Texto } },
            { "BA", new ConfiguracionColumna { Titulo = "Tipo Pedido", Formato = TipoFormato.Texto } },
            { "BB", new ConfiguracionColumna { Titulo = "Mercados", Formato = TipoFormato.Texto } },
            { "BC", new ConfiguracionColumna { Titulo = "Suma de Entrega", Formato = TipoFormato.Texto } },
            { "BD", new ConfiguracionColumna { Titulo = "Venta real", Formato = TipoFormato.Texto } },
            { "BE", new ConfiguracionColumna { Titulo = "Aplica", Formato = TipoFormato.Texto } },
            { "BF", new ConfiguracionColumna { Titulo = "Transporte", Formato = TipoFormato.Texto } },
            { "BG", new ConfiguracionColumna { Titulo = "Para facturar", Formato = TipoFormato.Texto } },
            { "BH", new ConfiguracionColumna { Titulo = "Venta sin confirmar", Formato = TipoFormato.Texto } },
            { "BI", new ConfiguracionColumna { Titulo = "SEMANA", Formato = TipoFormato.Texto } },
            { "BJ", new ConfiguracionColumna { Titulo = "DATOS B", Formato = TipoFormato.Texto } },
            { "BK", new ConfiguracionColumna { Titulo = "Estatus Factura", Formato = TipoFormato.Texto } },
            { "BL", new ConfiguracionColumna { Titulo = "MES creación", Formato = TipoFormato.Texto } },
            { "BM", new ConfiguracionColumna { Titulo = "Num Pedido", Formato = TipoFormato.Texto } },
            { "BN", new ConfiguracionColumna { Titulo = "MES Factura", Formato = TipoFormato.Texto } },
            { "BO", new ConfiguracionColumna { Titulo = "MES de Entrega", Formato = TipoFormato.Texto } },
            { "BP", new ConfiguracionColumna { Titulo = "Lugar Destino", Formato = TipoFormato.Texto } },
        };

        private void BtnSeguimientoUni_Click(object sender, EventArgs e)
        {
            if (!ValidarRutasSeleccionadas()) return;

            try
            {
                // Limpiar y vaciar el DataGrid antes de procesar los datos
                LimpiarDataGrid();
                _lastExecutedReport = "SeguimientoUnidades";

                // Leer archivos
                _listaSeguimientoUnidades = SeguimientoUnidad.LeerZLO22N(TbxRutaZLO22N.Text);
                List<SeguimientoUnidadZSD137> listaZSD137 = SeguimientoUnidad.LeerZSD137(TbxRutaZSD137.Text);
                List<SeguimientoUnidadZLO10> listaZLO10 = SeguimientoUnidad.LeerZLO10(TbxRutaZLO10.Text);
                List<SeguimientoUnidadVLO6F> listaVLO6F = SeguimientoUnidad.LeerVLO6F(TbxRutaVLO6F.Text);

                // Procesado de datos
                for (int i = 0; i < _listaSeguimientoUnidades.Count; i++)
                {
                    var entrega = _listaSeguimientoUnidades[i].E;
                    var zlo10 = listaZLO10.Find(x => x.Entrega == entrega);
                    var zlo6f = listaVLO6F.Find(x => x.Entrega == entrega);
                    var zsd137 = listaZSD137.Find(x => x.Material == _listaSeguimientoUnidades[i].Material && x.Pedido == _listaSeguimientoUnidades[i].C);

                    _listaSeguimientoUnidades[i].F = zlo10 == null ? string.Empty : zlo10.DT;
                    _listaSeguimientoUnidades[i].T = zlo6f == null ? string.Empty : zlo6f.Destinatario;

                    if (zsd137 != null)
                    {
                        _listaSeguimientoUnidades[i].G = zsd137.FechaCita;
                        _listaSeguimientoUnidades[i].H = zsd137.HoraCita;
                        _listaSeguimientoUnidades[i].L = zsd137.PrecioNeto;
                        _listaSeguimientoUnidades[i].O = zsd137.FechaEntrega;
                        _listaSeguimientoUnidades[i].P = _listaSeguimientoUnidades[i].L * _listaSeguimientoUnidades[i].K;
                    }
                }

                // Llenar Datagrid con la información procesada
                sfDataGridMain.DataSource = _listaSeguimientoUnidades;
                sfDataGridMain.AutoGenerateColumns = true;
                sfDataGridMain.AllowFiltering = true;
                sfDataGridMain.FilterRowPosition = RowPosition.Top;
                sfDataGridMain.AutoSizeController.AutoSizeCalculationMode = AutoSizeCalculationMode.SmartFit;
                sfDataGridMain.AutoSizeColumnsMode = AutoSizeColumnsMode.AllCells;

                // Habilitar botón para exportar a Excel
                BtnExportar.Visible = true;
            }
            catch
            {
                MessageBox.Show("Ocurrió un error al procesar los archivos!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private readonly Dictionary<string, ConfiguracionColumna> mapaColumnasSeguimientoUnidades = new Dictionary<string, ConfiguracionColumna>
        {
            { "A", new ConfiguracionColumna { Titulo = "Cliente", Formato = TipoFormato.Texto } },
            { "B", new ConfiguracionColumna { Titulo = "Orden de compra", Formato = TipoFormato.Texto } },
            { "C", new ConfiguracionColumna { Titulo = "Pedido", Formato = TipoFormato.Texto } },
            { "D", new ConfiguracionColumna { Titulo = "Pedido&Item", Formato = TipoFormato.Texto } },
            { "E", new ConfiguracionColumna { Titulo = "Entrega", Formato = TipoFormato.Texto } },
            { "F", new ConfiguracionColumna { Titulo = "DT", Formato = TipoFormato.Texto } },
            { "G", new ConfiguracionColumna { Titulo = "Fecha cita", Formato = TipoFormato.Fecha } },
            { "H", new ConfiguracionColumna { Titulo = "Hora cita", Formato = TipoFormato.Hora } },
            { "Factura", new ConfiguracionColumna { Titulo = "Factura", Formato = TipoFormato.Texto } },
            { "J", new ConfiguracionColumna { Titulo = "Centro", Formato = TipoFormato.Texto } },
            { "K", new ConfiguracionColumna { Titulo = "Cantidad entregada", Formato = TipoFormato.Entero } },
            { "L", new ConfiguracionColumna { Titulo = "Precio neto", Formato = TipoFormato.Moneda } },
            { "M", new ConfiguracionColumna { Titulo = "Organización", Formato = TipoFormato.Texto } },
            { "N", new ConfiguracionColumna { Titulo = "Canal", Formato = TipoFormato.Texto } },
            { "O", new ConfiguracionColumna { Titulo = "Fecha pref entrega", Formato = TipoFormato.Fecha } },
            { "P", new ConfiguracionColumna { Titulo = "Monto real", Formato = TipoFormato.Moneda } },
            { "Q", new ConfiguracionColumna { Titulo = "Incidencias", Formato = TipoFormato.Texto } },
            { "R", new ConfiguracionColumna { Titulo = "Tipo de incidencias", Formato = TipoFormato.Texto } },
            { "S", new ConfiguracionColumna { Titulo = "Fecha reprogramación", Formato = TipoFormato.Fecha } },
            { "T", new ConfiguracionColumna { Titulo = "Destinatario", Formato = TipoFormato.Texto } },
        };
    }
}
