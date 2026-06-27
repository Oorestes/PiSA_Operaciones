using PiSA_Operaciones.Classes;
using Syncfusion.WinForms.DataGrid.Enums;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PiSA_Operaciones
{
    public partial class FrmMain : Form
    {
        private static List<DtPrimarioRecord> _listaDtPrimarios = new List<DtPrimarioRecord>();
        private string _lastExecutedReport = string.Empty;
        public FrmMain()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("es-MX");
            System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("es-MX");

            InitializeComponent();

            BtnExportar.Visible = false;
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
            if (_lastExecutedReport == "DtPrimarios")
            {
                if(mapaColumnasDtPrimario.TryGetValue(e.Column.MappingName, out ConfiguracionColumna config)){
                    e.Column.HeaderText = config.Titulo;

                    switch (config.Formato)
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
            }
        }

        private void BtnExportar_Click(object sender, System.EventArgs e)
        {
            if(_lastExecutedReport == "DtPrimarios") DtPrimario.CrearExcel(_listaDtPrimarios);

            MessageBox.Show("Proceso finalizado!\nEl archivo generado se encuentra en el escritorio", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
