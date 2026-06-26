using PiSA_Operaciones.Classes;
using System.Collections.Generic;
using System.Windows.Forms;

namespace PiSA_Operaciones
{
    public partial class FrmMain : Form
    {
        public FrmMain() => InitializeComponent();

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

        private void BtnDtPrimarios_Click(object sender, System.EventArgs e)
        {
            if (!ValidarRutasSeleccionadas()) return;

            try
            {
                // Leer archivos
                List<DtPrimarioRecord> _listaDtPrimarios = DtPrimario.LeerZLO10(TbxRutaZLO10.Text);
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

                // Generar archivo Excel de salida
                DtPrimario.CrearExcel(_listaDtPrimarios);

                MessageBox.Show("Proceso finalizado!\nEl archivo generado se encuentra en el escritorio", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch
            {
                MessageBox.Show("Ocurrió un error al procesar los archivos!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
