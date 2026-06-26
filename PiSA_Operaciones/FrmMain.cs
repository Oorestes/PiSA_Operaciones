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
    }
}
