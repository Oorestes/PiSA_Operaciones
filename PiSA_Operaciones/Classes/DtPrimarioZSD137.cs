using System;

namespace PiSA_Operaciones.Classes
{
    internal class DtPrimarioZSD137
    {
        public string DocumentoComercial { get; set; }
        public float Monto { get; set; } = 0;
        public string OC { get; set; }
        public string Centro { get; set; }
        public DateTime? FechaCita { get; set; }
        public DateTime? HoraCita { get; set; }
    }
}
