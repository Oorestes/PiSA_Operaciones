using System;

namespace PiSA_Operaciones.Classes
{
    internal class SeguimientoUnidadRecord
    {
        public string A { get; set; } // Cliente
        public string B { get; set; } // OrdenCompra
        public string C { get; set; } // Pedido
        public string D { get; set; } // PedidoItem
        public string E { get; set; } // Entrega
        public string F { get; set; } // DT
        public DateTime? G { get; set; } // FechaCita
        public DateTime? H { get; set; } // HoraCita
        public DateTime? I { get; set; } // HoraArribo
        public string Factura { get; set; }
        public string J { get; set; } // Centro
        public int K { get; set; } // Cantidad
        public float L { get; set; } // PrecioNeto
        public string M { get; set; } // Organizacion
        public string N { get; set; } // Canal
        public DateTime? O { get; set; } // FechaEntrega
        public float P { get; set; } // MontoReal
        public string Q { get; set; } // Incidencias
        public string R { get; set; } // TipoIncidencias
        public string S { get; set; } // FechaReprogramacion
        public string T { get; set; } // Destinatario
        public string Material { get; set; }
    }
}
