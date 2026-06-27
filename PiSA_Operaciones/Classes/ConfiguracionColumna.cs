namespace PiSA_Operaciones.Classes
{
    public enum TipoFormato
    {
        Texto,
        Moneda,
        Fecha,
        Hora,
        Entero
    }

    internal class ConfiguracionColumna
    {
        public string Titulo { get; set; }
        public TipoFormato Formato { get; set; }
    }
}
