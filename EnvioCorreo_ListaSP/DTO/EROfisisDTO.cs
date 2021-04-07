using System;
using System.Collections.Generic;
using System.Text;

namespace EnvioCorreo_ListaSP.DTO
{
    public class EROfisisDTO
    {
        public string CodigoEmpresa { get; set; }
        public string ComprobanteTesoreria { get; set; }
        public string Moneda { get; set; }
        public string CorrelativoHelm { get; set; }
        public DateTime Fecha { get; set; }
        public decimal TipoCambio { get; set; }
    }
}
