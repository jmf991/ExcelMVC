using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDataTableToExcelInMVC4.Models
{
    public class ExportDataTableToExcelModel
    {
        public int CodigoCliente {get;set;}
        public int Importe {get;set;}
        public string Concepto {get;set;}
        public string TipoIva {get;set;}
        public int CodigoMarca { get; set; }
    }
}