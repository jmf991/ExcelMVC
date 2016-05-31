using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDataTableToExcelInMVC4.Models
{
    public class ImportarModel
    {
        //Variables de SUCCESS
        public bool SuccessGlobal { get; set; }
        public bool SuccessCodigoCliente { get; set; }
        public bool SuccessImporte { get; set; }
        public bool SuccessConcepto { get; set; }
        public bool SuccessCodigoMarca { get; set; }
        public bool SuccessTipoIva { get; set; }
        //Checker DB-Excel
        public bool SuccesCodigoClienteCheck { get; set; }
        //Variables de CONEXION
        public bool ExcelExtension { get; set; }
        public bool ExcelConnection { get; set; }     
        //Variables de INVALIDCHAR
        public bool CodigoClienteInvalidChar { get; set; }
        public bool ImporteInvalidChar { get; set; }
        public bool ConceptoInvalidChar { get; set; }
        public bool CodigoMarcaInvalidChar { get; set; }
        public bool TipoIvaInvalidChar { get; set; }
        public bool InvalidCharGlobal { get; set; }     
        //Variable GLOBAL
        public bool Validated { get; set; }
        public bool Importable { get; set; }   
    }
}