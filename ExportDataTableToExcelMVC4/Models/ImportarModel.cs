using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExportDataTableToExcelInMVC4.Models
{
    public class ImportarModel
    {        
        //Variables de VALIDATION
        public bool ExcelExtension { get; set; }
        public bool ExcelConnection { get; set; }        
        public bool SuccesCodigoClienteCheck { get; set; }
      
        //Variables de SUCCESS
        public bool SuccessCodigoCliente { get; set; }
        public bool SuccessImporte { get; set; }
        public bool SuccessConcepto { get; set; }
        public bool SuccessCodigoMarca { get; set; }
        public bool SuccessTipoIva { get; set; }
         
        //Variables de INVALIDCHAR
        public bool CodigoClienteInvalidChar { get; set; }
        public bool ImporteInvalidChar { get; set; }
        public bool ConceptoInvalidChar { get; set; }
        public bool CodigoMarcaInvalidChar { get; set; }
        public bool TipoIvaInvalidChar { get; set; }
      
        //Variable GLOBAL 
        public bool Validated { get; set; }         //Principal-Validation
        public bool SuccessGlobal { get; set; }     //Principal-Success
        public bool InvalidCharGlobal { get; set; } //Principal-InvalidChar
        public bool Importable { get; set; }        //Principal-Global   
    }
}