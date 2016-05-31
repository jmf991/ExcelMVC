using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;
using ExportDataTableToExcelInMVC4.Models;
using System.Configuration;
using System.IO;
using System.Data;
using System.Data.Entity;
using System.Data.Common;
using System.Data.OleDb;
using System.Data.SqlClient;

using System.Web.Mvc.Html;
using System.Text.RegularExpressions;

namespace ExportDataTableToExcelMVC4.Controllers
{
    public class Factura
    {
        public double CodigoCliente { get; set; }
        public double Importe { get; set; }
        public string Concepto { get; set; }
        public string TipoIva { get; set; }
        public double CodigoMarca { get; set; }
    }
    //static class DataRowExtensions
    //{
    //    public static object GetValue(this DataRow row, string column)
    //    {
    //        return row.Table.Columns.Contains(column) ? row[column] : null;
    //    }
    //}
  
    public class ImportDataController : Controller
    {
        public ActionResult Import()
        { return View(); }

        public ActionResult Importexcel()
        {
            //DEFINICIONES
            ImportarModel ImportarModelObj = new ImportarModel();//Se instancia la variable del model para utilizar sus valores estaticos
            //Listas de errores
            List<int> errorCodigoClienteList = new List<int>();//Se instancian las listas de errores 
            List<int> errorImporteList = new List<int>();      //para guardar los indices de celdas null
            List<int> errorConceptoList = new List<int>();
            List<int> errorTipoIvaList = new List<int>();
            List<int> errorCodigoMarcaList = new List<int>();
            List<int> errorCodigoClienteCheckListIndex = new List<int>();//Lista con indices de CodigoCliente no presentes en DB 
            List<object> errorCodigoClienteCheckList = new List<object>();//Lista con valores de CodigoCliente no presentes en DB
            //Listas de invalidChar
            List<int> invalidCharCodigoClienteList = new List<int>();//Se instancian las listas de invalidChar 
            List<int> invalidCharImporteList = new List<int>();      //para guardar los indices de celdas con invalidChar
            List<int> invalidCharConceptoList = new List<int>();
            List<int> invalidCharTipoIvaList = new List<int>();
            List<int> invalidCharCodigoMarcaList = new List<int>();
            //Lista con valores de CodigoCliente presentes en la DB
            List<object> codigoClienteDBList = new List<object>();

            //Loader de excel            
            string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/UploadedFolder"), Request.Files["FileUpload1"].FileName);
            if (System.IO.File.Exists(path1)) System.IO.File.Delete(path1);

            //CONNECTION STRINGS
            //Base de datos
            Request.Files["FileUpload1"].SaveAs(path1);
            string sqlConnectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=excelMvcDB;Integrated Security=True";
            //Excel
            string excelConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 +
                ";Extended Properties=Excel 12.0;Persist Security Info=False";
            OleDbConnection excelConnection = new OleDbConnection(excelConnectionString);
            //Comando SQL
            OleDbCommand cmd = new OleDbCommand("Select [CodigoCliente],[Importe],[Concepto],[TipoIva],[CodigoMarca]" +
                " from [Facturas$]", excelConnection);

            //INICIO CHECKERS
            //CHECKER CodigoCliente en DB: Se crea una lista con los "codigoCliente" disponibles en la DB
            using (SqlConnection conexionDB = new SqlConnection(sqlConnectionString))
            {
                SqlCommand getCodigoClienteDB = new SqlCommand("Select * from FacturasRecord", conexionDB);
                conexionDB.Open();
                using (IDataReader dataReader = getCodigoClienteDB.ExecuteReader())
                {
                    while (dataReader.Read())
                    {
                        ListItem CodigoClienteDB = new ListItem(Convert.ToString(dataReader["CodigoCliente"]));
                        codigoClienteDBList.Add(CodigoClienteDB);//Se añaden los "codigoCliente" de la DB a la lista
                        ViewBag.codigoClienteDBList = codigoClienteDBList;//Se alamacena la lista en viewBag xra su uso
                    }
                }
            }
            //Seleccion y Carga del archivo
            if (Request.Files["FileUpload1"].ContentLength > 0)
            {
                try
                {   //CHECKER FORMATO EXCEL  
                    string extension = System.IO.Path.GetExtension(Request.Files["FileUpload1"].FileName);
                    if (extension != ".xlsx" && extension != ".xls")//Si no es excel: ExcelExtension=false y PopUpExcelExtensionError
                    {
                        ImportarModelObj.ExcelExtension = false;
                        ImportarModelObj.SuccessGlobal = false;
                        ImportarModelObj.Validated = false;
                        return View("ImportPopUps", ImportarModelObj);
                    }
                    else { ImportarModelObj.ExcelExtension = true; }
                    //CHECKER CONEXION EXCEL
                    try
                    {//(La conexion depende del Comando SQL "cmd":Columnas, nombre de la hoja y "excelConnection" ) 
                        excelConnection.Open();
                        int a = cmd.ExecuteNonQuery();
                        excelConnection.Close();
                        ImportarModelObj.ExcelConnection = true;
                    }
                    catch
                    {//Si falla la conexion ExcelConnection=false y PopUpExcelConnectionError
                        excelConnection.Close();
                        ImportarModelObj.ExcelConnection = false;
                        ImportarModelObj.SuccessGlobal = false;
                        ImportarModelObj.Validated = false;
                        return View("ImportPopUps", ImportarModelObj);
                    }

                    //CHECKER FORMATO CELDAS
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);//no se abre conexion xq el datAdapter incluye el "cmd" 
                    DataSet ds = new DataSet(); //Se crea dataset               
                    ds.Tables.Add("xlsImport");//Se añaden tablas a ds
                    da.Fill(ds, "xlsImport");//Se añaden datos a ds con comando cmd del dataAdapter                    

                    //Se crean columnas en la tabla xlsImport
                    ds.Tables["xlsImport"].Columns[0].ColumnName = "CodigoCliente";
                    ds.Tables["xlsImport"].Columns[1].ColumnName = "Importe";
                    ds.Tables["xlsImport"].Columns[2].ColumnName = "Concepto";
                    ds.Tables["xlsImport"].Columns[3].ColumnName = "TipoIva";
                    ds.Tables["xlsImport"].Columns[4].ColumnName = "CodigoMarca";
                    excelConnection.Close();//Se cierra conexion con excel

                    //Una vez creada la tabla xlsImport se asignan los formatos de cada columna
                    var datos = ds.Tables["xlsImport"].AsEnumerable();
                    var query = datos.Where(x => x.Field<string>("Concepto") != null).Select(x =>
                                new Factura
                                {
                                    CodigoCliente = x.Field<double>("CodigoCliente"),
                                    Importe = x.Field<double>("Importe"),
                                    Concepto = x.Field<string>("Concepto"),
                                    TipoIva = x.Field<string>("TipoIva"),
                                    CodigoMarca = x.Field<double>("CodigoMarca")
                                });

                    foreach (DataRow row in ds.Tables["xlsImport"].Rows)
                    {
                        int index = ds.Tables["xlsImport"].Rows.IndexOf(row) + 2;//Se instancia el indice
                        object valueCodigoCliente = row["CodigoCliente"]; //Se instancian las celdas de las columnas seleccionadas
                        object valueImporte = row["Importe"];
                        object valueConcepto = row["Concepto"];
                        object valueTipoIva = row["TipoIva"];
                        object valueCodigoMarca = row["CodigoMarca"];

                        //CHEKER CODIGOCLIENTE DB-EXCEL
                        ListItem valueCodigoClienteStr = new ListItem(Convert.ToString(valueCodigoCliente));
                        if (!codigoClienteDBList.Contains(valueCodigoClienteStr))//Si no esta el valor en la DB:
                        {
                            errorCodigoClienteCheckList.Add(valueCodigoClienteStr);//lista con el valueCodigoCliente erroneo 
                            errorCodigoClienteCheckListIndex.Add(index);//lista con el index del valueCodigoCliente erroneo
                        }

                        //Se convierte el valor de la celda en string para quitar los carcateres invalidos
                        string valueCodigoClienteString = Convert.ToString(valueCodigoCliente);
                        string valueImporteString = Convert.ToString(valueImporte);
                        string valueConceptoString = Convert.ToString(valueConcepto);
                        string valueTipoIvaString = Convert.ToString(valueTipoIva);
                        string valueCodigoMarcaString = Convert.ToString(valueCodigoMarca);
                        //Se crean strings vacias para rellenarlas con la string sin invalidChars,y así compararlas con las string de origen
                        string valueCodigoClienteStringClean = ""; 
                        string valueImporteStringClean = "";
                        string valueConceptoStringClean = "";
                        string valueTipoIvaStringClean = "";
                        string valueCodigoMarcaStringClean = "";

                        //Se replazan los caracteres invalidos 
                        valueCodigoClienteStringClean = Regex.Replace(valueCodigoClienteString, "([^0-9a-nA-No-zO-ZçÇñÑáéíóúäëïöü_ ])", "").Trim();
                        valueImporteStringClean = Regex.Replace(valueImporteString, "([^0-9a-nA-No-zO-ZçÇñÑáéíóúäëïöü_ ])", "").Trim();
                        valueConceptoStringClean = Regex.Replace(valueConceptoString, "([^0-9a-nA-No-zO-ZçÇñÑáéíóúäëïöü_ ])", "").Trim();
                        valueTipoIvaStringClean = Regex.Replace(valueTipoIvaString, "([^0-9a-nA-No-zO-ZçÇñÑáéíóúäëïöü_ ])", "").Trim();
                        valueCodigoMarcaStringClean = Regex.Replace(valueCodigoMarcaString, "([^0-9a-nA-No-zO-ZçÇñÑáéíóúäëïöü_ ])", "").Trim();

                        //Si el valor anterior a la eliminacion de caracteres invalidos es el mismo que despues, no hay caracteres invalidos.
                        if (valueCodigoClienteString != valueCodigoClienteStringClean) { invalidCharCodigoClienteList.Add(index); }//CodigoCliente
                        if (valueImporteString != valueImporteStringClean) { invalidCharImporteList.Add(index); }//Importe
                        if (valueConceptoString != valueConceptoStringClean) { invalidCharConceptoList.Add(index); } //Concepto 
                        if (valueTipoIvaString != valueTipoIvaStringClean) { invalidCharTipoIvaList.Add(index); } //TipoIva
                        if (valueCodigoMarcaString != valueCodigoMarcaStringClean) { invalidCharCodigoMarcaList.Add(index); }//Importe
                                                                      
                        //FIN CHEKER CODIGOCLIENTE DB-EXCEL

                        //Si la celda esta en formato incorrecto saldrá null en la tabla "xlsImport"
                        //Si es null: añade el indice de la celda a su lista de errores  
                        if (valueCodigoCliente == DBNull.Value) { errorCodigoClienteList.Add(index); }  //CodigoCliente                       
                        if (valueImporte == DBNull.Value) { errorImporteList.Add(index); }              //Importe                        
                        if (valueConcepto == DBNull.Value) { errorConceptoList.Add(index); }            //Concepto                       
                        if (valueTipoIva == DBNull.Value) { errorTipoIvaList.Add(index); }              //TipoIva                       
                        if (valueCodigoMarca == DBNull.Value) { errorCodigoMarcaList.Add(index); }      //Importe
                        //FIN CHEKER FORMATO CELDAS
                    }

                    //Almacenamiento de listas de successErrors en ViewBag para utilizarse en las Views
                    ViewBag.errorCodigoClienteList = errorCodigoClienteList;                        //List-error-CodigoCliente
                    ViewBag.errorConceptoList = errorConceptoList;                                  //List-error-Concepto
                    ViewBag.errorCodigoMarcaList = errorCodigoMarcaList;                            //List-error-CodigoMarca
                    ViewBag.errorTipoIvaList = errorTipoIvaList;                                    //List-error-TipoIva
                    ViewBag.errorImporteList = errorImporteList;                                    //List-error-Importe
                    ViewBag.errorCodigoClienteCheckList = errorCodigoClienteCheckList;              //List-error-CodigoClienteCheck
                    ViewBag.errorCodigoClienteCheckListIndex = errorCodigoClienteCheckListIndex;    //List-error-CodigoClienteCheckIndex          
                    //Almacenamiento de listas de invalidChar
                    ViewBag.invalidCharCodigoClienteList = invalidCharCodigoClienteList;    //List-invalidChar-CodigoCliente
                    ViewBag.invalidCharCodigoClienteList = invalidCharImporteList;          //List-invalidChar-Importe
                    ViewBag.invalidCharConceptoList = invalidCharConceptoList;              //List-invalidChar-Concepto
                    ViewBag.invalidCharCodigoMarcaList = invalidCharCodigoMarcaList;        //List-invalidChar-CodigoMarca
                    ViewBag.invalidCharTipoIvaList = invalidCharTipoIvaList;                //List-invalidChar-TipoIva                    

                    //ASIGNACION VALORES PARCIALES
                    //CodigoClienteCheck
                    if (errorCodigoClienteCheckListIndex.Any())
                    { ImportarModelObj.SuccesCodigoClienteCheck = false; }                   //Validation-CodigoClienteCheck
                    else { ImportarModelObj.SuccesCodigoClienteCheck = true; }

                    //Success: Si las columnas NO tienen su lista de errores vacia: Success=False
                    if (errorCodigoClienteList.Any()) { ImportarModelObj.SuccessCodigoCliente = false; }
                    else { ImportarModelObj.SuccessCodigoCliente = true; }                          //Success-CodigoCliente
                    if (errorImporteList.Any()) { ImportarModelObj.SuccessImporte = false; }
                    else { ImportarModelObj.SuccessImporte = true; }                                //Success-Importe
                    if (errorCodigoMarcaList.Any()) { ImportarModelObj.SuccessCodigoMarca = false; }
                    else { ImportarModelObj.SuccessCodigoMarca = true; }                            //Success-CodigoMarca
                    if (errorTipoIvaList.Any()) { ImportarModelObj.SuccessTipoIva = false; }
                    else { ImportarModelObj.SuccessTipoIva = true; }                                //Success-TipoIva
                    if (errorConceptoList.Any()) { ImportarModelObj.SuccessConcepto = false; }
                    else { ImportarModelObj.SuccessConcepto = true; }                               //Success-Concepto                  

                    //InvalidChar: Si las columnas NO tienen su lista de invalidChar vacia: invalidChar=True
                    if (invalidCharCodigoClienteList.Any()) { ImportarModelObj.CodigoClienteInvalidChar = true; }
                    else { ImportarModelObj.CodigoClienteInvalidChar = false; }                          //invalidChar-CodigoCliente
                    if (invalidCharImporteList.Any()) { ImportarModelObj.ImporteInvalidChar = true; }
                    else { ImportarModelObj.CodigoClienteInvalidChar = false; }                          //invalidChar-Importe
                    if (invalidCharConceptoList.Any()) { ImportarModelObj.ConceptoInvalidChar = true; }
                    else { ImportarModelObj.CodigoClienteInvalidChar = false; }                          //invalidChar-Concepto
                    if (invalidCharCodigoMarcaList.Any()) { ImportarModelObj.CodigoMarcaInvalidChar = true; }
                    else { ImportarModelObj.CodigoClienteInvalidChar = false; }                          //invalidChar-CodigoMarca
                    if (invalidCharTipoIvaList.Any()) { ImportarModelObj.TipoIvaInvalidChar = true; }
                    else { ImportarModelObj.CodigoClienteInvalidChar = false; }                          //invalidChar-TipoIva

                    //ASIGNACION DE VALORES GLOBALES
                    //Validated: Si CodigoCliente presentes en DB + Conexion y Extension correctas
                    if (ImportarModelObj.SuccesCodigoClienteCheck && ImportarModelObj.ExcelExtension &&
                        ImportarModelObj.ExcelConnection)
                    { ImportarModelObj.Validated = true; }
                    else { ImportarModelObj.Validated = false; }                    //VALIDATED

                    //InavalidCharGlobal: Si TODAS las columnas tienen invalidChar=False: InvalidCharGlobal=false
                    if (!ImportarModelObj.CodigoClienteInvalidChar                  //invalidChar-CodigoCliente
                      && !ImportarModelObj.ImporteInvalidChar                       //invalidChar-Importe
                      && !ImportarModelObj.ConceptoInvalidChar                      //invalidChar-Concepto
                      && !ImportarModelObj.CodigoMarcaInvalidChar                   //invalidChar-CodigoMarca
                      && !ImportarModelObj.TipoIvaInvalidChar)                      //invalidChar-TipoIva                        
                    { ImportarModelObj.InvalidCharGlobal = false; }
                    else { ImportarModelObj.InvalidCharGlobal = true; }             //INVALIDCHARGLOBAL

                    //SuccessGlobal: Si TODAS las columnas tienen success=True: SuccessGlobal=True
                    if (ImportarModelObj.SuccessCodigoCliente                  //error-CodigoCliente       
                        && ImportarModelObj.SuccessImporte                     //error-Importe 
                        && ImportarModelObj.SuccessCodigoMarca                 //error-CodigoMarca
                        && ImportarModelObj.SuccessTipoIva                     //error-TipoIva
                        && ImportarModelObj.SuccessConcepto)                   //error-Concepto                      
                    { ImportarModelObj.SuccessGlobal = true; }
                    else { ImportarModelObj.SuccessGlobal = false; }            //SUCCESGLOBAL

                    //IMPORTABLE: si Validated=True + InvalidCharGlobal=False + SuccessGlobal=True :: Importable=True.
                    if (ImportarModelObj.Validated && !ImportarModelObj.InvalidCharGlobal && ImportarModelObj.SuccessGlobal)
                    { ImportarModelObj.Importable = true; }
                    else { ImportarModelObj.Importable = false; }

                    //IMPORTAR DATOS DE EXCEL A DB (Si Importable=True)
                    if (ImportarModelObj.Importable == true)
                    {
                        excelConnection.Open(); //Se abre conexion con el excel
                        OleDbDataReader dReader; //Se crea el dataReader
                        dReader = cmd.ExecuteReader(); //Se ejecuta el comando cmd con el dataReader       
                        SqlBulkCopy sqlBulk = new SqlBulkCopy(sqlConnectionString); //Se asocia conexion DB con sqlBulk
                        sqlBulk.DestinationTableName = "FacturasRecord"; //Se asocia tabla receptora en DB con sqlBulk
                        sqlBulk.WriteToServer(dReader); //Datos recogidos x el dataReader se vuelcan en la tabla de DB
                        excelConnection.Close(); //Se cierra conexion con el excel para poder utilizar el comando cmd mas adelante
                    }
                    //Tras asignar valor a todas las variables se muestran los PopUps Correspondientes
                    return View("ImportPopUps", ImportarModelObj);
                }
                catch //Si algun error no permite la ejecucion del try, SuccessGlobal=false
                {
                    excelConnection.Close();
                    ImportarModelObj.Importable = false;
                    return View("ImportPopUps", ImportarModelObj);
                }
            }
            //FIN CHECKERS
            return RedirectToAction("Import");
        }
    }
}
