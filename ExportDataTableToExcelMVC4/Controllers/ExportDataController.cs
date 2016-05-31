using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExportDataTableToExcelInMVC4.Models;
using ClosedXML;
using ClosedXML.Excel;
using System.IO;

namespace ExportDataTableToExcelInMVC4.Controllers
{
    public class ExportDataController : Controller
    {
        public ActionResult Index()
        {
            String constring = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;//Conexion de la DB 
            SqlConnection con = new SqlConnection(constring);// de donde saldrán los datos para el excel
            string query = "select * From Facturas";//DB 
            DataTable dt = new DataTable();//se crea dataTable
            con.Open();//Se abre conexion
            SqlDataAdapter da = new SqlDataAdapter(query, con);//Se crea dataAdapter para ejecutar el comando "query" en la conexion "con"
            da.Fill(dt);//Se llena el dataTable con los datos recogidos por el dataAdapter
            con.Close();//Se cierra conexion

            IList<ExportDataTableToExcelModel> model = new List<ExportDataTableToExcelModel>();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                model.Add(new ExportDataTableToExcelModel()//Se exportan datos de DB al modelo del excel
                {
                    CodigoCliente = Convert.ToInt32(dt.Rows[i]["CodigoCliente"]),
                    Importe = Convert.ToInt32(dt.Rows[i]["Importe"]),
                    Concepto = dt.Rows[i]["Concepto"].ToString(),
                    TipoIva = dt.Rows[i]["TipoIva"].ToString(),
                    CodigoMarca = Convert.ToInt32(dt.Rows[i]["CodigoMarca"])
                });
            }
            return View(model);// se muestra la view de Exportar de nuevo
        }

        public ActionResult ExportData()
        {
            String constring = ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString;
            SqlConnection con = new SqlConnection(constring);
            string query = "select * From Facturas";
            DataTable dt = new DataTable();
            dt.TableName = "Facturas";//se crea dataTable
            con.Open();
            SqlDataAdapter da = new SqlDataAdapter(query, con);
            da.Fill(dt); //Se llena la tabla Facturas con los datos recogido por el dataAdpater con la conexion "constring" y la query
            con.Close();

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);//se llena el worksheet con los datos del dataTable
                wb.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;//estilos del workSheet
                wb.Style.Font.Bold = true;

                Response.Clear();//borra cualquier información de HTML tamponada
                Response.Buffer = true;//el servidor va a contener la respuesta al navegador hasta que todos los scripts del servidor
                Response.Charset = "";// han sido procesados ​​, o hasta que el script llama al método Flush o Fin
                //añade el nombre del charset a la cabecera de tipo de contenido en el objeto Response .
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename= FacturasReport.xlsx");//Se añaden los headers

                using (MemoryStream MyMemoryStream = new MemoryStream())//Datos para guardar el excel
                {
                    wb.SaveAs(MyMemoryStream);
                    MyMemoryStream.WriteTo(Response.OutputStream);
                    Response.Flush();//envía el buffered output inmediatamente. 
                    Response.End();//Este método produce un error en tiempo de ejecución si Response.Buffer no se ha establecido en TRUE
                }
            }
            return RedirectToAction("Index", "ExportData");
        }

        private void releaseObject(object obj)
        {
            try
            {//Se disminuye el recuento de referencia asociado con el objeto COM especificado ?¿ .
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}