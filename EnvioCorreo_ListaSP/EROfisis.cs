using EnvioCorreo_ListaSP.DBContext;
using EnvioCorreo_ListaSP.DTO;
using EnvioCorreo_ListaSP.Helpers;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;

namespace EnvioCorreo_ListaSP
{
    public class EROfisis
    {
        public void EROfisisPendientes()
        {
            //Cambio se sube a GitHub
            InterfacesDBContext InterfacesDB = new InterfacesDBContext();
            try
            {
                bool envioCorreo = Convert.ToBoolean(ConfigurationManager.AppSettings["envioCorreo"].ToString());
                Logger.WriteLine("Inicia Llamado SP USP_LISTA_ER");
                var listaEROfisisDTO = InterfacesDB.Set<EROfisisDTO>().FromSqlRaw($"exec USP_LISTA_ER").ToList();

                //Inicio Crear archivo excel
                var stream = new MemoryStream();
                using (var package = new ExcelPackage(stream))
                {
                    var workSheet = package.Workbook.Worksheets.Add("ListadoER");
                    workSheet.Cells.LoadFromCollection(listaEROfisisDTO, true);
                    package.Save();
                }
                stream.Position = 0;
                var archivoExcel = stream.ToArray();
                string excelName = string.Format("EntregasRendir-{0}.xlsx", DateTime.Now.ToString("yyyyMMddHHmmssfff"));
                //Fin Crear archivo excel

                Correo correo = new Correo();
                
                correo.Adjuntos.Add(new Adjunto
                {
                    archivo = archivoExcel,
                    nombreArchivo = excelName
                });

                try
                {
                    correo.Asunto = $"Entregar a Rendir Pendientes";
                    string cuerpo = "Existen Entregas a Rendir Pendientes, se adjunta Excel.";

                    Helpers.Helper.ConstruirCorreoError(correo, cuerpo);
                    Helpers.Helper.EnviarCorreoElectronico(correo, true);                    
                }
                catch (Exception ex)
                {
                    Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
                }
            }
            catch (Exception ex)
            {
                Logger.WriteLine($"{ex.Message}\n{ex.InnerException?.Message}");
            }
            finally
            {
                Logger.WriteLine("Fin Ejecución");
            }
        }
    }
}
