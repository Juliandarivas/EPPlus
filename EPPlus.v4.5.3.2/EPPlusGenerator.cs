using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace EPPlus.v4._5._3._2
{
    public class EPPlusGenerator
    {
        public byte[] Generate<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos")
        {
            Stream stream = new MemoryStream(recurso);
            return GenerateExcelByteArray(stream, datos, nombreHojaDatos);
        }


        private static byte[] GenerateExcelByteArray<T>(Stream excelPlantilla, IEnumerable<T> datos, string nombreHojaDatos)
        {
            var excel = new ExcelPackage(excelPlantilla);
            var sheet = excel.Workbook.Worksheets[nombreHojaDatos];

            LoadData(datos, sheet);
            ConvertPropertiesIntoColumns<T>(sheet);

            return excel.GetAsByteArray();
        }

        private static void LoadData<T>(IEnumerable<T> datos, ExcelWorksheet sheet)
        {
            sheet.InsertRow(2, datos.Count() - 1);
            sheet.Cells["A2"].LoadFromCollection(datos);
        }

        private static void ConvertPropertiesIntoColumns<T>(ExcelWorksheet sheet)
        {
            var indiceColumna = 1;
            foreach (var propiedad in typeof(T).GetProperties())
            {
                var prop = propiedad.GetCustomAttribute<DisplayNameAttribute>();

                if (propiedad.PropertyType == typeof(DateTime))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = "dd/mm/yyyy";

                if (propiedad.PropertyType == typeof(decimal))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = "#,##0.00";

                sheet.SetValue(1, indiceColumna++, prop?.DisplayName ?? propiedad.Name);
            }

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            CrearPropiedadesAdicionales<T>(sheet);
        }

        private static void CrearPropiedadesAdicionales<T>(ExcelWorksheet sheet)
        {
            //ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            //string tableName = "TablaGenerada";
            //ExcelTable table = sheet.Tables.Add(range, tableName);
            //table.TableStyle = TableStyles.Medium2;

            //string referencia = "\"Referencia\"";
            //string descripcion = "\"Descripción\"";
            //string separador = "\" - \"";

            //var cell = sheet.Cells[2, 6];
            //cell.Formula = $"=CONCATENATE({referencia},{separador},{descripcion})";
        }
    }
}
