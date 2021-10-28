using Bgr.Base.Contracts;
using Bgr.Base.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Tests
{
    public class TestExportarExcel
    {
        public string ObtenerPath(string nameFile) => AppDomain.CurrentDomain.BaseDirectory + nameFile;

        [Test]
        public void TestExportar()
        {
            IExportExcel exportarExcel = new ExportExcel();
            exportarExcel.Log = TestContext.WriteLine;
            var paises = new List<Pais> { new Pais { Id = 1, Name = "Colombia" }, new Pais { Id = 2, Name = "Uruguay" } };

            
            exportarExcel.AddData(paises, "Paises");
            exportarExcel.ExportFile(ObtenerPath("Datos.xlsx"));
            var bytes = exportarExcel.ExportBytes();
            var stream = exportarExcel.Export();
            Assert.NotNull(stream);
        }
        [Test]
        public void TestImportar() 
        {
            var importar = new ImportExcel();
            var data = importar.Import(@".\ArchivosInicializacion\CodigosUnspsc.xlsx");
            Assert.Greater(data.Rows.Count, 0);
        }

        [Test]
        public void TestExportarClear()
        {
            try
            {
                IExportExcel exportarExcel = new ExportExcel();
                var stream = exportarExcel.Export();
            }
            catch (NoDataFoundException ex)
            {
                Assert.Pass("Generó excepctión por no tener datos" + ex.Message);
            }

        }
    }

    public class Contacto
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
    public class Pais
    {
        public int Id { get; set; }
        [ColumExcel("Nombre")]
        public string Name { get; set; }
    }
}