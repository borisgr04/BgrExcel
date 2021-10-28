using ByA.Base.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace Tests
{
    public class TestExportarExcel
    {
        [Test]
        public void TestExportar()
        {
            IExportarExcel exportarExcel = new ExportarExcel();
            exportarExcel.Log = TestContext.WriteLine;
            exportarExcel.AddData(new List<Contacto> { new Contacto { Id = 1, Name = "Boris" }, new Contacto { Id = 2, Name = "Boris" } }, "Contactos");
            var paises = new List<Pais> { new Pais { Id = 1, Name = "Colombia" }, new Pais { Id = 2, Name = "Uruguay" } };
            exportarExcel.AddData(paises, "Paises");
            exportarExcel.ExportarFile(@".\export.xls");
            //var bytes = exportarExcel.ExportarBytes();
            //var stream = exportarExcel.Exportar();
            Assert.Pass("Se exportó el archivo");
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
        [ColumExcel("Nombre Pais")]
        public string Name { get; set; }
    }


}
