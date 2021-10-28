using ByA.Base.Excel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Tests
{
    public class TestImportarExcel
    {
        //public string ObtenerPath(string nameFile) => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, nameFile);
        [Test]
        public void TestImportar() 
        {
            var importar = new ImportarExcel();
            var data = importar.Importar(@".\ArchivosInicializacion\CodigosUnspsc.xlsx");
            Assert.Greater(data.Rows.Count, 0);
        }

        
    }
    
}