using System;
using System.IO;
using NetOffice.ExcelApi;

namespace casaconcessionaria{
    
    public class ListarCarros
    {
        public void Listarcarros()
        {
        Application ex = new Application();
        ex.Workbooks.Open(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\vendas.xls");
        Console.WriteLine("Dados do Cliente:");
        Console.WriteLine(ex.Cells[1,1].Value + "; " + ex.Cells[1,2].Value + "; " + ex.Cells[1,3].Value + "; " + ex.Cells[1,4].Value + "; " + ex.Cells[1,5].Value + "; " + ex.Cells[1,6].Value);
        Console.WriteLine("Dados do Carro:");
        Console.WriteLine(ex.Cells[1,7].Value + "; " + ex.Cells[1,8].Value + "; " + ex.Cells[1,9].Value + "; Ar-condicionado: " + ex.Cells[1,10].Value + "; Airbag: " + ex.Cells[1,11].Value + "; ABS:" + ex.Cells[1,12].Value);
        ex.Quit();
        ex.Dispose();
        }
    }
}