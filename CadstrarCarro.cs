using System;
using System.IO;
using NetOffice.ExcelApi;

namespace casaconcessionária{
    
    public class CadastrarCarro{
    
        public void Cadastrarcarro()
        
            {Console.WriteLine("Qual o modelo do carro?");
            string modelocarro = Console.ReadLine();
            Console.WriteLine("Qual o ano do carro?");
            string anocarro = Console.ReadLine();
            Console.WriteLine("Qual o preço do carro(sem opcionais)");
            string precocarro = Console.ReadLine();
            Console.WriteLine("Opcionais");
            Console.WriteLine("Ar-condicionado? (s ou n)");
            Opcionais opcionais1 = new Opcionais();
            opcionais1.arcon = Console.ReadLine();
            Console.WriteLine("Airbag? (s ou n)");
            opcionais1.airbag = Console.ReadLine();
            Console.WriteLine("Freios ABS? (s ou n)");
            opcionais1.abs = Console.ReadLine();
            
            
        if(!File.Exists(@"carros.xls"))
        {
            Criarexcel(modelocarro, anocarro, precocarro, opcionais1);
        }
        else
        {
            Application ex = new Application();
            ex.Workbooks.Open(@"clientes.xls");
            int contador = 1;
            do
            {
                contador += 1;

            } while (ex.Cells[contador,1] != null);
            
            ex.Workbooks.Add();
            ex.Cells[contador,1].Value = modelocarro;
            ex.Cells[contador,2].Value = anocarro;
            ex.Cells[contador,3].Value = precocarro;
            ex.Cells[contador,4].Value = opcionais1.arcon;
            ex.Cells[contador,5].Value = opcionais1.airbag;
            ex.Cells[contador,6].Value = opcionais1.abs;
            ex.Quit();
        }
    }
    public void Criarexcel(string modelocarro, string anocarro, string precocarro,  Opcionais opcionais1)
    {    
        Application ex = new Application();
        ex.Workbooks.Add();
        ex.Cells[1,1].Value = modelocarro;
        ex.Cells[1,2].Value = anocarro;
        ex.Cells[1,3].Value = precocarro;
        ex.Cells[1,4].Value = opcionais1.arcon;
        ex.Cells[1,5].Value = opcionais1.airbag;
        ex.Cells[1,6].Value = opcionais1.abs;

        ex.ActiveWorkbook.SaveAs(@"carros.xls");
        ex.Quit();
    }
            


            }   
    }
