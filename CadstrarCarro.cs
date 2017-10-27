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
            string arcon = Convert.ToString(opcionais1.arcon);
            string airbag = Convert.ToString(opcionais1.airbag);
            string abs = Convert.ToString(opcionais1.abs);
            
        if(!File.Exists(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\carros.xls"))
        {
            Criarexcel(modelocarro, anocarro, precocarro, opcionais1.arcon, opcionais1.airbag, opcionais1.abs);
        }
        else
        {
            //modificar arquivo, sem criar outro ou sobrepor
        }
    }
    public void Criarexcel(string modelocarro, string anocarro, string precocarro, string arcon, string airbag, string abs)
    {    
        Application ex = new Application();
        ex.Workbooks.Add();
        ex.Cells[1,1].Value = modelocarro;
        ex.Cells[1,2].Value = anocarro;
        ex.Cells[1,3].Value = precocarro;
        ex.Cells[1,4].Value = arcon;
        ex.Cells[1,5].Value = airbag;
        ex.Cells[1,6].Value = abs;

        ex.ActiveWorkbook.SaveAs(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\carros.xls");
        ex.Quit();
    }
            


            }   
    }
}    
        //     StreamWriter cadastroproduto = new StreamWriter ("Cadastroproduto.txt", true);
        //     cadastroproduto.Write(nomeproduto + ";" + codigoproduto + ";" + descricaoproduto + ";" + precoproduto + ";");
        //     cadastroproduto.Close();
        //     FileInfo cabecalho = new FileInfo("Cadastroproduto.txt");
        //     if(cabecalho.Length == 0)
        //     {
        //         cadastroproduto.WriteLine ("NOME DO PRODUTO; CÓDIGO DO PRODUO; DESCRIÇÃO DO PRODUTO; PREÇO;");
        //     }
        // }
        //     }