using System;
using System.IO;
using NetOffice.ExcelApi;

namespace casaconcessionaria{
    
    public class VenderCarro
    {
        public void Vendercarro()
        {
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\40809588897\Desktop\Programar\Semana 4\casaconcessionaria\carros.xls");
            int contador = 1;
            Console.WriteLine("Carros Disponíveis:");
            do
            {
                Console.WriteLine(ex.Cells[contador,1].Value.ToString() + "; " + ex.Cells[contador,2].Value.ToString() + "; " + ex.Cells[contador,3].Value.ToString());
                string opcional1 = "";
                string opcional2 = "";
                string opcional3 = "";
                if (ex.Cells[contador,4].Value.ToString() == "s")
                {
                    opcional1 = "Ar-condicionado";
                }
                else{opcional1 = "Sem Ar-condicionado";}
                if (ex.Cells[contador,5].Value.ToString() == "s")
                {
                    opcional2 = "Airbag";
                }
                else{opcional1 = "Sem Airbag";}
                
                if (ex.Cells[contador,4].Value.ToString() == "s")
                {
                    opcional3 = "ABS";
                }
                else{opcional1 = "Sem freios ABS";}
                Console.WriteLine(opcional1 + "; " + opcional2 + "; " + opcional3 + ".");
                contador += 1;
                
            } while (ex.Cells[contador,1].Value != null);
            Console.WriteLine("Digite o nome do carro que deseja");
            string carroescolhido = Console.ReadLine();

            do
            {

            }
            while (carrosescolhido ex.Cells[contador,1].Value)

            Console.WriteLine("Como deseja pagar? (digite 1 para a vista com 5% de desconto e 2 para a prazo)");
            string vistaprazo = Console.ReadLine();
            do
            {
                
                if (vistaprazo == "1")
                {
                    double preco = Convert.ToDouble(ex.Cells[contador,3]);
                    preco = preco * 95/100;
                }
                else if (vistaprazo == "2")
                {
                    Console.WriteLine("Em quantas parcelas deseja pagar?");
                    double parcelas;
                    do
                    {
                        Console.WriteLine("2, 4 ou 8 parcelas?");
                        parcelas = Convert.ToDouble(Console.ReadLine());
                        Console.WriteLine("O preço fica:");
                        double preco = Convert.ToDouble(ex.Cells[contador,3]);
                        preco = preco/parcelas; 
                        Console.WriteLine(parcelas + "parcelas de" + preco);

                    }
                    while (parcelas != 2 && parcelas != 4 && parcelas != 8);
                }    
            }
            while (vistaprazo != "1" && vistaprazo != "2");
            {

            }
            
        }
    }
}
    