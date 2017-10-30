using System;
using System.IO;
using NetOffice.ExcelApi;

namespace casaconcessionaria{
    
    public class VenderCarro
    {
        public void Vendercarro()
        {
            Application ax = new Application();
            int contador1 = 1;
            int contador = 1;
            Console.WriteLine("Digite seu CPF/CNPJ");
            string cpfcnpj = Console.ReadLine();
            ax.Workbooks.Open(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\clientes.xls");
            do
            {
                contador1 += 1;   
            }
            while (ax.Cells[contador1,3].Value.ToString() != cpfcnpj);
            
            // nao pode travar (escolher um cpf que nao esta na lista)
            if(ax.Cells[contador1,3].Value.ToString() == cpfcnpj)
            {         
            Console.WriteLine("Carros Disponíveis:");
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\carros.xls");
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
            contador = 1;
            do
            {
                contador += 1;   
            }
            while (ex.Cells[contador,1].Value.ToString() != carroescolhido);
            // nao pode travar (escolher um carro que nao esta na lista)
            ex.Cells[contador,8].Value = "vendido";
            Console.WriteLine("Você escolheu o carro: " + carroescolhido);

            Console.WriteLine("Como deseja pagar? (digite 1 para a vista com 5% de desconto e 2 para a prazo)");
            string vistaprazo = Console.ReadLine();
            do
            {
                
                if (vistaprazo == "1")
                {
                    double preco = Convert.ToDouble(ex.Cells[contador,3].Value);
                    preco = preco * 95/100;
                    Console.WriteLine("O preço fica " + preco);
                }
                else if (vistaprazo == "2")
                {
                    Console.WriteLine("Em quantas parcelas deseja pagar?");
                    double parcelas;
                    do
                    {
                        Console.WriteLine("2, 4 ou 8 parcelas?");
                        parcelas = Convert.ToDouble(Console.ReadLine());
                        Console.WriteLine("O preço fica: ");
                        double preco = Convert.ToDouble(ex.Cells[contador,3].Value);
                        preco = preco/parcelas; 
                        Console.WriteLine(parcelas + " parcelas de " + preco + " reais");

                    }
                    while (parcelas != 2 && parcelas != 4 && parcelas != 8);
                }    
            }
            while (vistaprazo != "1" && vistaprazo != "2");
            string cl1 = ax.Cells[contador1,1].Value.ToString();
            string cl2 = ax.Cells[contador1,2].Value.ToString();
            string cl3 = ax.Cells[contador1,3].Value.ToString();
            string cl4 = ax.Cells[contador1,4].Value.ToString();
            string cl5 = ax.Cells[contador1,5].Value.ToString();
            string cl6 = ax.Cells[contador1,6].Value.ToString();
            string cr1 = ex.Cells[contador,1].Value.ToString();
            string cr2= ex.Cells[contador,2].Value.ToString();
            string cr3 = ex.Cells[contador,3].Value.ToString();
            string cr4 = ex.Cells[contador,4].Value.ToString();
            string cr5 = ex.Cells[contador,5].Value.ToString();
            string cr6 = ex.Cells[contador,6].Value.ToString();
        
        
        
        if(!File.Exists(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\vendas.xls"))
        {
            Criarexcelvenda(cl1, cl2, cl3, cl4, cl5, cl6, cr1, cr2, cr3, cr4, cr5, cr6);
            ax.Quit();
            ax.Dispose();
            ex.Quit();
            ex.Dispose();
        }
        else
        {
            Application ox = new Application();
            ox.DisplayAlerts = false;
            ox.Workbooks.Open(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\vendas.xls");
            int contador3 = 1;
            do
            {
                contador3 += 1;

            } while (ox.Cells[contador3,1].Value != null);
            
            
            ox.Cells[contador3,1].Value = ax.Cells[contador1,1].Value;
            ox.Cells[contador3,2].Value = ax.Cells[contador1,2].Value;
            ox.Cells[contador3,3].Value = ax.Cells[contador1,3].Value;
            ox.Cells[contador3,4].Value = ax.Cells[contador1,4].Value;
            ox.Cells[contador3,5].Value = ax.Cells[contador1,5].Value;
            ox.Cells[contador3,6].Value = ax.Cells[contador1,6].Value;
            ox.Cells[contador3,7].Value = ex.Cells[contador,1].Value;
            ox.Cells[contador3,8].Value = ex.Cells[contador,2].Value;
            ox.Cells[contador3,9].Value = ex.Cells[contador,3].Value;
            ox.Cells[contador3,10].Value = ex.Cells[contador,4].Value;
            ox.Cells[contador3,11].Value = ex.Cells[contador,5].Value;
            ox.Cells[contador3,12].Value = ex.Cells[contador,6].Value;
        
            ox.ActiveWorkbook.Save();
            ox.Quit();
            ox.Dispose();
            ax.Quit();
            ax.Dispose();
            ex.Quit();
            ex.Dispose();
        }
            }
        }
    public void Criarexcelvenda(string cl1, string cl2, string cl3, string cl4, string cl5, string cl6, string cr1, string cr2, string cr3, string cr4, string cr5, string cr6)
    {    
        Application ix = new Application();
        ix.Workbooks.Add();
        ix.Cells[1,1].Value = cl1;
        ix.Cells[1,2].Value = cl2;
        ix.Cells[1,3].Value = cl3;
        ix.Cells[1,4].Value = cl4;
        ix.Cells[1,5].Value = cl5;
        ix.Cells[1,6].Value = cl6;
        ix.Cells[1,7].Value = cr1;
        ix.Cells[1,8].Value = cr2;
        ix.Cells[1,9].Value = cr3;
        ix.Cells[1,10].Value = cr4;
        ix.Cells[1,11].Value = cr5;
        ix.Cells[1,12].Value = cr6;
        
        ix.ActiveWorkbook.SaveAs(@"C:\Users\Fabio Freller\Documents\Programar\casaconcessionaria\vendas.xls");
        ix.Quit();
        ix.Dispose();
    }
            
            
        }
    }
    