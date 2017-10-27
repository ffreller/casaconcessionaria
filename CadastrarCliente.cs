using System;
using System.IO;
    
public class CadastrarCliente
{
    public void Cadastrarcliente()
    {
        Console.WriteLine("Cadastro");
        Console.WriteLine("Qual é seu nome?");
        string nome = Console.ReadLine();
        Console.WriteLine("Qual é seu e-mail?");
        string email = Console.ReadLine();
        Console.WriteLine("Qual é seu CPF/CNPJ?");
        string cpfecnpj = Console.ReadLine();
        Console.WriteLine("Qual sua cidade?");
        string cidade = Console.ReadLine();
        Console.WriteLine("Qual seu bairro?");
        string bairro = Console.ReadLine();
        Console.WriteLine("Qual sua rua? (com número)");
        string rua = Console.ReadLine();
        
        


        // StreamWriter cadastro = new StreamWriter ("Cadastro.txt", true);
        // FileInfo cabecalho = new FileInfo("Cadastro.txt");
        // if(cabecalho.Length == 0)
        // {
        //     cadastro.WriteLine ("NOME; E-MAIL; CPF/CNPJ; DATA E HORA");
        // }
        // cadastro.WriteLine(nome + ";" + email + ";" + cpfecnpj + ";" + DateTime.Now);
        // cadastro.Close();
        
    }
}          