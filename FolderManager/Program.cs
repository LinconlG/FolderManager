using System;
using System.Collections.Generic;
using System.IO;

namespace FolderManager
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine(" 1 - Criacao de pastas  |  2 - Gerenciador de Revisados");
                int v = Convert.ToInt32(Console.ReadLine());

                if (v != 1 || v!= 2)
                {
                    Console.WriteLine("Opcao invalida...");
                    Console.ReadLine();
                    return;
                }
                switch (v)
                {
                    case 1:
                        CriacaoPastas();
                        break;
                    case 2:
                        Console.WriteLine();
                        break;
                    default: break;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.ReadLine();
            }
        }

        static void CriacaoPastas()
        {
            Console.WriteLine("Pasta para criação: ");
            DirectoryInfo diretorio = new DirectoryInfo($@"{Console.ReadLine()}");
            Console.WriteLine();

            Console.WriteLine("Caminho da planilha: ");
            string diretorioExcel = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Quantidade de pastas: ");
            int linhas = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine();

            string[] nomesPastas = new string[linhas];
            bool invalido;

            var planilha = new Microsoft.Office.Interop.Excel.Application();
            var wb = planilha.Workbooks.Open($@"{diretorioExcel}", ReadOnly: true);
            var ws = wb.Worksheets[1];
            var r = ws.Range["A1"].Resize[linhas, 1];
            var array = r.Value;

            for (int i = 1; i < linhas; i++)
            {
                nomesPastas[i-1] = Convert.ToString(array[i, 1]);
                invalido = false;

                foreach (char c in nomesPastas[i-1])
                {
                    foreach (var item in Path.GetInvalidFileNameChars())
                    {
                        if (c == item)
                        {
                            Console.WriteLine();
                            Console.WriteLine($"-- O nome na linha {i} contem caractere invalido, por favor verifique...");
                            Console.WriteLine();
                            invalido = true;
                            break;
                        }
                    }
                    if (invalido)
                        break;
                }

                if (!(invalido) && !(Directory.Exists($@"{diretorio}\{nomesPastas[i - 1]}")))
                {
                    DirectoryInfo dir = Directory.CreateDirectory($@"{diretorio}\{nomesPastas[i - 1]}");
                }
            }
            wb.Close();
            planilha.Quit();

            Console.WriteLine("Finalizado!");
            Console.ReadLine();
        }
        static void Gerenciador()
        {
            Console.WriteLine("Pasta dos arquivos revisados: ");
            DirectoryInfo revisados = new DirectoryInfo($@"{Console.ReadLine()}");
            Console.WriteLine();

            Console.WriteLine("Caminho da planilha da relacao arquivo e tanque: ");
            string diretorioExcel = Console.ReadLine();
            Console.WriteLine();

            Console.WriteLine("Quantidade de arquivos: ");
            int linhas = Convert.ToInt32(Console.ReadLine());
            Console.WriteLine();

            Console.WriteLine("Pasta dos arquivos superados: ");
            DirectoryInfo superados = new DirectoryInfo($@"{Console.ReadLine()}");
            Console.WriteLine();

            var planilha = new Microsoft.Office.Interop.Excel.Application();
            var wb = planilha.Workbooks.Open($@"{diretorioExcel}", ReadOnly: true);
            var ws = wb.Worksheets[1];
            var r = ws.Range["A1"].Resize[linhas, 2];
            var array = r.Value;

            //char a 
            //a++ ??
        }

    }
}
