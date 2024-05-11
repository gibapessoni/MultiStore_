using System;
using System.Data;
using System.Data.SqlClient;
using OfficeOpenXml;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string filePath = @"C:\Caminho\para\o\arquivo.xlsx";        

        string processedFolderPath = @"C:\Caminho\para\a\pasta\arquivos\processados";

        string connectionString = "Data Source=SeuServidor;Initial Catalog=DW_MultiStore;Integrated Security=True";

        try
        {
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    // Excluir dados anteriores da tabela Stage.MultiStore
                    using (SqlCommand deleteCommand = new SqlCommand("TRUNCATE TABLE Stage.MultiStore", connection))
                    {
                        deleteCommand.ExecuteNonQuery();
                    }

                    // Inserir novos dados na tabela Stage.MultiStore
                    using (SqlCommand insertCommand = new SqlCommand("INSERT INTO Stage.MultiStore VALUES (@VendaID, @DataVenda, @ProdutoID, @Quantidade, @ValorUnitario, @TotalVenda, @LojaID)", connection))
                    {
                        for (int row = 2; row <= rowCount; row++)
                        {
                            insertCommand.Parameters.Clear();
                            insertCommand.Parameters.AddWithValue("@VendaID", worksheet.Cells[row, 1].Value);
                            insertCommand.Parameters.AddWithValue("@DataVenda", worksheet.Cells[row, 2].Value);
                            insertCommand.Parameters.AddWithValue("@ProdutoID", worksheet.Cells[row, 3].Value);
                            insertCommand.Parameters.AddWithValue("@Quantidade", worksheet.Cells[row, 4].Value);
                            insertCommand.Parameters.AddWithValue("@ValorUnitario", worksheet.Cells[row, 5].Value);
                            insertCommand.Parameters.AddWithValue("@TotalVenda", worksheet.Cells[row, 6].Value);
                            insertCommand.Parameters.AddWithValue("@LojaID", worksheet.Cells[row, 7].Value);

                            insertCommand.ExecuteNonQuery();
                        }
                    }

                    // Mover o arquivo para a pasta de arquivos processados
                    string fileName = Path.GetFileName(filePath);
                    string processedFilePath = Path.Combine(processedFolderPath, fileName);
                    file.MoveTo(processedFilePath);

                    connection.Close();
                }
            }

            Console.WriteLine("Importação concluída com sucesso.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Erro: " + ex.Message);
        }
    }
}
