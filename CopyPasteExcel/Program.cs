using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using Microsoft.Extensions.Configuration;

namespace CopyPasteExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Starting the transfer, Please wait...");
                IConfigurationBuilder configuration = new ConfigurationBuilder()
                             .SetBasePath(Directory.GetCurrentDirectory())
                                .AddJsonFile($"appsettings.json");
                GetCSV(configuration);
            }
            catch (Exception e)
            {
                throw new("Could not connect to file appsettings.json...", e);
            }
        }
        private static string GetCSV(IConfigurationBuilder configuration)
        {
            try
            {
                var config = configuration.Build();
                var select = configuration.Build();
                var connectionString = config.GetConnectionString("DefaultConnection");
                string queryconn = select.GetValue<string>("Data:Query");
                using SqlConnection connection = new(connectionString);
                try
                {
                    connection.Open();
                    Console.WriteLine("Connected to Database... " + "\n" + connection.ConnectionString + "\n" + "Data Query= " + queryconn);
                    return CreateCSV(new SqlCommand(queryconn, connection).ExecuteReader(), configuration);
                }
                catch (Exception e)
                {
                    throw new("Could not connect to Database", e);
                }
                finally
                {
                    DateTime today = DateTime.Now;
                    Console.WriteLine("Application was run on..." + today + "\nExecute operation done...");
                    connection.Close();
                }
            }
            catch (Exception e)
            {
                throw new("Database path or request may be wrong...", e);
            }
        }
        private static string CreateCSV(SqlDataReader sqlDataReaderdf, IConfigurationBuilder configuration)
        {
            try
            {
                var path = configuration.Build();
                string file = path.GetValue<string>("FilePath:ExitPath");
                ICollection<string> lines = new List<string>();
                string headerLine = "";
                try
                {
                    if (sqlDataReaderdf.Read())
                    {
                        string[] columns = new string[sqlDataReaderdf.FieldCount];
                        for (int i = 0; i < sqlDataReaderdf.FieldCount; i++)
                        {
                            columns[i] = sqlDataReaderdf.GetName(i);
                        }
                        headerLine = string.Join(";", columns);
                        lines.Add(headerLine);
                    }
                    //data
                    while (sqlDataReaderdf.Read())
                    {
                        object[] values = new object[sqlDataReaderdf.FieldCount];
                        sqlDataReaderdf.GetValues(values);
                        lines.Add(string.Join(";", values));
                    }
                }
                catch (Exception e)
                {
                    throw new("Data processing failed...", e);
                }
                try
                {
                    //create file
                    System.IO.File.WriteAllLines(file, lines, Encoding.GetEncoding("iso-8859-9"));
                    Console.WriteLine("Transaction successful... " + "File Path= " + file);
                }
                catch (Exception e)
                {
                    throw new("File could not be created...", e);
                }
                return file;
            }
            catch (Exception e)
            {
                throw new("The file path may be incorrect...", e);
            }
        }
    }
}