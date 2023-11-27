using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace TIBR_WORKER
{
    public class TIBR_Service
    {
        public readonly ILogger<TIBR_Service> _logger;
        public TIBR_Service(ILogger<TIBR_Service> logger)
        {
            _logger = logger;
        }
        public void ProcessTIBRData()
        {
            // setting URL and connectionstring 
            string baseUrl = "https://nbg.gov.ge/gw/api/ct/MonetaryPolicy/TIBR/Index/Export/Excel/?";
            DateTime currentDate = DateTime.Now.Date;
            DateTime startDate = DateTime.Now.Date.AddDays(-4);
            DateTime endDate = currentDate.AddDays(1);
            string apiUrl = $"{baseUrl}end={endDate.ToString("yyyy-MM-dd")}T08%3A52%3A51.848Z&start={startDate.ToString("yyyy-MM-dd")}T20%3A00%3A00.000Z";
            string connectionString = "Server=(localdb)\\mssqllocaldb;Database=MyDataBase;";

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            // saving excel file and getting data from it
            using (WebClient client = new WebClient())
            {
                byte[] excelData = client.DownloadData(apiUrl);

                using (MemoryStream excelStream = new MemoryStream(excelData))
                {
                    using (ExcelPackage package = new ExcelPackage(excelStream))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                        int rowCount = worksheet.Dimension.Rows;
                        int colCount = worksheet.Dimension.Columns;

                        // making sql connection 
                        using (SqlConnection connection = new SqlConnection(connectionString))
                        {
                            connection.Open();
                            using (SqlCommand enableIdentityInsert = new SqlCommand("SET IDENTITY_INSERT TIBR_RATE ON", connection))
                            {
                                enableIdentityInsert.ExecuteNonQuery();
                            }

                            // executing query, converting and adding data into database table
                            for (int row = 2; row <= rowCount; row++)
                            {
                               
                                SqlCommand command = new SqlCommand("INSERT INTO TIBR_RATE (Id, Date, TIBR, TIBR1M, TIBR3M, TIBR6M, AddDateTime) VALUES (@Value1, @Value2, @Value3, @Value4, @Value5, @Value6, @Value7)", connection);
                                int id = Convert.ToInt32(worksheet.Cells[row, 1].Value);
                                string dateString = worksheet.Cells[row, 2].Value?.ToString();

                                //var arr = dateString.Split("/");
                                // var date = new DateTime(int.Parse(arr[0]), int.Parse( arr[1]), int.Parse(arr[2]));
                                DateTime date;
                                if (DateTime.TryParseExact(dateString, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
                                {
                                    command.Parameters.AddWithValue("@Value2", date);
                                }
                                decimal tibr = Convert.ToDecimal(worksheet.Cells[row, 3].Value);
                                decimal tibr1M = Convert.ToDecimal(worksheet.Cells[row, 4].Value);
                                decimal tibr3M = Convert.ToDecimal(worksheet.Cells[row, 5].Value);
                                decimal tibr6M = Convert.ToDecimal(worksheet.Cells[row, 6].Value);
                                string AddDateTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

                                command.Parameters.AddWithValue("@Value1", id);
                                command.Parameters.AddWithValue("@Value3", tibr);
                                command.Parameters.AddWithValue("@Value4", tibr1M);
                                command.Parameters.AddWithValue("@Value5", tibr3M);
                                command.Parameters.AddWithValue("@Value6", tibr6M);
                                command.Parameters.AddWithValue("@Value7", AddDateTime);

                                // error handling
                                try
                                {
                                    command.ExecuteNonQuery();
                                    _logger.LogInformation("Data inserted successfully for ID: {id}", id);
                                }
                                catch (SqlException ex)
                                {
                                    // catching if there is duplications and adding next record, if there is duplicate data
                                    if (ex.Number == 2627 || ex.Number == 2601)
                                    {
                                        _logger.LogWarning("Duplicate entry found for ID: {id}, skipping...", id);
                                        continue;                                      
                                    }
                                    else
                                    {
                                        _logger.LogError(ex, $"Error inserting data for ID: {id}. Error message: {ex.Message}");
                                        
                                    }
                                }

                            }
                        }
                    }
                }
            }
        }
    }
}
