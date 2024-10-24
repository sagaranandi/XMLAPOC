using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Diagnostics;

class Program
{
    static void Main()
    {
        string connectionString = "Data Source=powerbi://api.powerbi.com/v1.0/myorg/Infosuite - Dev;Initial Catalog=Infosuite Master Data Model";

        Console.WriteLine("Opening connection to Power BI dataset...");

        using (AdomdConnection connection = new AdomdConnection(connectionString))
        {
            try
            {
                connection.Open();
                Console.WriteLine("Connection successful!");

                // Start the stopwatch to measure execution time
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();

                string daxQuery = @"
                    EVALUATE
                    SUMMARIZECOLUMNS(
                        ROLLUPADDISSUBTOTAL('[Dim] Customer OPG'[Country], ""Country Total""),
                        '[Dim] Customer OPG'[CustomerKey],
                        ""Total Revenue"", [Total Revenue],
                        ""Total Orders"", [Total Orders],
                        ""Average Unit Price"", [Unit Price],
                        ""Total Cost"", [Total Cost],
                        ""Total Profit"", [Total Profit],
                        ""YTD Revenue"", [YTD Revenue],
                        ""Prev Month Revenue"", [Prev Month Revenue]
                    )";

                Console.WriteLine("Executing DAX query...");

                using (AdomdCommand command = new AdomdCommand(daxQuery, connection))
                {
                    using (AdomdDataReader reader = command.ExecuteReader())
                    {
                        int totalRows = 0;

                        while (reader.Read())
                        {
                            if (totalRows >= 15)
                            {
                                break; 
                            }

                            string country = reader.IsDBNull(0) ? "N/A" : reader.GetString(0);
                            long? customerKey = reader.IsDBNull(1) ? (long?)null : reader.GetInt64(1); // Change to Int64
                            decimal? totalRevenue = reader.IsDBNull(2) ? (decimal?)null : reader.GetDecimal(2);
                            long? totalOrders = reader.IsDBNull(3) ? (int?)null : reader.GetInt64(3);
                            decimal? averageUnitPrice = reader.IsDBNull(4) ? (decimal?)null : reader.GetDecimal(4);
                            decimal? totalCost = reader.IsDBNull(5) ? (decimal?)null : reader.GetDecimal(5);
                            decimal? totalProfit = reader.IsDBNull(6) ? (decimal?)null : reader.GetDecimal(6);
                            decimal? ytdRevenue = reader.IsDBNull(7) ? (decimal?)null : reader.GetDecimal(7);
                            decimal? prevMonthRevenue = reader.IsDBNull(8) ? (decimal?)null : reader.GetDecimal(8);

                            Console.WriteLine($@"
                                Country: {country}
                                Customer Key: {customerKey?.ToString() ?? "N/A"}
                                Total Revenue: {totalRevenue?.ToString("C2") ?? "N/A"}
                                Total Orders: {totalOrders?.ToString() ?? "N/A"}
                                Average Unit Price: {averageUnitPrice?.ToString("C2") ?? "N/A"}
                                Total Cost: {totalCost?.ToString("C2") ?? "N/A"}
                                Total Profit: {totalProfit?.ToString("C2") ?? "N/A"}
                                YTD Revenue: {ytdRevenue?.ToString("C2") ?? "N/A"}
                                Prev Month Revenue: {prevMonthRevenue?.ToString("C2") ?? "N/A"}
                            ");
                            totalRows++;
                        }

                        Console.WriteLine($"Total Rows Printed: {totalRows}");
                    }
                }

                stopwatch.Stop();
                TimeSpan elapsedTime = stopwatch.Elapsed;
                Console.WriteLine($"Total Time Taken: {elapsedTime.TotalSeconds} seconds");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
                Console.WriteLine("Please check if the measure names and columns exist in the dataset.");
            }
        }
        Console.WriteLine("Press any key to exit...");
        Console.ReadKey();
    }
}
