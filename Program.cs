using System;
using System.Data;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Diagnostics;



namespace Printer_Service
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string currentVersion = getCurrentVersion();
            Console.WriteLine($"Harithma Printer Service : Version {currentVersion}");

            autoUpdate();

            string rabbitQueue = "harithmaq";
            string rabbitUsername = Environment.GetEnvironmentVariable("RABBIT_MQ_USERNAME");
            string rabbitPassword = Environment.GetEnvironmentVariable("RABBIT_MQ_PASSWORD");
            string rabbitHost = Environment.GetEnvironmentVariable("RABBIT_MQ_HOST");

            var factory = new ConnectionFactory() { HostName = rabbitHost, UserName = rabbitUsername, Password = rabbitPassword };
            using (var connection = factory.CreateConnection())
            using (var channel = connection.CreateModel())
            {
                channel.QueueDeclare(queue: rabbitQueue, durable: false, exclusive: false, autoDelete: false, arguments: null);

                var consumer = new EventingBasicConsumer(channel);
                consumer.Received += (model, ea) =>
                {
                    var body = ea.Body.ToArray();
                    var message = System.Text.Encoding.UTF8.GetString(body);
                    Console.WriteLine("\nInvoice Submitted {0}", message);

                    dynamic jsonObject = JsonConvert.DeserializeObject(message);

                    if (jsonObject.invoice_type == 1)
                    {
                        printServiceReceipt(jsonObject);
                        Console.WriteLine("Invoice printed");
                    }
                    else if (jsonObject.invoice_type == 2)
                    {
                        printItemReceipt(jsonObject);
                        Console.WriteLine("Invoice printed");
                    }
                };
                channel.BasicConsume(queue: "harithmaq", autoAck: true, consumer: consumer);

                Console.WriteLine($"Host : {rabbitHost}");
                Console.WriteLine($"Queue : {rabbitQueue}.");
                Console.WriteLine("\nWaiting for print request... To exit press CTRL+C");
                Console.ReadLine(); // Keep the console application running until user closes it
            }
        }
        static void printItemReceipt(dynamic jsonObject)
        {
            Reports.ItemInvoiceReceipt itemInvoice = new Reports.ItemInvoiceReceipt();

            itemInvoice.SetDataSource(generateDataSet(jsonObject.invoice_details));
            itemInvoice.SetParameterValue("invoiceNumber", jsonObject.invoice_number);
            itemInvoice.SetParameterValue("invoiceTotalPrice", jsonObject.total_price);
            itemInvoice.SetParameterValue("invoiceDiscount", jsonObject.discount_pct);
            itemInvoice.SetParameterValue("invoicePaidAmount", jsonObject.paid_amount);
            itemInvoice.SetParameterValue("invoiceGrossPrice", jsonObject.gross_price);
            itemInvoice.SetParameterValue("invoicebalance", jsonObject.gross_price);

            itemInvoice.PrintOptions.PaperOrientation = CrystalDecisions.Shared.PaperOrientation.Portrait;
            itemInvoice.PrintToPrinter(1, false, 0, 0);
            //itemInvoice.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, $"Invoice_Service_{jsonObject.invoice_number}.pdf");

        }

        static void printServiceReceipt(dynamic jsonObject)
        {
            Reports.ServiceInvoiceReceipt serviceInvoice = new Reports.ServiceInvoiceReceipt();

            serviceInvoice.SetDataSource(generateDataSet(jsonObject.invoice_details));
            serviceInvoice.SetParameterValue("invoiceNumber", jsonObject.invoice_number);
            serviceInvoice.SetParameterValue("invoiceTotalPrice", jsonObject.total_price);
            serviceInvoice.SetParameterValue("invoiceDiscount", jsonObject.discount_pct);
            serviceInvoice.SetParameterValue("invoicePaidAmount", jsonObject.paid_amount);
            serviceInvoice.SetParameterValue("invoiceGrossPrice", jsonObject.gross_price);
            serviceInvoice.SetParameterValue("invoicebalance", jsonObject.gross_price);
            serviceInvoice.SetParameterValue("customerName", jsonObject.customer_name.ToString());
            serviceInvoice.SetParameterValue("employeeName", jsonObject.employee_name.ToString());
            serviceInvoice.SetParameterValue("vehicalNumber", jsonObject.vehical_number.ToString());
            serviceInvoice.SetParameterValue("washBay", jsonObject.wash_bay.ToString());
            serviceInvoice.SetParameterValue("currentMilage", jsonObject.current_milage);
            serviceInvoice.SetParameterValue("nextMilage", jsonObject.next_milage);

            serviceInvoice.PrintToPrinter(1, false, 0, 0);
            //serviceInvoice.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, $"Invoice_Service_{jsonObject.invoice_number}.pdf");
        }

        static dynamic generateDataSet(dynamic invoiceDetails)
        {
            DataSet InvoiceDetailsDS = new DataSet();
            DataTable InvoiceDetailsDT = new DataTable();

            InvoiceDetailsDT.Columns.Add("itemName", typeof(String));
            InvoiceDetailsDT.Columns.Add("itemUnitPrice", typeof(decimal));
            InvoiceDetailsDT.Columns.Add("itemQuantity", typeof(int));
            InvoiceDetailsDT.Columns.Add("itemTotalPrice", typeof(decimal));

            foreach (var item in invoiceDetails)
            {
                InvoiceDetailsDT.Rows.Add(item.item_name, item.unit_price, item.quantity, item.total_price);
            }
            InvoiceDetailsDS.Tables.Add(InvoiceDetailsDT);

            return InvoiceDetailsDS;
        }

        static void autoUpdate()
        {
            Console.WriteLine("Checking for updates...");

            string currentVersion = getCurrentVersion();
            if (string.IsNullOrEmpty(currentVersion))
            {
                Console.WriteLine("Failed to retrieve the current version. Update process aborted.");
                return;
            }

            try
            {
                using (WebClient webClient = new WebClient())
                {
                    string latestVersion = webClient.DownloadString("https://raw.githubusercontent.com/Madlhawa/Harithma-Printer-Service/refs/heads/master/Updates/Version.txt");

                    if (string.IsNullOrEmpty(latestVersion))
                    {
                        Console.WriteLine("Failed to retrieve the latest version. Update process aborted.");
                        return;
                    }

                    if (!currentVersion.Equals(latestVersion, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Newer version {latestVersion} found.");

                        string tempPath = Path.Combine(Path.GetTempPath(), "HarithmaPrinterServiceSetup.msi");
                        webClient.DownloadFile("https://github.com/Madlhawa/Harithma-Printer-Service/raw/refs/heads/master/Updates/HarithmaPrinterServiceSetup.msi", tempPath);

                        Console.WriteLine("Downloaded the update installer. Starting the installation...");

                        Process process = new Process
                        {
                            StartInfo = new ProcessStartInfo
                            {
                                FileName = "msiexec.exe",
                                Arguments = $"/i \"{tempPath}\"",
                                UseShellExecute = true
                            }
                        };
                        process.Start();

                        Console.WriteLine("Installation started. Exiting the application...");
                        Environment.Exit(0); // 0 indicates successful termination
                    }

                }
            }
            catch
            {

            }
        }

        static string getCurrentVersion()
        {
            try
            {
                // Read all text from the file and return it
                string filePath = Path.Combine(Directory.GetCurrentDirectory(), "Updates", "Version.txt");
                return File.ReadAllText(filePath);
            }
            catch (Exception ex)
            {
                // Handle any errors and return an error message
                Console.WriteLine($"Error reading file: {ex.Message}");
                return null; // or you can return an empty string if preferred
            }
        }
    }
}
