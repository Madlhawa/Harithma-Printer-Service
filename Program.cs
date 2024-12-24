using System;
using System.Data;
using RabbitMQ.Client;
using RabbitMQ.Client.Events;
using Newtonsoft.Json;
using System.IO;
using System.IO.Compression;
using System.Net;
using System.Diagnostics;



namespace Printer_Service
{
    internal class Program
    {
        static void Main(string[] args)
        {
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

                Console.WriteLine("Initialized Harithma Printer Service : Version 1.0.0.");
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
            WebClient webClient = new WebClient();
            var client = new WebClient();

            if (!webClient.DownloadString("https://raw.githubusercontent.com/Madlhawa/Harithma-Printer-Service/refs/heads/master/Updates/Version.txt").Contains("1.0.0"))
            {
                Console.WriteLine("New Version Found. Please type 'y' to update.");
                Console.ReadLine();

                try
                {
                    if (File.Exists(@".\MyAppSetup.msi")) { File.Delete(@".\MyAppSetup.msi"); }
                    client.DownloadFile("link to web host/MyAppSetup.zip", @"MyAppSetup.zip");
                    string zipPath = @".\MyAppSetup.zip";
                    string extractPath = @".\";
                    ZipFile.ExtractToDirectory(zipPath, extractPath);
                    Process process = new Process();
                    process.StartInfo.FileName = "msiexec.exe";
                    process.StartInfo.Arguments = string.Format("/i MyAppSetup.msi");
                    process.Start();
                }
                catch
                {
                }
            }
        }
    }
}
