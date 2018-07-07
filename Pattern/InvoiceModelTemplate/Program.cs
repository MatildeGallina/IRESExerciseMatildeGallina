using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace InvoiceModelTemplate
{
    class Program
    {
        static void Main(string[] args)
        {
            var models = new List<InvoiceModel>
            {
                new InvoiceModel(1, "Falciatrice", 74.99m, 2, 1, 2018, 6, 14, 9, 30, 12, "Mario"),
                new InvoiceModel(2, "Sega Elettrica", 150m, 3, 1, 2018, 6, 14, 9, 30, 12, "Mario"),
                new InvoiceModel(3, "Tritatutto", 39.99m, 1, 2, 2018, 6, 14, 20, 40, 00, "Mario"),
                new InvoiceModel(4, "Concime", 5.99m, 10, 3, 2018, 6, 16, 12, 12, 12, "Luigi"),
                new InvoiceModel(5, "Forbici", 7.89m, 10, 3, 2018, 6, 16, 12, 12, 12, "Luigi"),
                new InvoiceModel(6, "Tenaglie", 8.99m, 7, 3, 2018, 6, 16, 12, 12, 12, "Luigi"),
            };

            var flatDataPath = createDocument("FlatData.xlsx");
            var salesPerEmployeePath = createDocument("TotalSalesPerEmployee.xlsx");
            var billInvoicePath = createDocument("BillInvoice.xlsx");

            var exel1 = new FlatDataExporter(flatDataPath);
            var exel2 = new SalesPerEmployeeExporter(salesPerEmployeePath);
            var exel3 = new BillInvoiceExporter(billInvoicePath);

            exel1.Export(models);
            exel2.Export(models);
            exel3.Export(models);
        }

        private static string createDocument(string nameFile)
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                nameFile);
        }
    }

    public abstract class ExcelExporter<T>
    {
        private string _filePath;

        protected ExcelExporter(string filePath)
        {
            _filePath = filePath;
        }

        public void Export(IEnumerable<T> models)
        {
            var matrix = transform(models);

            writeOnExcel(matrix);
        }

        protected abstract List<string> createHeadersList();
        protected abstract void AddElaborateValues(List<Cell> cells, IEnumerable<T> models);

        protected IEnumerable<Cell> transform(IEnumerable<T> models)
        {
            var cells = new List<Cell>();

            AddHeaders(cells);
            AddElaborateValues(cells, models);

            return cells;
        }

        protected void AddHeaders(List<Cell> cells)
        {
            List<string> headers = createHeadersList();

            for (int i = 0; i < headers.Count; i++)
            {
                cells.Add(new Cell { Row = 1, Column = i + 1, Value = headers[i] });
            }
        }

        protected void writeOnExcel(IEnumerable<Cell> cells)
        {
            var outputFile = new FileInfo(_filePath);

            using (var package = new ExcelPackage(outputFile))
            {
                var ws = package.Workbook.Worksheets.Add("export");

                foreach (var cell in cells)
                    ws.Cells[cell.Row, cell.Column].Value = cell.Value.ToString();

                package.Save();
            }
        }
    }

    public class FlatDataExporter : ExcelExporter<InvoiceModel>
    {
        public FlatDataExporter(string filePath)
            : base(filePath)
        { }

        protected override List<string> createHeadersList()
        {
            return new List<string>
                {
                    "Id",
                    "ProductName",
                    "ProductPrice",
                    "Quantity",
                    "BillId",
                    "BillEmission",
                    "Employee"
                };
        }

        protected override void AddElaborateValues(List<Cell> cells, IEnumerable<InvoiceModel> models)
        {
            int row = 3;

            foreach (var value in models)
            {
                cells.Add(new Cell(row, 1, value.Id));
                cells.Add(new Cell(row, 2, value.ProductName));
                cells.Add(new Cell(row, 3, value.ProductPrice));
                cells.Add(new Cell(row, 4, value.Quantity));
                cells.Add(new Cell(row, 5, value.BillId));
                cells.Add(new Cell(row, 6, value.BillEmission));
                cells.Add(new Cell(row, 7, value.Employee));

                row++;
            }
        }
    }

    public class SalesPerEmployeeExporter : ExcelExporter<InvoiceModel>
    {
        public SalesPerEmployeeExporter(string filePath)
            : base(filePath)
        { }

        protected override List<string> createHeadersList()
        {
            return new List<string>
                {
                    "Employee",
                    "Total Sales"
                };
        }

        protected override void AddElaborateValues(List<Cell> cells, IEnumerable<InvoiceModel> models)
        {
            var result = models
                .GroupBy(x => x.Employee)
                .Select(g => new
                {
                    Name = g.Key,
                    TotalSales = g.Sum(i => i.ProductPrice * i.Quantity)
                });

            int row = 3;

            foreach (var employee in result)
            {
                cells.Add(new Cell(row, 1, employee.Name));
                cells.Add(new Cell(row, 2, employee.TotalSales));

                row++;
            }
        }
    }

    public class BillInvoiceExporter : ExcelExporter<InvoiceModel>
    {
        public BillInvoiceExporter(string filePath)
            : base(filePath)
        { }

        protected override List<string> createHeadersList()
        {
            return new List<string>
                {
                    "Bill Id",
                    "Emission Date",
                    "Employee Name",
                    "Invoice Id",
                    "Product Name",
                    "Product Price",
                    "Quantity"
                };
        }

        protected override void AddElaborateValues(List<Cell> cells, IEnumerable<InvoiceModel> models)
        {
            var result = models
                .GroupBy(x => x.BillId)
                .Select(g => new
                {
                    Id = g.Key,
                    Emission = g.First().BillEmission,
                    EmployeeName = g.First().Employee,
                    Invoices = g
                            .Select(im => new
                            {
                                InvoiceId = im.Id,
                                InvoiceName = im.ProductName,
                                InvoicePrice = im.ProductPrice,
                                InvoiceQuantity = im.Quantity
                            })
                            .ToList()
                });

            int row = 3;

            foreach (var value in result)
            {
                cells.Add(new Cell(row, 1, value.Id));
                cells.Add(new Cell(row, 2, value.Emission));
                cells.Add(new Cell(row, 3, value.EmployeeName));

                row++;

                foreach (var invoice in value.Invoices)
                {
                    cells.Add(new Cell(row, 4, invoice.InvoiceId));
                    cells.Add(new Cell(row, 5, invoice.InvoiceName));
                    cells.Add(new Cell(row, 6, invoice.InvoicePrice));
                    cells.Add(new Cell(row, 7, invoice.InvoiceQuantity));

                    row++;
                }

                row++;
            }

        }
    }

    public class Cell
    {
        public Cell() { }

        public Cell(int row, int column, object value)
        {
            Row = row;
            Column = column;
            Value = value;
        }

        public int Row { get; set; }
        public int Column { get; set; }
        public object Value { get; set; }
    }

    public class InvoiceModel
    {
        public InvoiceModel() { }

        public InvoiceModel(
            int id,
            string productName,
            decimal productPrice,
            int quantity,
            int billId,
            int year,
            int month,
            int day,
            int hour,
            int minute,
            int second,
            string employee)
        {
            Id = id;
            ProductName = productName;
            ProductPrice = productPrice;
            Quantity = quantity;
            BillId = billId;
            BillEmission = new DateTime(year, month, day, hour, minute, second);
            Employee = employee;
        }

        public int Id { get; set; }
        public string ProductName { get; set; }
        public decimal ProductPrice { get; set; }
        public int Quantity { get; set; }
        public int BillId { get; set; }
        public DateTime BillEmission { get; set; }
        public string Employee { get; set; }
    }
}
