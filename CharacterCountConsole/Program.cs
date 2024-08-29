using System.Drawing;
using System.Text.RegularExpressions;
using CountAppApp.Dtos;
using Npgsql;
using OfficeOpenXml;
using OfficeOpenXml.Style;

// Set the console encoding to UTF-8
Console.OutputEncoding = System.Text.Encoding.UTF8;
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string connectionString = "Host=10.111.15.32;Username=postgres;Password=postgres;Database=Ettn";
List<ProductDto> productList = new List<ProductDto>();

using (NpgsqlConnection connection = new NpgsqlConnection(connectionString))
{
    connection.Open();
    Console.WriteLine("Connected to PostgreSQL!");

    string sqlQuery = @"SELECT * FROM ""Products""";

    using (NpgsqlCommand command = new NpgsqlCommand(sqlQuery, connection))
    {
        using (NpgsqlDataReader reader = command.ExecuteReader())
        {
            while (reader.Read())
            {
                ProductDto data = new ProductDto
                {
                    Name = reader.GetString(reader.GetOrdinal("name"))
                };

                productList.Add(data);
            }
        }
    }
}

string spacePattern = @"\s+";

Regex regexSpace = new Regex(spacePattern);
var charList = new List<CharCount>();

foreach (var item in productList)
{
    string wordWithoutSpace = regexSpace.Replace(item.Name, "");

    var localCharList = wordWithoutSpace.GroupBy(x => x)
        .Select(s => new CharCount { ItemChar = s.Key, ItemCount = s.Count() }).ToList();
    charList.AddRange(localCharList);
}

var totalCharList = charList.GroupBy(x => x.ItemChar).Select(s => new CharCount()
{
    ItemChar = s.Key,
    ItemCount = s.Sum(x => x.ItemCount)
}).ToList();

// foreach (var item in totalCharList.OrderBy(x=>x.ItemCount))
// {
//     Console.WriteLine($"Symbol: {item.ItemChar} | Count: {item.ItemCount} \n");
// }    Console.ReadKey();

var totalCharListLast = new List<CharCount>();

foreach (var item in totalCharList)
{
    totalCharListLast.Add(new CharCount()
    {
        ItemChar = item.ItemChar,
        ItemCount = item.ItemCount,
        WhatIsLetter = DefineLetter(item.ItemChar)
    });
}

//Writing file
string filePath = "C:\\Users\\melis\\Desktop\\TextFiles\\File.xlsx";


using (var package = new ExcelPackage())
{
    // Add a new worksheet to the Excel package
    var sheet = package.Workbook.Worksheets.Add("Sheet1");

    // Write data to the worksheet
    sheet.Cells["A1"].Value = "№";
    sheet.Cells["B1"].Value = "Symbol";
    sheet.Cells["C1"].Value = "Letter";
    sheet.Cells["D1"].Value = "Count";
    int row = 1;
    foreach (var item in totalCharListLast.OrderByDescending(x=>x.ItemCount))
    {
        row++;
        sheet.Cells[row, 1].Value = (row - 1).ToString();
        sheet.Cells[row, 2].Value = item.ItemChar;
        sheet.Cells[row, 3].Value = item.WhatIsLetter;
        sheet.Cells[row, 4].Value = item.ItemCount;

        sheet.Cells[row, 1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells[row, 1].Style.Font.Color.SetColor(Color.FromArgb(156, 87, 0));
        sheet.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
        sheet.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(255, 235, 156));

        sheet.Row(1).Height = 20;
        sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        sheet.Column(1).Width = 20;
        sheet.Column(2).Width = 20;
        sheet.Column(3).Width = 20;
        sheet.Column(4).Width = 20;
        sheet.Cells[sheet.Dimension.Address].Style.Font.Size = 12;
        sheet.Cells[sheet.Dimension.Address].Style.Font.Name = "Century Gothic";
        sheet.Cells[sheet.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

        sheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

        sheet.Cells["A1:L1"].Style.Font.Bold = true;
        sheet.Cells["A1:L1"].Style.Font.Color.SetColor(Color.FromArgb(0, 97, 0));
        sheet.Cells["A1:L1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        sheet.Cells["A1:L1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        sheet.Cells["A1:L1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        sheet.Cells["A1:L1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(198, 239, 206));

        sheet.Cells[sheet.Dimension.Address].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        sheet.Cells[sheet.Dimension.Address].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        sheet.Cells[sheet.Dimension.Address].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        sheet.Cells[sheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        sheet.Cells[sheet.Dimension.Address].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
    }

    // Save the Excel package to a file
    package.SaveAs(new System.IO.FileInfo(filePath));
}

Console.WriteLine("Excel file has been created.");
Console.ReadKey();

static string DefineLetter(char c)
{
    // Define Regex patterns for Latin and Cyrillic characters
    string latinPatternWithoutNumbers = @"^\p{IsBasicLatin}$";
    string cyrillicPatternWithoutNumbers = @"^\p{IsCyrillic}$";
    string symbolPattern = @"[<>"":;\[\]{}|\\/,.'!@#$%^&*()\-_=+№?]";
    string numberPattern = @"[\d+]";

    if (Regex.IsMatch(c.ToString(), symbolPattern))
    {
        return "Symbol";
    }

    if (Regex.IsMatch(c.ToString(), latinPatternWithoutNumbers))
    {
        if (Regex.IsMatch(c.ToString(), numberPattern))
        {
            return "Number";
        }

        return "Latin";
    }

    if (Regex.IsMatch(c.ToString(), cyrillicPatternWithoutNumbers))
    {
        if (Regex.IsMatch(c.ToString(), numberPattern))
        {
            return "Number";
        }

        return "Cyrillic";
    }

    {
        return "Other";
    }
}

class CharCount
{
    public char ItemChar { get; set; }
    public int ItemCount { get; set; }
    public string WhatIsLetter { get; set; }
}