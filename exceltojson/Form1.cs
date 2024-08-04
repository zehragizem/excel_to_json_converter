using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using OfficeOpenXml;
using Newtonsoft.Json;
using Formatting = Newtonsoft.Json.Formatting;

namespace exceltojson
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filePath;
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xlsx";
            openFileDialog1.Title = "Excel dosyasýný seçin";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog1.FileName;

                // Dosya yolunu TextBox'a yazdýrýn
                textBox1.Text = filePath;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            // Dictionaries for both tables
            var rowHeadersKalkýþ = new Dictionary<string, List<int>>
        {
            { "Önceki servis / Previous service ", new List<int>() },
            { "Açýklamalar / Remarks ", new List<int>() },
            { "Varýþ noktasý / Destination Codes  ", new List<int>() },
            { "Hat-sefer-trip / Line-run-trip ", new List<int>() },
            { "Sonraki servis / Next service", new List<int>() },
            { "Açýklamalar / Remarks", new List<int>() },
        };
            var resultsRowKalkýþ = new Dictionary<string, List<string>>
        {
            { "Önceki servis / Previous service ", new List<string>() },
            { "Açýklamalar / Remarks ", new List<string>() },
            { "Varýþ noktasý / Destination Codes  ", new List<string>() },
            { "Hat-sefer-trip / Line-run-trip ", new List<string>() },
            { "Sonraki servis / Next service", new List<string>() },
            { "Açýklamalar / Remarks", new List<string>() },
        };

            var rowHeadersVarýþ = new Dictionary<string, List<int>>
        {
            { "Önceki servis / Previous service ", new List<int>() },
            { "Açýklamalar / Remarks ", new List<int>() },
            { "Varýþ noktasý / Destination Codes  ", new List<int>() },
            { "Hat-sefer-trip / Line-run-trip ", new List<int>() },
            { "Sonraki servis / Next service", new List<int>() },
            { "Açýklamalar / Remarks", new List<int>() },
        };
            var resultsRowVarýþ = new Dictionary<string, List<string>>
        {
            { "Önceki servis / Previous service ", new List<string>() },
            { "Açýklamalar / Remarks ", new List<string>() },
            { "Varýþ noktasý / Destination Codes  ", new List<string>() },
            { "Hat-sefer-trip / Line-run-trip ", new List<string>() },
            { "Sonraki servis / Next service", new List<string>() },
            { "Açýklamalar / Remarks", new List<string>() },
        };

            var resultsKalkýþ = new Dictionary<string, List<string>>
        {
            { "km", new List<string>() },
            { "Ýstasyonlar / Stations", new List<string>() },
            { "Ýst. Arasý Yolculuk Süresi", new List<string>() },
            { "Ýstasyon Bekleme Süreleri", new List<string>() },
            { "Toplam Yolculuk Süresi", new List<string>() },
            { "kalkýþ", new List<string>() },
        };

            var resultsVarýþ = new Dictionary<string, List<string>>
        {
            { "km", new List<string>() },
            { "Ýstasyonlar / Stations", new List<string>() },
            { "Ýst. Arasý Yolculuk Süresi", new List<string>() },
            { "Ýstasyon Bekleme Süreleri", new List<string>() },
            { "Toplam Yolculuk Süresi", new List<string>() },
            { "varýþ", new List<string>() },
        };

            var kalkýþListesi = new List<List<string>>();
            var varýþListesi = new List<List<string>>();


            try
            {
                UpdateHeaderPositions(filePath, rowHeadersKalkýþ, rowHeadersVarýþ);
                int firstKmHeaderRow, firstKmHeaderCol;
                int secondKmHeaderRow, secondKmHeaderCol;
                FindKmHeaderPositions(filePath, out firstKmHeaderRow, out firstKmHeaderCol, out secondKmHeaderRow, out secondKmHeaderCol);
                ReadAndDisplayRowData(filePath, rowHeadersKalkýþ, resultsRowKalkýþ, rowHeadersVarýþ, resultsRowVarýþ);
                ReadAndDisplayColumnData(filePath, resultsKalkýþ, kalkýþListesi, resultsVarýþ, varýþListesi);
                var myClassKalkýþ = new MyClass(); // Initialize for kalkýþ
                var myClassVarýþ = new MyClass(); // Initialize for varýþ
                var (counter_kalkis, items_kalkis) = countList(kalkýþListesi);
                var (counter_varis, items_varis) = countList(varýþListesi);

                var result_kalkis = setJson(resultsKalkýþ, "kalkýþ", myClassKalkýþ, items_kalkis, kalkýþListesi, resultsRowKalkýþ, counter_kalkis);
                string json_kalkis = JsonConvert.SerializeObject(result_kalkis, Formatting.Indented);

                var result_varis = setJson(resultsVarýþ, "varýþ", myClassVarýþ, items_varis, varýþListesi, resultsRowVarýþ, counter_varis);
                string json_varis = JsonConvert.SerializeObject(result_varis, Formatting.Indented);

                // Save JSON files using SaveFileDialog
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    Title = "Save JSON File"
                };

                // Save kalkýþ JSON
                saveFileDialog.FileName = "results_kalkis.json";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, json_kalkis);

                }

                // Save varýþ JSON
                saveFileDialog.FileName = "results_varis.json";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    File.WriteAllText(saveFileDialog.FileName, json_varis);

                }
                MessageBox.Show("Dosyalar baþarýyla oluþturuldu");
                textBox1.Text = ""; // TextBox'ý temizler
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }



        static ResultsCombine setJson(Dictionary<string, List<string>> results, string result_key, MyClass myClass, int items, List<List<string>> liste, Dictionary<string, List<string>> resultsRow, int counter)
        {


            foreach (var result in results)
            {
                if (result.Key != result_key)
                {
                    switch (result.Key)
                    {
                        case "km":
                            myClass.Km = string.Join(", ", result.Value);
                            break;
                        case "Ýstasyonlar / Stations":
                            myClass.Istasyon = string.Join(", ", result.Value);
                            break;
                        case "Ýst. Arasý Yolculuk Süresi":
                            myClass.IstasyonArasiYolculuk = string.Join(", ", result.Value);
                            break;
                        case "Ýstasyon Bekleme Süreleri":
                            myClass.IstasyonBeklemeSuresi = string.Join(", ", result.Value);
                            break;
                        case "Toplam Yolculuk Süresi":
                            myClass.ToplamYolculukSuresi = string.Join(", ", result.Value);
                            break;
                    }
                }
            }

            myClass.MyList = new List<MyList>();
            int index = 0;

            for (int i = 0; i <= items; i++)
            {
                foreach (var subList in liste)
                {
                    if (subList.Count > i)
                    {
                        var kalkýþ = subList[i];
                        var myListItem = new MyList { Saat = kalkýþ };

                        foreach (var rowResult in resultsRow)
                        {
                            if (rowResult.Value.Count > index)
                            {
                                switch (rowResult.Key)
                                {
                                    case "Önceki servis / Previous service ":
                                        myListItem.ÖncekiServis = rowResult.Value[index];
                                        break;
                                    case "Açýklamalar / Remarks ":
                                        myListItem.Açýklama = rowResult.Value[index];
                                        break;
                                    case "Varýþ noktasý / Destination Codes  ":
                                        myListItem.VarýþNoktasý = rowResult.Value[index];
                                        break;
                                    case "Hat-sefer-trip / Line-run-trip ":
                                        myListItem.HatSefer = rowResult.Value[index];
                                        break;
                                    case "Sonraki servis / Next service":
                                        myListItem.SonrakiSefer = rowResult.Value[index];
                                        break;
                                }
                            }
                        }

                        myClass.MyList.Add(myListItem);
                        index++;

                        if (index >= resultsRow.Values.First().Count)
                        {
                            index = 0;
                        }
                    }
                }
            }

            var results_combine = new ResultsCombine();
            var kmValues = myClass.Km.Split(", ");
            var istasyonValues = myClass.Istasyon.Split(", ");
            var arasiYolculukValues = myClass.IstasyonArasiYolculuk.Split(", ");
            var beklemeSuresiValues = myClass.IstasyonBeklemeSuresi.Split(", ");
            var toplamYolculukSuresiValues = myClass.ToplamYolculukSuresi.Split(", ");

            int maxCount = Math.Max(kmValues.Length, Math.Max(istasyonValues.Length, Math.Max(arasiYolculukValues.Length, Math.Max(beklemeSuresiValues.Length, toplamYolculukSuresiValues.Length))));

            int chunkSize = counter; // Number of items to add at a time
            int offset = 0;

            for (int x = 0; x < maxCount; x++)
            {
                var myResults = new MyResults
                {
                    Km = kmValues.Length > x ? kmValues[x] : "",
                    Istasyon = istasyonValues.Length > x ? istasyonValues[x] : "",
                    IstasyonArasiYolculuk = arasiYolculukValues.Length > x ? arasiYolculukValues[x] : "",
                    IstasyonBeklemeSuresi = beklemeSuresiValues.Length > x ? beklemeSuresiValues[x] : "",
                    ToplamYolculukSuresi = toplamYolculukSuresiValues.Length > x ? toplamYolculukSuresiValues[x] : "",
                    MyList = new List<MyList>() // Initialize the list
                };

                // Add items from myLists in chunks
                for (int y = offset; y < offset + chunkSize - 1 && y < myClass.MyList.Count; y++)
                {
                    var firstItem = myClass.MyList[y];
                    var myListItem = new MyList // Assuming MyList is a class
                    {
                        Saat = firstItem.Saat,
                        ÖncekiServis = firstItem.ÖncekiServis,
                        Açýklama = firstItem.Açýklama,
                        VarýþNoktasý = firstItem.VarýþNoktasý,
                        HatSefer = firstItem.HatSefer,
                        SonrakiSefer = firstItem.SonrakiSefer
                    };

                    myResults.MyList.Add(myListItem);
                }

                results_combine.MyResultsList.Add(myResults);

                // Increment offset for the next chunk
                offset += chunkSize;
            }

            return results_combine;
        }


        static void UpdateHeaderPositions(string filePath, Dictionary<string, List<int>> rowHeadersKalkýþ, Dictionary<string, List<int>> rowHeadersVarýþ)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                int firstTableEndRow = 0;
                bool foundFirstTableEnd = false;

                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;

                        // Update positions for 'kalkýþ' headers
                        if (rowHeadersKalkýþ.ContainsKey(cellValue) && rowHeadersKalkýþ[cellValue].Count == 0)
                        {
                            rowHeadersKalkýþ[cellValue].Add(row);
                            rowHeadersKalkýþ[cellValue].Add(col);
                        }

                        // If first table end is not yet found, mark it and switch to second table headers
                        if (rowHeadersKalkýþ.Values.All(v => v.Count > 1) && !foundFirstTableEnd)
                        {
                            firstTableEndRow = row;
                            foundFirstTableEnd = true;
                        }

                        // Update positions for 'varýþ' headers only after the end of the first table
                        if (foundFirstTableEnd && rowHeadersVarýþ.ContainsKey(cellValue) && rowHeadersVarýþ[cellValue].Count == 0)
                        {
                            rowHeadersVarýþ[cellValue].Add(row);
                            rowHeadersVarýþ[cellValue].Add(col);
                        }
                    }
                }
            }
        }


        static (int countListe, int itemIndex) countList(List<List<string>> Liste)
        {
            int itemIndex = 0;
            int countListe = 0;

            foreach (var subList in Liste)
            {
                itemIndex = 0;  // Reset itemIndex for each subList
                foreach (var item in subList)
                {
                    itemIndex++;
                }
                countListe++;
            }

            return (countListe, itemIndex);
        }
        static void FindKmHeaderPositions(string filePath, out int firstKmHeaderRow, out int firstKmHeaderCol, out int secondKmHeaderRow, out int secondKmHeaderCol)
        {
            firstKmHeaderRow = -1;
            firstKmHeaderCol = -1;
            secondKmHeaderRow = -1;
            secondKmHeaderCol = -1;

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int occurrenceCount = 0;

                for (int row = 1; row <= worksheet.Dimension.Rows; row++)
                {
                    for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text;

                        if (cellValue.Contains("km", StringComparison.OrdinalIgnoreCase))
                        {
                            occurrenceCount++;
                            if (occurrenceCount == 1)
                            {
                                firstKmHeaderRow = row;
                                firstKmHeaderCol = col;
                            }
                            else if (occurrenceCount == 2)
                            {
                                secondKmHeaderRow = row;
                                secondKmHeaderCol = col;
                                break;
                            }
                        }
                    }
                    if (secondKmHeaderRow != -1) break;
                }
            }
        }

        static void ReadAndDisplayRowData(
            string filePath,
            Dictionary<string, List<int>> rowHeadersKalkýþ,
            Dictionary<string, List<string>> resultsRowKalkýþ,
            Dictionary<string, List<int>> rowHeadersVarýþ,
            Dictionary<string, List<string>> resultsRowVarýþ)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Process kalkýþ table
                foreach (var header in rowHeadersKalkýþ)
                {
                    if (header.Value.Count == 2)
                    {
                        int row = header.Value[0];
                        int startCol = header.Value[1];
                        int targetCol = startCol + 7;
                        int totalFilledColumns = CountTotalFilledColumns(worksheet);

                        while (targetCol <= totalFilledColumns + 1)
                        {
                            string cellValue = worksheet.Cells[row, targetCol].Text;
                            resultsRowKalkýþ[header.Key].Add(!string.IsNullOrEmpty(cellValue) ? cellValue : "null");
                            targetCol++;
                        }

                        Console.WriteLine($"Processed kalkýþ header: {header.Key}");
                    }
                }

                // Process varýþ table
                foreach (var header in rowHeadersVarýþ)
                {
                    if (header.Value.Count == 2)
                    {
                        int row = header.Value[0];
                        int startCol = header.Value[1];
                        int targetCol = startCol + 7;
                        int totalFilledColumns = CountTotalFilledColumns(worksheet);

                        while (targetCol <= totalFilledColumns + 1)
                        {
                            string cellValue = worksheet.Cells[row, targetCol].Text;
                            resultsRowVarýþ[header.Key].Add(!string.IsNullOrEmpty(cellValue) ? cellValue : "null");
                            targetCol++;
                        }

                        Console.WriteLine($"Processed varýþ header: {header.Key}");
                    }
                }
            }
        }

        static int CountTotalFilledColumns(ExcelWorksheet worksheet)
        {
            int totalFilledColumns = 0;
            for (int col = 1; col <= worksheet.Dimension.Columns; col++)
            {
                if (!string.IsNullOrEmpty(worksheet.Cells[1, col].Text))
                {
                    totalFilledColumns = col;
                }
            }
            return totalFilledColumns;
        }

        static void ReadAndDisplayColumnData(string filePath,
            Dictionary<string, List<string>> resultsKalkýþ, List<List<string>> kalkýþListesi,
            Dictionary<string, List<string>> resultsVarýþ, List<List<string>> varýþListesi)
        {
            int firstKmHeaderRow, firstKmHeaderCol, secondKmHeaderRow, secondKmHeaderCol;
            FindKmHeaderPositions(filePath, out firstKmHeaderRow, out firstKmHeaderCol, out secondKmHeaderRow, out secondKmHeaderCol);

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // Process kalkýþ data if first km header is found
                if (firstKmHeaderCol != -1)
                {
                    ProcessColumnData(worksheet, firstKmHeaderRow, firstKmHeaderCol, resultsKalkýþ, kalkýþListesi, "kalkýþ");
                }
                else
                {
                    Console.WriteLine("First 'km' header not found for kalkýþ table.");
                }

                // Process varýþ data if second km header is found
                if (secondKmHeaderCol != -1)
                {
                    ProcessColumnData(worksheet, secondKmHeaderRow, secondKmHeaderCol, resultsVarýþ, varýþListesi, "varýþ");
                }
                else
                {
                    Console.WriteLine("Second 'km' header not found for varýþ table.");
                }
            }
        }

        static void ProcessColumnData(ExcelWorksheet worksheet, int startRow, int col,
            Dictionary<string, List<string>> results, List<List<string>> list, string tableName)
        {
            int[] adjacentCols = { col + 1, col + 3, col + 4, col + 6, col + 8 };
            int count = 0;

            // Read the "km" column data
            for (int row = startRow + 1; row <= worksheet.Dimension.Rows; row++)
            {
                string cellValue = worksheet.Cells[row, col].Text;
                if (!string.IsNullOrEmpty(cellValue))
                {
                    results["km"].Add(cellValue);
                    count++;
                }
                else
                {
                    break;
                }
            }

            // Read data for other columns
            List<string> columns = new List<string> { "Ýstasyonlar / Stations", "Ýst. Arasý Yolculuk Süresi", "Ýstasyon Bekleme Süreleri", "Toplam Yolculuk Süresi", tableName };
            int totalFilledColumns = CountTotalFilledColumns(worksheet);

            for (int i = 0; i < adjacentCols.Length; i++)
            {
                string cheader = columns[i];

                if (cheader == tableName)
                {
                    int initialRow = startRow + 1;
                    int currentCol = adjacentCols[i];

                    for (int x = currentCol; x <= totalFilledColumns + 1; x++)
                    {
                        var tempList = new List<string>();

                        for (int y = initialRow; y < initialRow + count; y++)
                        {
                            string cellValue = worksheet.Cells[y, x].Text;
                            tempList.Add(!string.IsNullOrEmpty(cellValue) ? cellValue : "null");
                        }

                        list.Add(tempList);
                    }
                }
                else
                {
                    for (int row = startRow + 1; row < startRow + 1 + count; row++)
                    {
                        string adjacentCellValue = worksheet.Cells[row, adjacentCols[i]].Text;
                        results[cheader].Add(!string.IsNullOrEmpty(adjacentCellValue) ? adjacentCellValue : "null");
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

