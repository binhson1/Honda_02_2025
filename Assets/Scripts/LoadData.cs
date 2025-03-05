using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using UnityEngine;

public class LoadData : MonoBehaviour
{
    public string dataPath = "";

    [System.Serializable]
    public class PlayerData
    {
        public int id;
        public string manhanvien;
        public string hovaten;
        public string phong;
        public string note;
        public bool isWon;
        public string prizeName;
    }

    public class Prize
    {
        public string Name;
        public int TotalQuantity;
        public int RemainingQuantity;
        public int BatchSize;
    }

    public List<Prize> Morningprizes = new List<Prize>
    {
        new Prize { Name = "Vali", TotalQuantity = 100, RemainingQuantity = 100, BatchSize = 10 },
        new Prize { Name = "Nồi Chiên", TotalQuantity = 20, RemainingQuantity = 20, BatchSize = 10 },
        new Prize { Name = "Máy Giặt", TotalQuantity = 15, RemainingQuantity = 15, BatchSize = 15 },
        new Prize { Name = "Tủ Lạnh", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
        new Prize { Name = "Tivi", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 }
    };

    public List<Prize> Afternoonprizes = new List<Prize>
    {
        new Prize { Name = "Máy Phát Điện Honda", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
        new Prize { Name = "Xe Máy Honda Wave Alpha", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
        new Prize { Name = "Xe Máy Honda Blade", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
        new Prize { Name = "Xe Máy Honda Vision", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
        new Prize { Name = "Xe Máy Honda Lead", TotalQuantity = 3, RemainingQuantity = 3, BatchSize = 3 },
        new Prize { Name = "Giải đặc biệt", TotalQuantity = 4, RemainingQuantity = 4, BatchSize = 4 },
        new Prize { Name = "Giải thưởng bất ngờ", TotalQuantity = 1, RemainingQuantity = 1, BatchSize = 1 }
    };

    private string CurrentSheetName;
    public List<Prize> prizes;
    public SelectPrize selectPrize;
    public List<PlayerData> playerDataList = new List<PlayerData>();
    public List<PlayerData> wonPlayer = new List<PlayerData>();

    void Start()
    {
        string filePath = Path.Combine(Path.GetDirectoryName(Application.dataPath), "Data");
        dataPath = filePath + "/Data.xlsx";
        // dataPath = "D:\\STUDYING\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Data.xlsx";
        LoadOrCreatePrizes(dataPath);
        ReadPlayerData(dataPath);
        prizes = Morningprizes;
        CurrentSheetName = "MorningPrizes";
    }

    public void SwitchToMorning()
    {
        prizes = Morningprizes;
        CurrentSheetName = "MorningPrizes";
    }

    public void WritePrizeRemainQuantity()
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(dataPath)))
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets[CurrentSheetName];
            if (sheet != null)
            {
                sheet.Cells[selectPrize.currentPrizeIndex + 2, 3].Value = prizes[selectPrize.currentPrizeIndex].RemainingQuantity;
                // package.Save();
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream);
                    File.WriteAllBytes(dataPath, memoryStream.ToArray());
                }
            }
        }
    }

    public void SaveExcel(List<PlayerData> wonPlayer, string prizeName)
    {
        if (File.Exists(dataPath))
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(dataPath)))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];
                if (sheet != null)
                {
                    foreach (var player in wonPlayer)
                    {
                        int rowIndex = player.id + 2; // Giả sử dòng đầu tiên là header
                        sheet.Cells[rowIndex, 6].Value = prizeName;
                    }
                }

                // Sử dụng MemoryStream để giảm thời gian thao tác file
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    package.SaveAs(memoryStream); // Lưu vào RAM trước
                    File.WriteAllBytes(dataPath, memoryStream.ToArray()); // Sau đó ghi 1 lần vào file
                }
            }
        }
    }

    public void SwitchToAfternoon()
    {
        prizes = Afternoonprizes;
        CurrentSheetName = "AfternoonPrizes";
    }

    private void LoadOrCreatePrizes(string path)
    {
        FileInfo fileInfo = new FileInfo(path);
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            WriteOrRead(package, "MorningPrizes", Morningprizes);
            WriteOrRead(package, "AfternoonPrizes", Afternoonprizes);
            package.Save();
        }
    }

    private void WriteOrRead(ExcelPackage package, string sheetName, List<Prize> prizeList)
    {
        ExcelWorksheet sheet = package.Workbook.Worksheets[sheetName] ?? package.Workbook.Worksheets.Add(sheetName);

        if (sheet.Dimension != null)
        {
            prizeList.Clear();
            for (int i = 2; i <= sheet.Dimension.Rows; i++)
            {
                prizeList.Add(new Prize
                {
                    Name = sheet.Cells[i, 1].Text,
                    TotalQuantity = int.Parse(sheet.Cells[i, 2].Text),
                    RemainingQuantity = int.Parse(sheet.Cells[i, 3].Text),
                    BatchSize = int.Parse(sheet.Cells[i, 4].Text)
                });
            }
        }
        else
        {
            sheet.Cells[1, 1].Value = "Name";
            sheet.Cells[1, 2].Value = "TotalQuantity";
            sheet.Cells[1, 3].Value = "RemainingQuantity";
            sheet.Cells[1, 4].Value = "BatchSize";
            for (int i = 0; i < prizeList.Count; i++)
            {
                sheet.Cells[i + 2, 1].Value = prizeList[i].Name;
                sheet.Cells[i + 2, 2].Value = prizeList[i].TotalQuantity;
                sheet.Cells[i + 2, 3].Value = prizeList[i].RemainingQuantity;
                sheet.Cells[i + 2, 4].Value = prizeList[i].BatchSize;
            }
        }
    }

    private void ReadPlayerData(string path)
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
        {
            ExcelWorksheet sheet = package.Workbook.Worksheets["Data"];
            if (sheet != null)
            {
                for (int i = 3; i <= sheet.Dimension.Rows; i++)
                {
                    string prize = sheet.Cells[i, 6].Text;
                    PlayerData data = new PlayerData
                    {
                        id = int.Parse(sheet.Cells[i, 1].Text),
                        manhanvien = sheet.Cells[i, 2].Text,
                        hovaten = sheet.Cells[i, 3].Text,
                        phong = sheet.Cells[i, 4].Text,
                        note = sheet.Cells[i, 5].Text,
                        isWon = !string.IsNullOrEmpty(prize),
                        prizeName = prize
                    };
                    if (data.isWon)
                        wonPlayer.Add(data);
                    else
                        playerDataList.Add(data);
                }
            }
        }
    }
}
