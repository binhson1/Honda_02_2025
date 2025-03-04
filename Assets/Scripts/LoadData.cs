using System.Collections;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using UnityEngine;

public class LoadData : MonoBehaviour
{
    public string dataPath = "";
    // public string prizePath = "D:\\STUDYING\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Prize.xlsx";    
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
        public int BatchSize; // Số lượng quay mỗi lần
    }

    public List<Prize> Morningprizes = new List<Prize>{
            new Prize { Name = "Vali", TotalQuantity = 100, RemainingQuantity = 100, BatchSize = 10 },
            new Prize { Name = "Nồi Chiên", TotalQuantity = 20, RemainingQuantity = 20, BatchSize = 10 },
            new Prize { Name = "Máy Giặt", TotalQuantity = 15, RemainingQuantity = 15, BatchSize = 15 },
            new Prize { Name = "Tủ Lạnh", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Tivi", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },

    };
    public List<Prize> Afternoonprizes = new List<Prize>{
            new Prize { Name = "Máy Phát Điện Honda", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Wave Alpha", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Blade", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Vision", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
            new Prize { Name = "Xe Máy Honda Lead", TotalQuantity = 3, RemainingQuantity = 3, BatchSize = 3 },
            new Prize { Name = "Giải đặc biệt", TotalQuantity = 4, RemainingQuantity = 4, BatchSize = 4 },
    };
    private string CurrentSheetName;
    public SelectPrize selectPrize;
    public List<Prize> prizes;
    public List<PlayerData> playerDataList = new List<PlayerData>();
    public List<PlayerData> wonPlayer = new List<PlayerData>();

    // Start is called before the first frame update
    void Start()
    {
        string filePath = Path.Combine(Path.GetDirectoryName(Application.dataPath), "Data");
        if (!Directory.Exists(filePath))
        {
            Directory.CreateDirectory(filePath);
        }
        dataPath = filePath + "/Data.xlsx";
        dataPath = "D:\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Data.xlsx";
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
        using (FileStream stream = new FileStream(dataPath, FileMode.Open, FileAccess.ReadWrite))
        {
            IWorkbook workbook = new XSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheet(CurrentSheetName);
            if (sheet != null)
            {
                IRow row = sheet.GetRow(selectPrize.currentPrizeIndex + 1);
                if (row != null)
                {
                    row.GetCell(2).SetCellValue(prizes[selectPrize.currentPrizeIndex].RemainingQuantity);
                }
            }
            using (FileStream writeStream = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(writeStream);
            }
        }
    }
    public void SaveExcel(List<PlayerData> wonPlayer, string prizeName)
    {
        if (File.Exists(dataPath))
        {
            IWorkbook workbook = OpenAndReadWorkbook(dataPath);
            ISheet sheet = workbook.GetSheet("Data");
            if (sheet != null)
            {
                for (int i = 0; i < wonPlayer.Count; i++)
                {
                    IRow row = sheet.GetRow(wonPlayer[i].id + 1);
                    if (row != null)
                    {
                        row.CreateCell(5).SetCellValue(prizeName);
                    }
                }
            }
            using (FileStream writeStream = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(writeStream);
            }
        }
    }
    public void SwitchToAfternoon()
    {
        prizes = Afternoonprizes;
        CurrentSheetName = "AfternoonPrizes";
    }
    private IWorkbook OpenAndReadWorkbook(string path)
    {
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        {
            return new XSSFWorkbook(fs);
        }
    }
    private void LoadOrCreatePrizes(string path)
    {
        if (File.Exists(path))
        {
            IWorkbook workbook = OpenAndReadWorkbook(path);
            WriteOrRead(workbook, "MorningPrizes", Morningprizes);
            WriteOrRead(workbook, "Afternoonprizes", Afternoonprizes);
            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }
    private void WriteOrRead(IWorkbook workbook, string sheetName, List<Prize> prizeList)
    {
        ISheet sheet = workbook.GetSheet(sheetName);
        if (sheet != null)
        {
            prizeList.Clear();
            for (int i = 1; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    Prize prize = new Prize
                    {
                        Name = row.GetCell(0).StringCellValue,
                        TotalQuantity = (int)row.GetCell(1).NumericCellValue,
                        RemainingQuantity = (int)row.GetCell(2).NumericCellValue,
                        BatchSize = (int)row.GetCell(3).NumericCellValue
                    };
                    prizeList.Add(prize);
                }
            }
        }
        else
        {
            sheet = workbook.CreateSheet(sheetName);
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Name");
            headerRow.CreateCell(1).SetCellValue("TotalQuantity");
            headerRow.CreateCell(2).SetCellValue("RemainingQuantity");
            headerRow.CreateCell(3).SetCellValue("BatchSize");

            for (int i = 0; i < prizeList.Count; i++)
            {
                IRow row = sheet.GetRow(i + 1) ?? sheet.CreateRow(i + 1);
                row.CreateCell(0).SetCellValue(prizeList[i].Name);
                row.CreateCell(1).SetCellValue(prizeList[i].TotalQuantity);
                row.CreateCell(2).SetCellValue(prizeList[i].RemainingQuantity);
                row.CreateCell(3).SetCellValue(prizeList[i].BatchSize);
            }
        }
    }

    private void ReadPlayerData(string path)
    {
        if (File.Exists(path))
        {
            IWorkbook workbook = OpenAndReadWorkbook(path);
            ISheet sheet = workbook.GetSheet("Data");
            if (sheet != null)
            {
                for (int i = 2; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    PlayerData data;
                    if (row == null) continue;

                    string cell5 = row?.GetCell(5)?.ToString() ?? "";
                    if (cell5 == "")
                    {
                        data = new PlayerData
                        {
                            id = int.Parse(row.GetCell(0).ToString()),
                            manhanvien = row.GetCell(1).ToString(),
                            hovaten = row.GetCell(2).ToString(),
                            phong = row.GetCell(3).ToString(),
                            note = row.GetCell(4).ToString() ?? string.Empty,
                            isWon = false,
                            prizeName = row?.GetCell(5)?.ToString() ?? "",
                        };
                        playerDataList.Add(data);
                    }
                    else
                    {
                        data = new PlayerData
                        {
                            id = int.Parse(row.GetCell(0).ToString()),
                            manhanvien = row.GetCell(1).ToString(),
                            hovaten = row.GetCell(2).ToString(),
                            phong = row.GetCell(3).ToString(),
                            note = row.GetCell(4).ToString() ?? string.Empty,
                            isWon = true,
                            prizeName = row?.GetCell(5)?.ToString() ?? "",
                        };
                        wonPlayer.Add(data);
                    }
                }
                Debug.Log($"Đã tải {playerDataList.Count} dòng dữ liệu.");
            }
        }
    }
    void Update()
    {
    }
}
