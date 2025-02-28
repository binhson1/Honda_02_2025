using UnityEngine;
using System.Collections.Generic;
using System.IO;
using SFB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using UnityEngine.UI;
using System;
using TMPro;
using NPOI.SS.Formula.Functions;

public class LoadData : MonoBehaviour
{
    // public Button selectFileButton;
    public string dataPath = "";
    // public string prizePath = "D:\\STUDYING\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Prize.xlsx";
    public string wonPath = "";
    public TextMeshProUGUI result;

    [System.Serializable]
    public class PlayerData
    {
        public int id;
        public string manhanvien;
        public string hovaten;
        public string phong;
        public string note;

        public bool isWon;

    }

    public class Prize
    {
        public string Name;
        public int TotalQuantity;
        public int RemainingQuantity;
        public int BatchSize; // Số lượng quay mỗi lần
    }
    public List<Prize> Morningprizes = new List<Prize>{
            new Prize { Name = "Máy Giặt", TotalQuantity = 15, RemainingQuantity = 15, BatchSize = 15 },
            new Prize { Name = "Nồi Chiên", TotalQuantity = 20, RemainingQuantity = 20, BatchSize = 10 },
            new Prize { Name = "Tivi", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
            new Prize { Name = "Tủ Lạnh", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Vali", TotalQuantity = 100, RemainingQuantity = 100, BatchSize = 10 },

    };
    public List<Prize> Afternoonprizes = new List<Prize>{
            new Prize { Name = "Xe Máy Honda Blade", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Wave Alpha", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Lead", TotalQuantity = 3, RemainingQuantity = 3, BatchSize = 3 },
            new Prize { Name = "Xe Máy Honda Vision", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
            new Prize { Name = "Máy Phát Điện Honda", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Giải đặc biệt", TotalQuantity = 4, RemainingQuantity = 4, BatchSize = 4 },
    };
    private string CurrentSheetName;
    public List<PlayerData> WonPlayer = new List<PlayerData>();
    public List<Prize> prizes;
    void Start()
    {
        string folderName = "Data";
        string baseDirectory = Path.GetDirectoryName(Application.dataPath);
        string filePath = Path.Combine(baseDirectory, folderName);
        if (!Directory.Exists(filePath))
        {
            Directory.CreateDirectory(filePath);
        }
        dataPath = filePath + "/Data.xlsx";    
        // dataPath = "D:\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Data.xlsx";
        dataPath = "D:\\STUDYING\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Data.xlsx";
        // public string prizePath = "D:\\STUDYING\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Prize.xlsx";
        // selectFileButton.onClick.AddListener(LoadExcel);
        LoadOrCreatePrizes(dataPath);
        prizes = Morningprizes;
        CurrentSheetName = "MorningPrizes";
        ReadExcel(dataPath);
        ReadWinPlayer(dataPath);
    }
    public void SwitchToMorning()
    {
        prizes = Morningprizes;
        CurrentSheetName = "MorningPrizes";
    }

    public void SwitchToAfternoon()
    {
        prizes = Afternoonprizes;
        CurrentSheetName = "AfternoonPrizes";
    }
    public List<PlayerData> playerDataList = new List<PlayerData>();
    public List<PlayerData> wonPlayer = new List<PlayerData>();
    public void LoadExcel()
    {
        var extensions = new[] { new ExtensionFilter("Excel Files", "xlsx", "xls") };
        string[] paths = StandaloneFileBrowser.OpenFilePanel("Open Excel File", "", extensions, false);

        if (paths.Length > 0 && !string.IsNullOrEmpty(paths[0]))
        {
            ReadExcel(paths[0]);
        }
    }

    private void LoadOrCreatePrizes(string filePath)
    {
        IWorkbook workbook;
        if (File.Exists(filePath))
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fs);
                ReadPrizesFromSheet(workbook, "MorningPrizes", Morningprizes);
                ReadPrizesFromSheet(workbook, "AfternoonPrizes", Afternoonprizes);
            }
        }
        else
        {
            workbook = new XSSFWorkbook();
            WritePrizesToSheet(workbook, "MorningPrizes", Morningprizes);
            WritePrizesToSheet(workbook, "AfternoonPrizes", Afternoonprizes);
            using (FileStream fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fs);
            }
        }
    }

    private void ReadPrizesFromSheet(IWorkbook workbook, string sheetName, List<Prize> prizeList)
    {
        ISheet sheet = workbook.GetSheet(sheetName);
        if (sheet == null) return;

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
    private void ReadWinPlayer(string path)
    {
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite))
        {
            IWorkbook workbook = new XSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheet("Win");

            if (sheet != null)
            {
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        PlayerData player = new PlayerData
                        {
                            id = int.Parse(row.GetCell(0).ToString()),
                            manhanvien = row.GetCell(1).ToString(),
                            hovaten = row.GetCell(2).ToString(),
                            phong = row.GetCell(3).ToString(),
                            note = row.GetCell(4).ToString() ?? string.Empty,
                            isWon = true,
                        };
                        wonPlayer.Add(player);
                    }
                }
            }
        }

    }
    private void WritePrizesToSheet(IWorkbook workbook, string sheetName, List<Prize> prizeList)
    {
        ISheet sheet = workbook.GetSheet(sheetName) ?? workbook.CreateSheet(sheetName);
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Name");
        sheet.GetRow(0).CreateCell(1).SetCellValue("TotalQuantity");
        sheet.GetRow(0).CreateCell(2).SetCellValue("RemainingQuantity");
        sheet.GetRow(0).CreateCell(3).SetCellValue("BatchSize");

        for (int i = 0; i < prizeList.Count; i++)
        {
            IRow row = sheet.GetRow(i + 1) ?? sheet.CreateRow(i + 1);
            row.CreateCell(0).SetCellValue(prizeList[i].Name);
            row.CreateCell(1).SetCellValue(prizeList[i].TotalQuantity);
            row.CreateCell(2).SetCellValue(prizeList[i].RemainingQuantity);
            row.CreateCell(3).SetCellValue(prizeList[i].BatchSize);
        }
    }
    public void SaveExcel(List<PlayerData> wonPlayer, string prizeName)
    {
        // Kiểm tra file có tồn tại hay không
        IWorkbook workbook;
        if (File.Exists(dataPath))
        {
            // Mở file hiện tại nếu có
            using (FileStream stream = new FileStream(dataPath, FileMode.Open, FileAccess.ReadWrite))
            {
                workbook = new XSSFWorkbook(stream);
            }
        }
        else
        {
            // Tạo file mới nếu chưa có
            workbook = new XSSFWorkbook();
        }

        // Kiểm tra sheet "Win" có tồn tại hay không
        ISheet sheet = workbook.GetSheet("Win") ?? workbook.CreateSheet("Win");

        // Nếu sheet mới tạo, thêm dòng tiêu đề
        if (sheet.PhysicalNumberOfRows == 0)
        {
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("ID");
            headerRow.CreateCell(1).SetCellValue("Mã Nhân Viên");
            headerRow.CreateCell(2).SetCellValue("Họ và Tên");
            headerRow.CreateCell(3).SetCellValue("Phòng");
            headerRow.CreateCell(4).SetCellValue("Ghi Chú");
            headerRow.CreateCell(5).SetCellValue("Giải Thưởng");
        }

        // Thêm dữ liệu vào sheet
        int lastRow = sheet.LastRowNum + 1;
        foreach (var player in wonPlayer)
        {
            IRow row = sheet.CreateRow(lastRow++);
            row.CreateCell(0).SetCellValue(player.id);
            row.CreateCell(1).SetCellValue(player.manhanvien);
            row.CreateCell(2).SetCellValue(player.hovaten);
            row.CreateCell(3).SetCellValue(player.phong);
            row.CreateCell(4).SetCellValue(player.note);
            row.CreateCell(5).SetCellValue(prizeName);
        }

        // Lưu lại file
        using (FileStream file = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
        {
            workbook.Write(file);
        }
    }
    public void WritePrizeRemainQuantity()
    {
        using (FileStream stream = new FileStream(dataPath, FileMode.Open, FileAccess.ReadWrite))
        {
            IWorkbook workbook = new XSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheet(CurrentSheetName);
            if (sheet != null)
            {
                for (int i = 0; i < prizes.Count; i++)
                {
                    IRow row = sheet.GetRow(i + 1);
                    if (row != null)
                    {
                        row.GetCell(2).SetCellValue(prizes[i].RemainingQuantity);
                    }
                }
            }
            using (FileStream writeStream = new FileStream(dataPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(writeStream);
            }
        }
    }
    private void ReadExcel(string dataPath)
    {
        playerDataList.Clear();

        using (FileStream stream = new FileStream(dataPath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = null;

            if (Path.GetExtension(dataPath) == ".xlsx")
            {
                workbook = new XSSFWorkbook(stream);
            }
            else if (Path.GetExtension(dataPath) == ".xls")
            {
                workbook = new HSSFWorkbook(stream);
            }

            if (workbook != null)
            {
                ISheet sheet = workbook.GetSheet("Data");
                for (int i = 2; i <= sheet.LastRowNum; i++) // Bỏ qua dòng tiêu đề
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    PlayerData data = new PlayerData
                    {
                        id = int.Parse(row.GetCell(0).ToString()),
                        manhanvien = row.GetCell(1).ToString(),
                        hovaten = row.GetCell(2).ToString(),
                        phong = row.GetCell(3).ToString(),
                        note = row.GetCell(4).ToString() ?? string.Empty,
                        isWon = false
                    };

                    playerDataList.Add(data);
                }
                Debug.Log(playerDataList.Count);
                Debug.Log($"Đã tải {playerDataList.Count} dòng dữ liệu.");
            }
            else
            {
                Debug.LogError("Không tìm thấy sheet thứ 2 trong file Excel.");
            }
        }
    }

}
