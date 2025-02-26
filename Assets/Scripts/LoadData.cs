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
    // public List<Prize> prizes = new List<Prize>{
    //         new Prize { Name = "Vali", TotalQuantity = 100, RemainingQuantity = 100, BatchSize = 10 },
    //         new Prize { Name = "Máy Giặt", TotalQuantity = 15, RemainingQuantity = 15, BatchSize = 10 },
    //         new Prize { Name = "Tivi", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
    //         new Prize { Name = "Nồi Chiên", TotalQuantity = 20, RemainingQuantity = 20, BatchSize = 10 },
    //         new Prize { Name = "Tủ Lạnh", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
    // };
    public List<Prize> prizes = new List<Prize>{
            new Prize { Name = "Xe Máy Honda Blade", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Wave Alpha", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Xe Máy Honda Lead", TotalQuantity = 3, RemainingQuantity = 3, BatchSize = 3 },
            new Prize { Name = "Xe Máy Honda Vision", TotalQuantity = 5, RemainingQuantity = 5, BatchSize = 5 },
            new Prize { Name = "Máy Phát Điện Honda", TotalQuantity = 10, RemainingQuantity = 10, BatchSize = 10 },
            new Prize { Name = "Giải đặc biệt", TotalQuantity = 4, RemainingQuantity = 4, BatchSize = 4 },
    };
    void Start()
    {
        // string baseDirectory = Path.GetDirectoryName(Application.dataPath);
        // string folderName = "Data";
        // string filePath = Path.Combine(baseDirectory, folderName);
        // if (!Directory.Exists(filePath))
        // {
        //     Directory.CreateDirectory(filePath);
        // }

        // wonPath = filePath + "Won.xlsx";
        // dataPath = filePath + "Data.xlsx";

        wonPath = "D:\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Won.xlsx";
        dataPath = "D:\\Unity\\Honda_02_2025\\Assets\\ExcelTest\\Data.xlsx";

        // selectFileButton.onClick.AddListener(LoadExcel);
        ReadExcel(dataPath);
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
    public void SaveExcel(List<PlayerData> wonPlayer, string prizeName)
    {
        // create file if not exist 
        if (!File.Exists(wonPath))
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("ID");
            headerRow.CreateCell(1).SetCellValue("Mã Nhân Viên");
            headerRow.CreateCell(2).SetCellValue("Họ và Tên");
            headerRow.CreateCell(3).SetCellValue("Phòng");
            headerRow.CreateCell(4).SetCellValue("Ghi Chú");
            headerRow.CreateCell(5).SetCellValue("Giải Thưởng");
            FileStream file = new FileStream(wonPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
        }
        // save data to excel 
        using (FileStream stream = new FileStream(wonPath, FileMode.Open, FileAccess.ReadWrite))
        {
            IWorkbook workbook = new XSSFWorkbook(stream);
            ISheet sheet = workbook.GetSheetAt(0);
            int lastRow = sheet.LastRowNum + 1;
            result.text = "";
            for (int i = 0; i < wonPlayer.Count; i++)
            {
                IRow row = sheet.CreateRow(lastRow + i);
                row.CreateCell(0).SetCellValue(wonPlayer[i].id);
                row.CreateCell(1).SetCellValue(wonPlayer[i].manhanvien);
                row.CreateCell(2).SetCellValue(wonPlayer[i].hovaten);
                row.CreateCell(3).SetCellValue(wonPlayer[i].phong);
                row.CreateCell(4).SetCellValue(wonPlayer[i].note);
                row.CreateCell(5).SetCellValue(prizeName);
                result.text += wonPlayer[i].hovaten;
                result.text += '\t';
            }
            stream.Close();
            FileStream file = new FileStream(wonPath, FileMode.Create);
            workbook.Write(file);
            file.Close();
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

            if (workbook != null && workbook.NumberOfSheets >= 2)
            {
                ISheet sheet = workbook.GetSheetAt(1); // Lấy sheet thứ 2

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

                Debug.Log($"Đã tải {playerDataList.Count} dòng dữ liệu.");
            }
            else
            {
                Debug.LogError("Không tìm thấy sheet thứ 2 trong file Excel.");
            }
        }
    }

}
