using UnityEngine;
using System.Collections.Generic;
using System.IO;
using SFB;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using UnityEngine.UI;

public class PlayerReader : MonoBehaviour
{
    public Button selectFileButton;

    [System.Serializable]
    public class PlayerData
    {
        public int id;
        public string manhanvien;
        public string hovaten;
        public string phong;
        public string note;
    }
    void Start()
    {
        selectFileButton.onClick.AddListener(LoadExcel);
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

    private void ReadExcel(string path)
    {
        playerDataList.Clear();

        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook = null;

            if (Path.GetExtension(path) == ".xlsx")
            {
                workbook = new XSSFWorkbook(stream);
            }
            else if (Path.GetExtension(path) == ".xls")
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
                        note = row.GetCell(4).ToString() ?? string.Empty
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
