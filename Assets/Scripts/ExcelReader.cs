using UnityEngine;
using UnityEngine.UI;
using System.Collections.Generic;
using SFB; // StandaloneFileBrowser
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using System.IO;
using TMPro;

public class ExcelReader : MonoBehaviour
{
    public Button selectFileButton;
    public TextMeshProUGUI filePathText;
    public TextMeshProUGUI playerListText;

    void Start()
    {
        selectFileButton.onClick.AddListener(OpenFileExplorer);
    }

    void OpenFileExplorer()
    {
        var paths = StandaloneFileBrowser.OpenFilePanel("Select Excel File", "", new[] {
            new ExtensionFilter("Excel Files", "xlsx", "xls")
        }, false);

        if (paths.Length > 0 && !string.IsNullOrEmpty(paths[0]))
        {
            filePathText.text = "Selected File: " + paths[0];
            ReadExcel(paths[0]);
        }
    }

    void ReadExcel(string path)
    {
        playerListText.text = "Loading players...\n";

        IWorkbook workbook;
        using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read))
        {
            if (Path.GetExtension(path) == ".xls")
                workbook = new HSSFWorkbook(stream); // Excel 97-2003
            else
                workbook = new XSSFWorkbook(stream); // Excel 2007+

            ISheet sheet = workbook.GetSheetAt(0); // Lấy sheet đầu tiên

            List<string> players = new List<string>();

            for (int i = 0; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    string playerName = row.GetCell(0)?.ToString(); // Lấy dữ liệu cột đầu tiên
                    if (!string.IsNullOrEmpty(playerName))
                    {
                        players.Add(playerName);
                    }
                }
            }

            playerListText.text = "Players List:\n" + string.Join("\n", players);
        }
    }
}
