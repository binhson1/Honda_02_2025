using System;
using System.Collections.Generic;
using System.IO;
using UnityEngine;
using UnityEngine.UI;
using TMPro;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;


public class ExcelReader : MonoBehaviour
{
    public Button selectFileButton;
    public TextMeshProUGUI filePathText;
    public TextMeshProUGUI playerListText;

    void Start()
    {
        string filePath = Application.dataPath + "/ExcelTest/data.xlsx";

        List<Employee> employees = ReadEmployeesFromExcel(filePath);
        List<Gift> gifts = ReadGiftsFromExcel(filePath);

        foreach (var emp in employees)
        {
            Debug.Log($"ID: {emp.Id}, Name: {emp.Name}, UpdatedAt: {emp.UpdatedAt}");
        }

        foreach (var gift in gifts)
        {
            Debug.Log($"ID: {gift.Id}, Name: {gift.Name}, Quantity: {gift.Quantity} , Remaining: {gift.Remaining}, UpdatedAt: {gift.UpdatedAt}");
        }
    }


    public static List<Employee> ReadEmployeesFromExcel(string filePath)
    {

        List<Employee> employees = new List<Employee>();

        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook;
            if (Path.GetExtension(filePath) == ".xls")
            {
                workbook = new HSSFWorkbook(file); // Excel 97-2003
            }
            else
            {
                workbook = new XSSFWorkbook(file); // Excel 2007+
            }

            ISheet sheet = workbook.GetSheetAt(0);

            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue; // Bỏ qua dòng trống

                try
                {
                    Employee emp = new Employee(
                        (int)row.GetCell(0).NumericCellValue,
                        row.GetCell(1).StringCellValue,
                        row.GetCell(2).StringCellValue,
                        row.GetCell(3).StringCellValue,
                        row.GetCell(4)?.StringCellValue,
                        row.GetCell(5)?.StringCellValue,
                        GetDateValue(row.GetCell(6)),
                        GetDateValue(row.GetCell(7))
                    );

                    employees.Add(emp);
                }
                catch (Exception ex)
                {
                    Debug.Log($"Lỗi đọc employee dòng {rowIndex}: {ex.Message}");
                }
            }
        }
        return employees;
    }
    public static List<Gift> ReadGiftsFromExcel(string filePath)
    {

        List<Gift> gifts = new List<Gift>();

        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
            IWorkbook workbook;
            if (Path.GetExtension(filePath) == ".xls")
            {
                workbook = new HSSFWorkbook(file); // Excel 97-2003
            }
            else
            {
                workbook = new XSSFWorkbook(file); // Excel 2007+
            }

            ISheet sheet = workbook.GetSheetAt(1);

            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue; // Bỏ qua dòng trống

                try
                {
                    Gift gift = new Gift(
                        (int)row.GetCell(0).NumericCellValue,
                        row.GetCell(1).StringCellValue,
                        (int)row.GetCell(2).NumericCellValue,
                        row.GetCell(3) == null ? (int)row.GetCell(2).NumericCellValue : (int)row.GetCell(3).NumericCellValue,
                        GetDateValue(row.GetCell(4))
                    );

                    gifts.Add(gift);
                }
                catch (Exception ex)
                {
                    Debug.Log($"Lỗi gift đọc dòng {rowIndex}: {ex.Message}");
                }
            }
        }
        return gifts;
    }
    private static DateTime? GetDateValue(ICell cell)
    {
        if (cell == null || cell.CellType == CellType.Blank)
            return null;

        if (cell.CellType == CellType.Numeric && DateUtil.IsCellDateFormatted(cell))
            return cell.DateCellValue;

        return null;
    }

}