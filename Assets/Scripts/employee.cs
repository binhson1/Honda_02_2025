using System;

[Serializable]
public class Employee
{
    public int Id { get; set; }
    public string Code { get; set; }
    public string Name { get; set; }
    public string Department { get; set; }
    public string? Note { get; set; }
    public string? Gift { get; set; }
    public DateTime? UpdatedAt { get; set; }
    public DateTime? DeletedAt { get; set; }

    public Employee(int id, string code, string name, string department, string? note, string? gift, DateTime? updatedAt, DateTime? deletedAt)
    {
        this.Id = id;
        this.Code = code;
        this.Name = name;
        this.Department = department;
        this.Note = note;
        this.Gift = null;
        this.UpdatedAt = null;
        this.DeletedAt = null; 
    }

    public void UpdateGift(string gift)
    {
        if (string.IsNullOrEmpty(gift)) return;

        try
        {
            this.Gift = gift;
            this.UpdatedAt = DateTime.Now;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error setting gift: {ex.Message}");
            throw;
        }
    }

    public void Delete()
    {
        try
        {
            if (this.DeletedAt == null)
            {
                this.DeletedAt = DateTime.Now;
                Console.WriteLine($"Employee {this.Name} has been deleted.");
            }
            else
            {
                Console.WriteLine($"Employee {this.Name} was already deleted.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error deleting employee: {ex.Message}");
        }
    }
}

public class ExcelReader
{
    public static List<Employee> ReadEmployeesFromExcel(string filePath)
    {
        List<Employee> employees = new List<Employee>();

        using (FileStream file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
        {
              if (Path.GetExtension(path) == ".xls")
            {
                workbook = new HSSFWorkbook(stream); // Excel 97-2003
            }
            else
            {
                workbook = new XSSFWorkbook(stream); // Excel 2007+
            }

            ISheet sheet = workbook.GetSheetAt(0); 

            for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++) 
            {
                IRow row = sheet.GetRow(rowIndex);
                if (row == null) continue; // Bỏ qua dòng trống

                try
                {
                    Employee emp = new Employee
                    {
                        Id = (int)row.GetCell(0).NumericCellValue,
                        Code = row.GetCell(1).StringCellValue, 
                        Name = row.GetCell(2).StringCellValue, 
                        Department = row.GetCell(3).StringCellValue,
                        Note = row.GetCell(4)?.StringCellValue, 
                        Gift = row.GetCell(5)?.StringCellValue, 
                        UpdatedAt = GetDateValue(row.GetCell(6)),
                        DeletedAt = GetDateValue(row.GetCell(7)) 
                    };

                    employees.Add(emp);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi đọc dòng {rowIndex}: {ex.Message}");
                }
            }
        }

        return employees;
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