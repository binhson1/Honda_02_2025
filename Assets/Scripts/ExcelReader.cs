using UnityEngine;
using UnityEngine.UI;
using System.Collections.Generic;
using SFB;
using System.IO;
using TMPro;

public class ExcelReader : MonoBehaviour
{
    public Button selectFileButton;
    public TextMeshProUGUI filePathText;
    public TextMeshProUGUI playerListText;

    void Start()
    {
        ReadEmployeesFromExcel(Application.dataPath + "/ExcelTest/data.xlsx", 0);
        string filePath = "employees.xlsx";

        List<Employee> employees = ExcelReader.ReadEmployeesFromExcel(filePath);

    foreach (var emp in employees)
    {
        Console.WriteLine($"ID: {emp.Id}, Name: {emp.Name}, UpdatedAt: {emp.UpdatedAt}");
    }
    }


}