using System;
using UnityEngine;
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
            Debug.Log($"Error setting gift: {ex.Message}");
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
                Debug.Log($"Employee {this.Name} has been deleted.");
            }
            else
            {
                Debug.Log($"Employee {this.Name} was already deleted.");
            }
        }
        catch (Exception ex)
        {
            Debug.Log($"Error deleting employee: {ex.Message}");
        }
    }
}