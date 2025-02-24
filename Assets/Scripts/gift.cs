using System;
using UnityEngine;
public class Gift
{
    public int Id { get; set; }
    public string Name { get; set; }
    public int Quantity { get; set; }
    public int? Remaining { get; set; }
    public DateTime? UpdatedAt { get; set; }

    public Gift(int id, string name, int quantity, int? remaining, DateTime? updatedAt)
    {
        this.Id = id;
        this.Name = name;
        this.Quantity = quantity;
        this.Remaining = remaining ?? quantity;
        this.UpdatedAt = null;
    }

    public void UpdateQuantity(int quantity)
    {
        if (quantity <= 0)
        {
            Debug.LogError("Quantity must be greater than 0.");
            throw new ArgumentException("Quantity must be greater than 0.", nameof(quantity));
        }

        try
        {
            this.Quantity = quantity;
            this.UpdatedAt = DateTime.Now;
        }
        catch (Exception ex)
        {
            Debug.LogError($"Error setting gift: {ex.Message}");
            throw;
        }
    }
}