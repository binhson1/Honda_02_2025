using UnityEngine;
using TMPro;

public class DropdownSync : MonoBehaviour
{
    public TMP_Dropdown originalDropdown; // Dropdown gốc
    public TMP_Dropdown clonedDropdown;   // Dropdown cần sao chép

    void Start()
    {
        if (originalDropdown != null && clonedDropdown != null)
        {
            Invoke("StartClone", 1.2f);
        }
    }

    public void StartClone()
    {
        // Đồng bộ danh sách options (nếu chưa có)
        clonedDropdown.options = new System.Collections.Generic.List<TMP_Dropdown.OptionData>(originalDropdown.options);

        // Đặt giá trị mặc định
        clonedDropdown.value = originalDropdown.value;

        // Khi giá trị thay đổi, cập nhật dropdown clone
        originalDropdown.onValueChanged.AddListener(delegate { SyncDropdown(); });
        clonedDropdown.onValueChanged.AddListener(delegate { SyncValue(); });
    }
    public void SyncDropdown()
    {
        if (clonedDropdown != null)
        {
            clonedDropdown.value = originalDropdown.value;
            clonedDropdown.RefreshShownValue(); // Cập nhật UI
        }
    }
    public void CloneOption()
    {
        clonedDropdown.options = new System.Collections.Generic.List<TMP_Dropdown.OptionData>(originalDropdown.options);

    }
    public void SyncValue()
    {
        originalDropdown.value = clonedDropdown.value;
    }
}
