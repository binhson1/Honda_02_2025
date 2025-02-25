using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class SelectPrize : MonoBehaviour
{
    public List<GameObject> prizes;
    public TMPro.TMP_Dropdown dropdownlistprize;
    public int prizeIndex;
    // Start is called before the first frame update
    void Start()
    {
        dropdownlistprize.onValueChanged.AddListener(delegate
        {
            DropdownValueChanged(dropdownlistprize);
        });
        for (int i = 0; i < prizes.Count; i++)
        {
            if (i == 0)
            {
                prizes[i].SetActive(true);
            }
            else
            {
                prizes[i].SetActive(false);
            }
        }
    }

    void DropdownValueChanged(TMPro.TMP_Dropdown change)
    {
        for (int i = 0; i < prizes.Count; i++)
        {
            prizes[i].SetActive(false);
        }
        prizes[change.value].SetActive(true);
    }

    // Update is called once per frame
    void Update()
    {

    }
}
