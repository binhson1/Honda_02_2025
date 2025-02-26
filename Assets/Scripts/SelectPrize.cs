using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class SelectPrize : MonoBehaviour
{
    public List<GameObject> batchprizes;
    public TMPro.TMP_Dropdown dropdownlistprize;
    public LoadData loadData;
    public int currentPrizeIndex;

    public List<GameObject> backGround;
    // Start is called before the first frame update
    void Start()
    {
        dropdownlistprize.onValueChanged.AddListener(delegate
        {
            DropdownValueChanged(dropdownlistprize);
        });
        Load();
    }

    public void Load()
    {
        List<string> prizeNames = new List<string>();
        for (int i = 0; i < loadData.prizes.Count; i++)
        {
            prizeNames.Add(loadData.prizes[i].Name + " : " + loadData.prizes[i].RemainingQuantity);
        }
        string batchPrizeName = loadData.prizes[0].BatchSize.ToString();
        dropdownlistprize.AddOptions(prizeNames);
        string prizeName = loadData.prizes[0].Name;
        for (int i = 0; i < batchprizes.Count; i++)
        {
            if (batchprizes[i].name == batchPrizeName)
            {
                batchprizes[i].SetActive(true);
                currentPrizeIndex = 0;
            }
            else
            {
                batchprizes[i].SetActive(false);
            }
        }
        for (int i = 0; i < backGround.Count; i++)
        {
            if (backGround[i].name == prizeName)
            {
                backGround[i].SetActive(true);
            }
            else
            {
                backGround[i].SetActive(false);
            }
        }
    }
    void DropdownValueChanged(TMPro.TMP_Dropdown change)
    {
        // check current dropdown value
        int prizeIndex = change.value;

        string batchPrizeName = loadData.prizes[prizeIndex].BatchSize.ToString();
        for (int i = 0; i < batchprizes.Count; i++)
        {
            if (batchprizes[i].name == batchPrizeName)
            {
                batchprizes[i].SetActive(true);
                currentPrizeIndex = change.value;
            }
            else
            {
                batchprizes[i].SetActive(false);
            }
        }
        string prizeName = loadData.prizes[prizeIndex].Name;
        for (int i = 0; i < backGround.Count; i++)
        {
            if (backGround[i].name == prizeName)
            {
                backGround[i].SetActive(true);
            }
            else
            {
                backGround[i].SetActive(false);
            }
        }
    }

    // Update is called once per frame
    void Update()
    {

    }
}
