using System.Collections;
using System.Collections.Generic;
using UnityEngine;

public class NextPage : MonoBehaviour
{
    public List<GameObject> pages;
    // Start is called before the first frame update
    void Start()
    {
        pages[0].SetActive(true);
        for (int i = 1; i < pages.Count; i++)
        {
            pages[i].SetActive(false);
        }
    }

    public void Next()
    {
        for (int i = 0; i < pages.Count; i++)
        {
            if (pages[i].activeSelf)
            {
                pages[i].SetActive(false);
                if (i == pages.Count - 1)
                {
                    pages[0].SetActive(true);
                }
                else
                {
                    pages[i + 1].SetActive(true);
                }
                break;
            }
        }
    }
    // Update is called once per frame
    void Update()
    {

    }
}
