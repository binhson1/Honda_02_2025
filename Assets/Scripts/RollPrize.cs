using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

class RollPrize : MonoBehaviour
{
    public PlayerReader playerReader;
    public List<GameObject> players;
    // create a list to store player data
    public List<PlayerReader.PlayerData> playerDataList = new List<PlayerReader.PlayerData>();
    public void Roll()
    {
        for (int i = 0; i < players.Count; i++)
        {
            int randomIndex = Random.Range(0, playerReader.playerDataList.Count);
            string playerName = playerReader.playerDataList[randomIndex].hovaten;
            // get player object child text have name "Name"
            TMPro.TextMeshProUGUI playerNameText = players[i].transform.Find("Name").GetComponent<TMPro.TextMeshProUGUI>();
            playerNameText.text = playerName;
            string playerid = playerReader.playerDataList[randomIndex].manhanvien;
            // get player object child text have name "ID"
            TMPro.TextMeshProUGUI playerIDText = players[i].transform.Find("ID").GetComponent<TMPro.TextMeshProUGUI>();
            playerIDText.text = playerid;
            string playerphong = playerReader.playerDataList[randomIndex].phong;
            // get player object child text have name "Phong"
            TMPro.TextMeshProUGUI playerPhongText = players[i].transform.Find("Phong").GetComponent<TMPro.TextMeshProUGUI>();
            playerPhongText.text = playerphong;
            string playernote = playerReader.playerDataList[randomIndex].note;
            // get player object child text have name "Note"
            TMPro.TextMeshProUGUI playerNoteText = players[i].transform.Find("Note").GetComponent<TMPro.TextMeshProUGUI>();
            playerNoteText.text = playernote;
        }
    }

    public void SelectPlayer()
    {
        // check current gameobject that provoke this function
        GameObject selectedPlayer = UnityEngine.EventSystems.EventSystem.current.currentSelectedGameObject;
        Debug.Log(selectedPlayer);
    }
    public void ConfirmSave()
    {
        // Save to excel
    }
}