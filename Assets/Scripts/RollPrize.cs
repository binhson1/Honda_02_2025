using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using TMPro;

class RollPrize : MonoBehaviour
{
    public LoadData loadData;
    public List<GameObject> players;
    // create a list to store player data
    public List<LoadData.PlayerData> playingPlayer = new List<LoadData.PlayerData>();
    
    public List<LoadData.PlayerData> wonPlayer = new List<LoadData.PlayerData>();
    public SelectPrize selectPrize;

    public TextMeshProUGUI currentPrizeLeft;

    void Start()
    {
        
    }
    public void Roll()
    {
        if(loadData.playerDataList.Count == 0)
        {
            return;
        }
        if(playingPlayer.Count == 0)
        {
            for(int i = 0; i < players.Count; i++)
            {
                playingPlayer.Add(new LoadData.PlayerData());
            }
        }
        SavePlayer();
        for(int i = 0; i < players.Count; i++)
        {
            if (!playingPlayer[i].isWon)
            { 
                int random = Random.Range(0, loadData.playerDataList.Count);
                players[i].transform.Find("Text").GetComponent<TMPro.TextMeshProUGUI>().text = loadData.playerDataList[random].manhanvien + " - " + loadData.playerDataList[random].hovaten + " - " + loadData.playerDataList[random].phong + " - " + loadData.playerDataList[random].note;
                playingPlayer[i] = loadData.playerDataList[random];
                playingPlayer[i].isWon = true;
                wonPlayer.Add(playingPlayer[i]);
            }
        }
    }

    public void SavePlayer()
    {
        // check if wonPlayer.count == players.count
        if(wonPlayer.Count == players.Count)
        {
            // save to excel
            loadData.SaveExcel(wonPlayer, loadData.prizes[selectPrize.currentPrizeIndex].Name);
            // reset all player            
            for(int i = 0; i < players.Count; i++)
            {
                players[i].transform.Find("Text").GetComponent<TMPro.TextMeshProUGUI>().text = "";
                playingPlayer[i] = new LoadData.PlayerData();                
            }
            wonPlayer.Clear();
            loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity -= loadData.prizes[selectPrize.currentPrizeIndex].BatchSize;
            currentPrizeLeft.text = loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString();
        }
    }
    public void SelectPlayer()
    {
        // check current gameobject that provoke this function
        GameObject selectedPlayer = UnityEngine.EventSystems.EventSystem.current.currentSelectedGameObject;
        Debug.Log(selectedPlayer);
        // check who is this player in the gameobject list
        for(int i = 0; i < players.Count; i++)
        {
            if(players[i].name == selectedPlayer.name)
            {
                // check if this player is already won
                if(playingPlayer[i].isWon)
                {
                    // remove from won list
                    playingPlayer[i].isWon = false;
                    wonPlayer.Remove(playingPlayer[i]);
                    players[i].transform.Find("Text").GetComponent<TMPro.TextMeshProUGUI>().text = "";                    
                }
                else
                {
                    // add to won list
                    playingPlayer[i].isWon = true;
                    wonPlayer.Add(playingPlayer[i]);
                    players[i].transform.Find("Text").GetComponent<TMPro.TextMeshProUGUI>().text = playingPlayer[i].manhanvien + " - " + playingPlayer[i].hovaten + " - " + playingPlayer[i].phong + " - " + playingPlayer[i].note;
                }
            }
        }

    }
    public void ConfirmSave()
    {
        // Save to excel
    }
}