using System.Collections.Generic;
using UnityEngine;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using TMPro;
using System.Collections;
using System.Linq;

class RollPrize : MonoBehaviour
{
    public LoadData loadData;
    public List<GameObject> players;
    // create a list to store player data
    public List<LoadData.PlayerData> playingPlayer = new List<LoadData.PlayerData>();

    public List<LoadData.PlayerData> wonPlayer = new List<LoadData.PlayerData>();

    public List<LoadData.PlayerData> resultList = new List<LoadData.PlayerData>();
    public SelectPrize selectPrize;
    public TextMeshProUGUI buttonText;
    public TextMeshProUGUI currentPrizeLeft;
    public float fadeDuration = 5f; // Thời gian fade (giây)

    public List<TextMeshProUGUI> resultTMP = new List<TextMeshProUGUI>();
    private bool isPlaying = false;
    private bool isStop = false;
    public List<TextMeshProUGUI> resultTextTMP;
    // list of people that won 
    void Start()
    {

    }
    public void Roll()
    {
        if (loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity <= 0)
        {
            return;
        }
        if (!isPlaying)
        {            
            if (playingPlayer.Count == 0)
            {
                for (int i = 0; i < players.Count; i++)
                {
                    playingPlayer.Add(new LoadData.PlayerData());
                    wonPlayer.Add(new LoadData.PlayerData());
                }
            }
            SavePlayer();
            if(!isStop)
            {
                List<int> randomList = new List<int>();
                for (int i = 0; i < players.Count; i++)
                {
                    if (!playingPlayer[i].isWon)
                    {                    
                        int random = Random.Range(0, loadData.playerDataList.Count);
                        while (loadData.wonPlayer.Contains(loadData.playerDataList[random]) || randomList.Contains(random))
                        {
                            random = Random.Range(0, loadData.playerDataList.Count);
                        }
                        players[i].transform.Find("Anim").gameObject.SetActive(true);
                        StartCoroutine(Fade(players[i].transform.Find("Anim").gameObject, true));
                        players[i].transform.Find("Text (TMP)").gameObject.SetActive(false);
                        string note = string.IsNullOrEmpty(loadData.playerDataList[random].note) ? " " : (" - " + loadData.playerDataList[random].note);
                        players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = loadData.playerDataList[random].manhanvien + " - " + loadData.playerDataList[random].hovaten + " - " + loadData.playerDataList[random].phong + note;
                        playingPlayer[i] = loadData.playerDataList[random];
                        playingPlayer[i].isWon = true;
                        wonPlayer[i] = playingPlayer[i];
                        randomList.Add(random);
                    }
                    buttonText.text = "STOP";
                }
            } 
            else
            {
                isStop = false;
            }           
            isPlaying = true;
        }
        else
        {
            for (int i = 0; i < players.Count; i++)
            {
                players[i].transform.Find("Anim").gameObject.SetActive(false);
                players[i].transform.Find("Text (TMP)").gameObject.SetActive(true);
            }
            isPlaying = false;
            buttonText.text = "ROLL";
        }
    }   
    
    // Gọi để làm hiện dần
    public void StartFadeIn(GameObject target)
    {
        StartCoroutine(Fade(target, true));
    }

    // Gọi để làm ẩn dần
    public void StartFadeOut(GameObject target)
    {
        StartCoroutine(Fade(target, false));
    }

    private IEnumerator Fade(GameObject target, bool fadeIn)
    {
        SpriteRenderer sr = target.GetComponent<SpriteRenderer>();
        if (sr == null) yield break;

        float startAlpha = fadeIn ? 0f : 1f;
        float endAlpha = fadeIn ? 1f : 0f;

        Color color = sr.color;
        color.a = startAlpha;
        sr.color = color;

        float elapsedTime = 0f;

        while (elapsedTime < fadeDuration)
        {
            color.a = Mathf.Lerp(startAlpha, endAlpha, elapsedTime / fadeDuration);
            sr.color = color;
            elapsedTime += Time.deltaTime;
            yield return null;
        }

        // Đảm bảo giá trị alpha đạt chính xác 0 hoặc 1
        color.a = endAlpha;
        sr.color = color;
    }
    public void SavePlayer()
    {
        bool allWon = true;
        List<LoadData.PlayerData> savePlayer = new List<LoadData.PlayerData>();                
        for (int i = 0; i < players.Count; i++)
        {
            if(playingPlayer[i].isWon)
            {
                if(wonPlayer[i] != null )
                {
                    loadData.wonPlayer.Add(wonPlayer[i]);
                    resultList.Add(wonPlayer[i]);
                    string note = string.IsNullOrEmpty(wonPlayer[i].note) ? " " : (" - " + wonPlayer[i].note);
                    resultTextTMP[i].text = wonPlayer[i].manhanvien + " - " + wonPlayer[i].hovaten + " - " + wonPlayer[i].phong + note;
                    savePlayer.Add(wonPlayer[i]);
                    wonPlayer[i] = null;
                }                                                     
            }
            else
            {
                allWon = false;
            }
        }
        loadData.SaveExcel(savePlayer, loadData.prizes[selectPrize.currentPrizeIndex].Name);            
        if(allWon)
        {
            for (int i = 0; i < players.Count; i++)
            {                
                playingPlayer[i].isWon = false;
                players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = "";
            }
            isStop = true;
        }    
        if(loadData.prizes[selectPrize.currentPrizeIndex].BatchSize == savePlayer.Count)        
        {
            // isStop = true;    
        }      
        loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity -= savePlayer.Count;
        currentPrizeLeft.text = loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString();
        loadData.WritePrizeRemainQuantity();        
    }
    public void clearText()
    {
        currentPrizeLeft.text = loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString();
        for (int i = 0; i < players.Count; i++)
        {
            players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = "";
            SpriteRenderer spriteRenderer = players[i].transform.Find("Anim").GetComponent<SpriteRenderer>();

            // Lấy màu hiện tại
            Color color = spriteRenderer.color;

            // Set giá trị alpha thành 0
            color.a = 0f;

            // Gán lại màu sắc đã thay đổi vào SpriteRenderer
            spriteRenderer.color = color;
        }
        for (int i = 0; i < resultTextTMP.Count; i++)
        {
            resultTextTMP[i].text = "";
        }
        buttonText.text = "ROLL";
    }
    public void SelectPlayer()
    {
        if(playingPlayer == null)        
        {
            return;
        }
        // check current gameobject that provoke this function
        GameObject selectedPlayer = UnityEngine.EventSystems.EventSystem.current.currentSelectedGameObject;
        // Debug.Log(selectedPlayer);
        // check who is this player in the gameobject list
        for (int i = 0; i < players.Count; i++)
        {
            if (players[i].name == selectedPlayer.name && wonPlayer[i] != null)
            {
                // check if this player is already won
                if (playingPlayer[i].isWon)
                {
                    // remove from won list
                    playingPlayer[i].isWon = false;                    
                    wonPlayer[i].isWon = false;
                    players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = "";
                }
                else
                {
                    // add to won list
                    playingPlayer[i].isWon = true;
                    wonPlayer[i].isWon = true;
                    players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = playingPlayer[i].manhanvien + " - " + playingPlayer[i].hovaten + " - " + playingPlayer[i].phong + " - " + playingPlayer[i].note;
                }
            }
        }
    }
    public void ShowResult()
    {
        for (int i = 0; i < resultTMP.Count; i++)
        {
            resultTMP[i].text = resultList[i].manhanvien + " " + resultList[i].hovaten;
        }
    }
    void Update()
    {

    }
}