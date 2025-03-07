using System.Collections.Generic;
using UnityEngine;
using UnityEngine.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using TMPro;
using System.Collections;
using System.Linq;

class RollPrize : MonoBehaviour
{
    public LoadData loadData;
    private int currentPage = 0;
    private const int pageSize = 20;
    public List<GameObject> players;
    public List<Sprite> sprites;
    // create a list to store player data
    public List<LoadData.PlayerData> playingPlayer = new List<LoadData.PlayerData>();

    public List<LoadData.PlayerData> wonPlayer = new List<LoadData.PlayerData>(); // list này dùng để lưu người đã điểm danh rồi hoặc chưa điểm danh

    public List<LoadData.PlayerData> resultList = new List<LoadData.PlayerData>();
    public SelectPrize selectPrize;
    public TextMeshProUGUI buttonText;
    public TextMeshProUGUI currentPrizeLeft;
    public float fadeDuration = 5f; // Thời gian fade (giây)
    private bool isPlaying = false;
    private bool isStop = false;
    public List<TextMeshProUGUI> resultTextTMP;
    public List<TextMeshProUGUI> bigResultTextTMP;
    public Sprite redSprite;
    public Sprite greenSprite;
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
            if (!isStop)// quay tiếp không dừng lại cho trường hợp buổi sáng.
            {
                List<int> randomList = new List<int>();
                for (int i = 0; i < players.Count; i++)
                {
                    if (!playingPlayer[i].isWon)
                    {
                        int random = Random.Range(0, loadData.playerDataList.Count);
                        while (loadData.wonPlayer.Contains(loadData.playerDataList[random]) || randomList.Contains(random) || loadData.playerDataList[random].isWon == true)
                        {
                            random = Random.Range(0, loadData.playerDataList.Count);
                        }
                        players[i].transform.Find("Anim").gameObject.SetActive(true);
                        StartCoroutine(Fade(players[i].transform.Find("Anim").gameObject, true));
                        players[i].transform.Find("Text (TMP)").gameObject.SetActive(false);
                        string note = string.IsNullOrEmpty(loadData.playerDataList[random].note) ? " " : (" - " + loadData.playerDataList[random].note);
                        players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = loadData.playerDataList[random].manhanvien + " - " + loadData.playerDataList[random].hovaten + " - " + loadData.playerDataList[random].phong + note;
                        // change color
                        // Image imageColor = players[i].transform.Find("ImageBG").GetComponent<Image>();
                        // imageColor.color = new Color(255, 255, 255);
                        playingPlayer[i] = loadData.playerDataList[random];
                        playingPlayer[i].isWon = true;
                        wonPlayer[i] = playingPlayer[i];
                        randomList.Add(random);
                    }
                    buttonText.text = "STOP (Dừng Quay)";
                }
                isPlaying = true;// là cho cái animations;
                isStop = true;
            }
            else
            {
                SavePlayer();
            }
        }
        else
        {
            for (int i = 0; i < players.Count; i++)
            {
                players[i].transform.Find("Anim").gameObject.SetActive(false);
                players[i].transform.Find("Text (TMP)").gameObject.SetActive(true);
            }
            isPlaying = false;
            // buttonText.text = "ROLL";
            buttonText.text = "SAVE (Lưu Danh Sách)";
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
        buttonText.text = "START (Bắt Đầu Quay)";
        bool allWon = true;
        List<LoadData.PlayerData> savePlayer = new List<LoadData.PlayerData>();
        for (int i = 0; i < players.Count; i++)
        {
            if (playingPlayer[i].isWon)
            {
                if (wonPlayer[i] != null)
                {
                    wonPlayer[i].isWon = false;
                    loadData.playerDataList.Remove(wonPlayer[i]);
                    wonPlayer[i].prizeName = loadData.prizes[selectPrize.currentPrizeIndex].Name;
                    wonPlayer[i].isWon = true;
                    loadData.wonPlayer.Add(wonPlayer[i]);
                    resultList.Add(wonPlayer[i]);
                    // string note = string.IsNullOrEmpty(wonPlayer[i].note) ? " " : (" - " + wonPlayer[i].note);
                    Image image = players[i].transform.Find("ImageBG").GetComponent<Image>();
                    image.sprite = redSprite;
                    // image.color = new Color(255, 255, 255);
                    // if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 10)
                    // {
                    //     bigResultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + wonPlayer[i].manhanvien + " - " + wonPlayer[i].hovaten + " - " + wonPlayer[i].phong + note;
                    // }
                    // else
                    // {
                    //     resultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + wonPlayer[i].manhanvien + " - " + wonPlayer[i].hovaten + " - " + wonPlayer[i].phong + note;
                    // }
                    savePlayer.Add(wonPlayer[i]);
                    wonPlayer[i] = null;
                }
            }
            else
            {
                allWon = false;
            }
        }
        if (!allWon)
        {
            for (int i = 0; i < resultList.Count; i++)
            {
                string note = string.IsNullOrEmpty(resultList[i].note) ? " " : (" - " + resultList[i].note);
                if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 10)
                {
                    bigResultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + resultList[i].manhanvien + " - " + resultList[i].hovaten + " - " + resultList[i].phong + note;
                }
                else
                {
                    resultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + resultList[i].manhanvien + " - " + resultList[i].hovaten + " - " + resultList[i].phong + note;
                }
            }
        }
        loadData.SaveExcel(savePlayer, loadData.prizes[selectPrize.currentPrizeIndex].Name);
        isStop = false;
        if (loadData.prizes[selectPrize.currentPrizeIndex].Name != "Giải thưởng bất ngờ")
        {
            loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity -= savePlayer.Count;
            currentPrizeLeft.text = "Còn lại : " + loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString() + " lượt";
            loadData.WritePrizeRemainQuantity();
        }
        else
        {
            currentPrizeLeft.text = resultList.Count().ToString();
        }
        if (allWon)
        {
            for (int i = 0; i < players.Count; i++)
            {
                playingPlayer[i].isWon = false;
                // change Image Sprite 
                if (loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity > 0 && loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity >= 20)
                {
                    players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = "";
                    Image image = players[i].transform.Find("Image").GetComponent<Image>();
                    image.sprite = sprites[i + (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity - loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity)];
                }
            }
            if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 100 || loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 20)
            {
                ShowMultiResult();
            }
            if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 1)
            {
                for (int i = 0; i < resultList.Count; i++)
                {
                    string note = string.IsNullOrEmpty(resultList[i].note) ? " " : (" - " + resultList[i].note);
                    resultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + resultList[i].manhanvien + " - " + resultList[i].hovaten + " - " + resultList[i].phong + note;
                }
            }
        }
    }

    public void ShowHistoryResult()
    {
        if (loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity == 0 && gameObject.name == loadData.prizes[selectPrize.currentPrizeIndex].BatchSize.ToString())
        {
            List<LoadData.PlayerData> showList = loadData.wonPlayer.Where(p => p.prizeName == loadData.prizes[selectPrize.currentPrizeIndex].Name).ToList();
            if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity != 100)
            {
                for (int i = 0; i < showList.Count; i++)
                {
                    string note = string.IsNullOrEmpty(showList[i].note) ? " " : (" - " + showList[i].note);

                    if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity != 10)
                    {
                        resultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + showList[i].manhanvien + " - " + showList[i].hovaten + " - " + showList[i].phong + note;
                    }
                    else
                    {
                        bigResultTextTMP[i].text = "(" + (i + 1) + ")" + " - " + showList[i].manhanvien + " - " + showList[i].hovaten + " - " + showList[i].phong + note;
                    }
                }
            }
            else
            {
                resultList = showList;
                currentPage = 0;
                ShowMultiResult();
            }
        }
    }

    public void ShowMultiResult()
    {
        if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 20)
        {
            for (int i = 0; i < resultList.Count; i++)
            {
                if (resultList[i] != null)
                {
                    string note = string.IsNullOrEmpty(resultList[i].note) ? " " : (" - " + resultList[i].note);
                    resultTextTMP[i].text = $"({i + 1}) - {resultList[i].manhanvien} - {resultList[i].hovaten} - {resultList[i].phong}{note}";
                }
            }
            return;
        }
        int startIndex = currentPage * pageSize;
        int endIndex = Mathf.Min(startIndex + pageSize, resultList.Count);

        for (int i = 0; i < pageSize; i++)
        {
            if (startIndex + i < resultList.Count && resultList[startIndex + i] != null)
            {
                string note = string.IsNullOrEmpty(resultList[startIndex + i].note) ? " " : (" - " + resultList[startIndex + i].note);
                resultTextTMP[i].text = $"({startIndex + i + 1}) - {resultList[startIndex + i].manhanvien} - {resultList[startIndex + i].hovaten} - {resultList[startIndex + i].phong}{note}";
            }
            else
            {
                resultTextTMP[i].text = ""; // Xóa nội dung nếu không có dữ liệu
            }
        }
    }

    public void NextPage()
    {
        if ((currentPage + 1) * pageSize < resultList.Count)
        {
            currentPage++;
            ShowMultiResult();
        }
    }

    public void PreviousPage()
    {
        if (currentPage > 0)
        {
            currentPage--;
            ShowMultiResult();
        }
    }
    public void clearText()
    {
        // currentPrizeLeft.text = loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString();
        if (loadData.prizes[selectPrize.currentPrizeIndex].TotalQuantity == 1)
        {
            currentPrizeLeft.text = resultList.Count.ToString();
        }
        else
        {
            currentPrizeLeft.text = "Còn lại : " + loadData.prizes[selectPrize.currentPrizeIndex].RemainingQuantity.ToString() + " lượt";
        }
        for (int i = 0; i < players.Count; i++)
        {
            players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = "";
            Image imagebg = players[i].transform.Find("ImageBG").GetComponent<Image>();
            imagebg.sprite = redSprite;
            // imagebg.color = new Color(255, 255, 255);
            SpriteRenderer spriteRenderer = players[i].transform.Find("Anim").GetComponent<SpriteRenderer>();
            if (players.Count == 10)
            {
                Image image = players[i].transform.Find("Image").GetComponent<Image>();
                image.sprite = sprites[i];
            }
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
        for (int i = 0; i < bigResultTextTMP.Count; i++)
        {
            bigResultTextTMP[i].text = "";
        }
        resultList.Clear();
        resultList = new List<LoadData.PlayerData>();
        buttonText.text = "START (Bắt Đầu Quay)";
    }
    public void SelectPlayer()
    {
        if (playingPlayer == null)
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
                    // change color
                    Image image = players[i].transform.Find("ImageBG").GetComponent<Image>();
                    image.sprite = greenSprite;
                    // image.color = new Color(0, 255, 255);
                }
                else
                {
                    // add to won list
                    playingPlayer[i].isWon = true;
                    wonPlayer[i].isWon = true;
                    players[i].transform.Find("Text (TMP)").GetComponent<TMPro.TextMeshProUGUI>().text = playingPlayer[i].manhanvien + " - " + playingPlayer[i].hovaten + " - " + playingPlayer[i].phong + " - " + playingPlayer[i].note;
                    // change color
                    Image image = players[i].transform.Find("ImageBG").GetComponent<Image>();
                    image.sprite = redSprite;
                    // image.color = new Color(255, 255, 255);
                }
            }
        }
    }
    void Update()
    {

    }
}