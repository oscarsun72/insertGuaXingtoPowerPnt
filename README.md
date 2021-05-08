# insertGuaXingtoPowerPnt
insert Yi Jing Gua Xing to PowerPnt 在Powerpoint簡報中插入《易經》卦形
已擴充到Word 也能用，且不止於卦形，還有任何圖片（主要是字圖，如小篆、行書、隸書、甲骨、金文……）均能用

【至於如例由字型轉字圖，可見此專案（rspo）：https://github.com/oscarsun72/PPTtoDoc_PowerPoint_xchg_Word 】

# 調整字圖大小（20210508） Resize Pictures and images
* 為了：在沒有必要以大圖插入的文件或投影片中，加速插入字圖的速度
* 效能：1萬多字當不出十分鐘即可調整完畢
* 應用範疇：可應用於所有圖形檔案
## 操作方式：
* 在主介面「調整字圖大小」右邊方塊先指定縮放比（以%(100)分為單位，如30%就指定「30」）
* 按住Ctrl鍵再按下滑鼠左鍵即可執行左邊清單中選中的字型字圖調整大小功能。原圖會移置原目錄的560x690資料夾下。此乃原圖的尺寸大小。
* 開始執行時，指定縮放比的方塊會變色，執行完成後則會恢復，且會循環播放系統提示樂音。要停止樂音，可關閉表單（本應用程式）、或將縮放比的方塊數值歸零。
* 無誤完成後會自行開啟操作的目錄（資料夾）以供檢視執行結果。
### 擬增功能：
- 在已經有「560x690」子目錄時，自行由此目錄中擷取原始圖檔來進行調整大小。並將最近的結果放此字圖的最頂層（母）目錄（top folder）中。

# 擬增功能
+ 由字型製作來源，判斷可能對應之字集。（如中國大陸與臺灣、日本製作者，字體對應即異）
+ 異體字對應功能（非楷體多是異體通用，如行草書、碑體字等，故凡係異體，均可流用）

以上期以增進字與字型的對應而減少特殊意義缺字的可能

# 以下實作經驗記錄
## 實作展示（[字型轉字圖開發過程結果演示](https://youtu.be/1FS9TZ0tWRk)）：
餘記錄片詳本人頻道：[https://www.youtube.com/c/%E5%AD%AB%E5%AE%88%E7%9C%9F](https://www.youtube.com/c/%E5%AD%AB%E5%AE%88%E7%9C%9F)

字型畢竟是要在本機電腦上安裝，才能正常顯示。若常在不同電腦間使用字型，又礙於可能對方電腦並不方便安裝字型，則不如將要展示的字型轉作字圖，來得方便、通用、跨平台，一了百了。
　　末學自己正在嘗試用C# 寫這樣的應用程式，已建立不少字型的字圖檔庫與插入字圖至pptx或docx檔的功能。不妨參考。感恩感恩　南無阿彌陀佛
ps. 實際測試，在Office內嵌字型，較以等量字圖插入Oiffe文件中，其檔案要冗肥十倍有餘。如內嵌字型後的ppt檔有30MB，同一個檔案以插入字圖的代換，則約只3MB。阿彌陀佛
建置字型字圖專案計畫：[https://github.com/oscarsun72/PPTtoDoc_PowerPoint_xchg_Word](https://github.com/oscarsun72/PPTtoDoc_PowerPoint_xchg_Word)

插入字圖專案機制：[https://github.com/oscarsun72/insertGuaXingtoPowerPnt](https://github.com/oscarsun72/insertGuaXingtoPowerPnt)

[https://free.com.tw/adobe-fonts-arphic-types/#more-83672](https://free.com.tw/adobe-fonts-arphic-types/?fb_comment_id=4258204694213311_4259613094072471)

[https://www.facebook.com/oscarsun72/posts/3689146467863127](https://www.facebook.com/oscarsun72/posts/3689146467863127)

