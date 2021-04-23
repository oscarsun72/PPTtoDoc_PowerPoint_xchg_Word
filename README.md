# PPTtoDoc_PowerPoint_xchg_Word_DocSwitchPPT
 pptx and docx switch .ppt and doc, PowerPoint and Word automation exchange and interoperation
 
 PowerPoint ppt pptx 檔與 Word doc docx 檔之間的互動操作程式設計與開發；也藉以嫻熟自己這方面的知識與技術
 
## 【本計畫之緣起】：

菩薩慈悲：看來，目前最高端的 Word docx 文件轉 成PowerPoint pptx 的技巧，就屬[豬式繪設教學頻道：《【超實用】將WORD檔案轉換成POWERPOINT檔》此片](https://youtu.be/3YMx5zAsqq0)所介紹的，如果不用程式，不寫程式的話。可見本計畫是有實作的必要了。感恩感恩　南無阿彌陀佛


# 以下實作經驗記錄
## 實作展示（[開發過程結果演示](https://youtu.be/1FS9TZ0tWRk)）：
餘記錄片詳本人頻道：[https://www.youtube.com/c/%E5%AD%AB%E5%AE%88%E7%9C%9F](https://www.youtube.com/c/%E5%AD%AB%E5%AE%88%E7%9C%9F)

字型畢竟是要在本機電腦上安裝，才能正常顯示。若常在不同電腦間使用字型，又礙於可能對方電腦並不方便安裝字型，則不如將要展示的字型轉作字圖，來得方便、通用、跨平台，一了百了。
　　末學自己正在嘗試用C# 寫這樣的應用程式，已建立不少字型的字圖檔庫與插入字圖至pptx或docx檔的功能。不妨參考。感恩感恩　南無阿彌陀佛
ps. 實際測試，在Office內嵌字型，較以等量字圖插入Oiffe文件中，其檔案要冗肥十倍有餘。如內嵌字型後的ppt檔有30MB，同一個檔案以插入字圖的代換，則約只3MB。阿彌陀佛
建置字型字圖專案計畫：[https://github.com/oscarsun72/PPTtoDoc_PowerPoint_xchg_Word](https://github.com/oscarsun72/PPTtoDoc_PowerPoint_xchg_Word)

插入字圖專案機制：[https://github.com/oscarsun72/insertGuaXingtoPowerPnt](https://github.com/oscarsun72/insertGuaXingtoPowerPnt)

[https://free.com.tw/adobe-fonts-arphic-types/#more-83672](https://free.com.tw/adobe-fonts-arphic-types/?fb_comment_id=4258204694213311_4259613094072471)

[https://www.facebook.com/oscarsun72/posts/3689146467863127](https://www.facebook.com/oscarsun72/posts/3689146467863127)

## 【字型轉字圖】各字型所需大小設定恐有不同，實測後略列如下（或與畫布大小與比例也有關）：
### 【目前專注】：
VBA已達極限，準擬以C#統合之。期能做到一鍵搞定。雖然VBA可能也可以，但以C#或VB.Net寫成可執行棕檔畢竟仍較方便，不必受制於Office介面，得以獨立作業及後續擴展。感恩感恩　南無阿彌陀佛


### 字型名稱=字型大小
+ Adobe 仿宋 Std R=450
+ 教育部隸書=450（右緣已被吃到些了！）
+ 華康行書體=410
+ 文鼎行楷L=450
+ 餘詳\各字型檔(不含缺字)相關\下的pptm檔內的字型大小資訊（第二張slide之設定）


### 程式碼效能測試：
整套 Adobe 仿宋 Std R 字型約20,957字圖在我母校華岡學習雲的公用雲端電腦跑完也要2、3天左右。故若電腦沒故障，須有耐心等待。程式碼並無出現錯誤，記憶體也不會不夠用。茲略錄其系統配置下：
* Windows 8.1 企業版
* 處理器：Intel(R)Xeon(R) CPU E5-2680 v3 @ 2.50GHz 
* 記憶體 8.00GB
* 系統類型：64位元作業系統，x64型處理器
* 電腦名稱：Pub8_1325
* 完整電腦名稱：pub8_1325.pccuad.pccu.edu.tw
* 網域：pccuad.pccu.edu.tw
* Office 2013


至於教職員個人雲則不到一天就跑完了「華康行書體」13,727字，規格如下：
* Windows 10 企業版 2016 長期維護
* 處理器：Intel(R) Xeon(R) CPU E5-2680 v4 @ 2.40GHz(2個處理器）
* 記憶體：8.00GB
* 系統類型：64位元作業系統，x64型處理器
* 電腦名稱：Personal-0331
* 完整電腦名稱：Personal-0331.pccuad.pccu.edu.tw
* 網域：pccuad.pccu.edu.tw
* Office 2016
* 在末學編輯這段說明時，不到半個鐘頭，或不出十幾分鐘吧，13,727的「華康行書體」字圖原樣 slide 已匯出成 png 完成 20210410日1059時。茲附上其 pptm檔，以省後人/未來菩薩續測時間。

【經驗：約保持在5000字左右在以上計算機性能上會有極佳的效能；若不似公用雲那般（那台的PowerPoint幾經測試其他功能（如插入古文字圖後清除）跑起來似乎特別慢），則1萬字或許是不錯的選擇。感恩感恩　南無阿彌陀佛。佛弟子孫守真任真甫合十敬啟。】


