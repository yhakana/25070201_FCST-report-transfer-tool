# FCST report transfer tool

- 需求編號: 25070201
- 需求者: Darlene Weng
- 原作者: Harry Liu -> Shawn Wang (Harry 沒留 source code, 故重寫此需求)

---

### 使用步驟

1. 建立 `resource` 資料夾 (如果沒有的話)
2. 將目標檔案 `.xlsm` 放入 `resource` 資料夾中 (一次可執行多個檔案)
3. 執行 `main.exe`
4. 會在 `export` 資料夾中得到輸出檔 `{原檔名}_export.xlsm`

- 其他: `log` 資料夾可以查看執行狀況與錯誤訊息

- p.s. 此程式會偵測是否已有 `export`, `resource`, `log` 資料夾, 若沒有會自動建立
- p.s. 此程式會檢查 `export` 資料夾是否有先前執行的結果, 如果有會進行移除

---

### 開發者步驟

```
# 請在命令列中輸入以下指令, 建立虛擬環境
$ python -m venv venv

# 啟用 venv (如果你的作業系統為 Windows)
$ venv/Scripts/activate
# 啟用 venv (如果你的作業系統為 Linux / macOS)
$ source venv/bin/activate

# 安裝依賴套件
$ pip install -r requirements.txt

# (如果你有安裝 / 解除依賴套件)
$ pip freeze > requirements.txt

# 啟用程式
$ python main.py

```

---

### 主要程式

- main.py

---

### 需求內容

- 將原始 .xlsm 檔案中的 `CP Summary` 與 `FT Summary` 兩個分頁的內容整理至 `Data presentation` 分頁中
- 會擷取原始檔案的部分內容, 輸出至 `Data presentation` 的表格中
- 僅有整理內容, 沒有進行特殊計算或分析
