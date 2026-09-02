# OfflineDatabase Windows Portable v2

這個版本已整合「成品名 Brand Name」右側的生產記錄入口，以及正式版原有功能。

## 最重要的使用方式

使用者端 **不需要安裝 Python、PyInstaller 或 openpyxl**。

Windows EXE 由 GitHub Actions 的 Windows 建置環境自動產生。

## GitHub 一鍵建置

1. 建立一個 GitHub Repository。
2. 將本資料夾內的所有檔案上傳到 Repository 根目錄。
3. 確認 `.github/workflows/build-windows.yml` 也一併上傳。
4. 在 GitHub 開啟該 Repository。
5. 點選 **Actions**。
6. 左側選 **Build Windows Portable EXE**。
7. 點選 **Run workflow**。
8. 等待工作完成。
9. 在完成的 workflow 頁面底部的 **Artifacts** 下載：
   - `OfflineDatabase-Windows-Portable-ZIP`

解壓縮後即可得到 Windows Portable 版本。

## Portable 結構

```text
OfflineDatabase\
├─ OfflineDatabase.exe
├─ initial_data.xlsx
├─ README_使用說明.md
├─ data\
├─ backup\
└─ （PyInstaller 所需的程式檔案）
```

請整個 `OfflineDatabase` 資料夾一起攜帶，不要只複製 EXE。

## 自訂圖示

如果你有原本的 `OfflineDatabase.ico`，請把它放在 Repository 根目錄，與 `app.py` 同一層。

GitHub Actions 建置時會自動：

- 使用該 `.ico` 作為 EXE 圖示。
- 將 `.ico` 一併複製到 Portable 資料夾。

如果沒有 `.ico`，仍然可以正常建置，只是會使用 PyInstaller 預設圖示。

## 生產記錄

主資料表的「成品名 Brand Name」右側有生產記錄入口。

點擊後會開啟獨立的「生產記錄」視窗，可新增、編輯、刪除：

- 生產日期
- 批號
- 生產數量
- 備註

資料會儲存在正式 SQLite 資料庫中。

## build_windows.bat

`build_windows.bat` 是給需要在 Windows 建置機自行打包時使用的。一般使用者不需要執行它。
