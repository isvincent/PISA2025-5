# PISA 2025 臺灣學習羅盤

這是一個可直接部署到 GitHub Pages 的純靜態互動網站，不需要安裝套件或執行建置指令。

## GitHub Pages 結構

```text
.
├── index.html
├── .nojekyll
├── assets/
│   └── images/
│       └── oecd-country-note-cover.jpg
└── downloads/
    └── PISA_2025_臺灣學習羅盤.docx
```

## 上架方式

1. 將根目錄的 `index.html`、`.nojekyll`、`assets`、`downloads` 推送到 GitHub repository。
2. 在 repository 的 **Settings → Pages** 開啟 GitHub Pages。
3. Source 選擇 **Deploy from a branch**，Branch 選擇 `main`，資料夾選擇 `/(root)`。
4. 儲存後等待 GitHub 顯示正式網址。

網站內所有本機資源均使用相對路徑，可部署於專案型網址，例如 `https://username.github.io/repository-name/`。
