# 台灣彩券策略儀表板

響應式的大樂透與威力彩策略分析網站，顯示最新開獎資料、策略推薦、rolling backtest、隨機基準比較、包牌組合數與號碼排序訊號。

**公開網站：** https://samyiqrs.github.io/taiwan-lottery-dashboard/

## 網站內容

- 大樂透與威力彩切換
- 最新期號與開獎號碼
- 多種觀察策略
- Rolling backtest 策略比較
- 排序訊號與包牌成本提醒
- 手機、平板與桌機響應式版面
- 每晚 22:30（Asia/Taipei）由 GitHub Actions 嘗試更新官方資料
- 官方 API 失敗時自動重試並驗證期數、日期及號碼範圍；驗證失敗不會覆蓋既有網站
- 開獎資料沒有變化時不產生無效更新 commit

## 本機更新

本專案只使用 Python 標準函式庫：

```bash
python3 scripts/update_lottery.py
```

指令會更新：

- `index.html`（GitHub Pages 首頁）
- `台灣彩券號碼推薦器.html`
- `台灣彩券分析報告.xlsx`
- `latest_official_draws.json`
- `演算法審查與UI優化建議.md`

## 免責聲明

彩券開獎近似獨立隨機事件。本站提供的推薦與分數是統計觀察及策略排序訊號，不是中獎機率，也不保證獲利。請量力而為。

資料來源為台灣彩券官方公開資料；本專案不隸屬於台灣彩券。

## 更新失敗排查

到 repository 的 **Actions → Update lottery data** 查看紀錄，或手動執行 `workflow_dispatch`。若官方 API 暫時不可用，workflow 會失敗並保留前一次可用的 GitHub Pages 內容。
