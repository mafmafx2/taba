# taba

PTSランキングを自動表示し、該当銘柄の直近適時開示リンクを一覧化するシングルページアプリです。

## 使い方

```bash
cd /workspace/taba
python -m http.server 8000 --directory public
```

ブラウザで `http://localhost:8000` を開いてください。

## データ更新

- `public/data/pts-ranking.json` にPTSランキング情報を保存します。
- `public/data/disclosures.json` に適時開示情報を保存します。
- 実運用時はAPI連携に置き換えてください。
