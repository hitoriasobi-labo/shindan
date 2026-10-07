# PC ルーティン「VICTAS マーケ月次報告書」の指示文

Claude デスクトップアプリ → ルーティン → 新しいルーティン で、スケジュールを「毎月1日 8:52」にして、下の枠の中をそのまま貼り付ける。

```
VICTAS マーケティング本部の月次報告書（前月分）を作り、板橋靖和（yasukazu.itabashi@victas.com）に Outlook でメール送信してください。板橋が毎月1日に全社へ出す報告書の下書きです。2026年9月度までは坂田悦代さんが作成していました。

対象月 = 実行日の前月（例: 11/1 実行なら 2026年10月度）。

■ 形式
坂田さんの「202609_マーケティング本部月次報告書.xlsx」と同じ2シート構成（「WEB・SNS数値」「トピック」）。
過去の数値と作成スクリプトは GitHub の hitoriasobi-labo/shindan リポジトリ、ブランチ claude/victas-sns-metrics-rsuhof の reports/ フォルダにある（README.md 参照）。
- reports/data/metrics.csv: 月末の数値（X / Instagram / YouTube / LINE のフォロワー数、アクセス数、WEBサイト UU数）
- reports/topics/YYYY-MM.tsv: 実施内容（カテゴリー1、カテゴリー2、詳細 のタブ区切り）
- reports/build_report.py: 上の2つから Excel を作る（python3 build_report.py YYYY-MM）
リポジトリが PC になければ clone、あれば pull する。Python や openpyxl が使えなければ、前月の報告書と同じレイアウトの Excel を別の方法で作る。

■ 手順
1. SNS の月末フォロワー数を調べる。X (@victas_inc)、Instagram (@victas.inc)、YouTube (@VICTAS-Inc)、LINE公式アカウント。Facebook は運用終了のため対象外。
   ブラウザで各アカウントのページや管理画面を開いて数字を読む。読めなかったものは空欄にして「要確認」とする。
2. アクセス数・WEBサイト（UU数）は、取得元が決まるまで空欄にして「要確認」とする。
3. 実施内容は、対象月のマーケティング関連の資料（SharePoint の marketing サイト「0.Report・Schedule」の週次MTG資料など）、メール、Teams から拾って、カテゴリー1/カテゴリー2/詳細 に整理する。
4. metrics.csv に対象月の行を追加し（source 列に取得元を書く）、topics/YYYY-MM.tsv を作り、Excel を作る。
5. 変更した CSV・TSV・Excel を同じブランチにコミットして push する。push できなければ、その理由をメールに書く。
6. 板橋に Outlook でメールを送る。件名「【下書き】YYYY年M月度 マーケティング本部 月次報告書」、Excel を添付。本文には次を書く:
   - SNS の数字と前月比
   - 空欄・要確認の項目の一覧
   - 実施内容の要約
   全社宛て（all@victas.com）には送らない。板橋が確認してから自分で送る。
```
