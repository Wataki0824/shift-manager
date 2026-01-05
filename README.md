# 勤務表自動生成ツール

月間の勤務表を自動生成するWebアプリケーションです．

## 機能

- 📅 年月を選択してカレンダー表示
- 👥 スタッフ名を自由に設定（3名）
- 🖱️ クリックで休み（公/年）を簡単入力
- 🚀 ワンクリックで勤務表を自動生成
- 📥 Excelファイルでダウンロード

## ルール

### 平日（月〜金）
- 出勤者2人以上: 1人が9時，他は10時，10時の1人にYマーク
- 出勤者1人: 9時出勤かつYマーク

### 土日
- 出勤者はY9（9時出勤+Yマーク）

### Yマーク分配
- 月間で各スタッフが均等（±1以内）になるよう自動調整
- 休み明けの人にはできる限りYを避ける

---

## セットアップ

### 必要環境
- Python 3.8以上
- pip（Pythonパッケージマネージャー）

### インストール

```bash
# 1. プロジェクトフォルダに移動
cd /path/to/tnc

# 2. 必要なパッケージをインストール
pip install -r requirements.txt
```

### 起動

```bash
# Webアプリを起動
python app.py
```

起動後，ブラウザで以下のURLを開きます：
```
http://127.0.0.1:5001
```

---

## 使い方

### Step 1: 年月を選択
プルダウンから対象の年と月を選択します．

### Step 2: スタッフ名を入力
3名のスタッフ名を入力します（デフォルト: A, B, C）．

### Step 3: 休みを選択
カレンダーの各セルをクリックして休みを入力します．
- **1回クリック**: 公（公休）
- **2回クリック**: 年（年休）
- **3回クリック**: 空欄（出勤）に戻る

土日は背景が赤くハイライトされます．

### Step 4: 生成
「🚀 勤務表を生成」ボタンをクリックします．

### Step 5: ダウンロード
生成完了後，「📥 Excelをダウンロード」ボタンでExcelファイルを取得します．

---

## 出力形式

Excelファイルの形式：

| 日付 | 曜日 | A | | B | | C | |
|-----|-----|---|---|---|---|---|---|
| 1 | 月 | Y | 10 | | 9 | | 10 |
| 2 | 火 | | 9 | Y | 10 | | 10 |
| ... | ... | ... | ... | ... | ... | ... | ... |

- 各スタッフは2列（左: Yマーク，右: 時間/公/年）
- ヘッダー行のスタッフ名は結合されます

---

## コマンドライン版

Webアプリを使わずに，コマンドラインからも実行できます．

```bash
# CSV入力 → Excel出力
python generate_shift.py sample_input.csv output.xlsx

# CSV入力 → CSV出力
python generate_shift.py sample_input.csv output.csv
```

### 入力CSVの形式

```csv
日付,曜日,スタッフA,スタッフB,スタッフC
1,月,,,
2,火,公,,
3,水,,公,
```

- 休みの日に「公」または「年」を入力
- 出勤日は空欄

---

## ファイル構成

```
tnc/
├── app.py                 # Webアプリ（Flask）
├── generate_shift.py      # コマンドライン版
├── requirements.txt       # 依存パッケージ
├── sample_input.csv       # サンプル入力
├── templates/
│   └── index.html         # WebアプリのHTML
└── static/
    └── style.css          # WebアプリのCSS
```

---

## トラブルシューティング

### ポート5001が使用中
```bash
# 環境変数でポートを指定
PORT=5002 python app.py
```

### Excelファイルが開けない
openpyxlがインストールされているか確認：
```bash
pip install openpyxl
```

---

## クラウドデプロイ（Render.com）

Pythonを知らない人でも使えるよう，Webにデプロイできます．

### Step 1: GitHubにアップロード

```bash
cd /path/to/tnc
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/shift-schedule.git
git push -u origin main
```

### Step 2: Render.comでデプロイ

1. [Render.com](https://render.com) にサインアップ
2. 「New +」→「Web Service」
3. GitHubリポジトリを選択
4. 以下を設定：
   - **Name**: shift-schedule
   - **Runtime**: Python
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app`
5. 「Create Web Service」をクリック

### Step 3: 完了

デプロイ後，以下のようなURLでアクセスできます：
```
https://shift-schedule.onrender.com
```

ユーザーはこのURLを開くだけで勤務表を作成できます！

---

## Author
Taiki Watanabe

## License
MIT
