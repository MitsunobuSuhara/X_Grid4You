# 【完全版】プロジェクトフォルダを安全に移動させる手順

このマニュアルは、`X-Grid-Package`プロジェクトフォルダの場所を、デスクトップからドキュメントフォルダなど、別の場所へ安全に移動させるための手順をまとめたものです。

> **なぜ、ただフォルダを移動するだけではダメなのか？**
>
> プロジェクトフォルダの中には、ファイルの場所を「絶対パス」（例: `C:\Users\mitsu\Desktop\...`）で記憶している設定ファイルがいくつか存在します。
> -   **Pythonの仮想環境 (`venv`フォルダ):** ライブラリの場所などを古い住所で記憶しています。
> -   **Inno Setupの設計図 (`.iss`ファイル):** インストーラーに含めるファイルの場所を古い住所で記憶しています。
>
> そのため、フォルダを移動した後に、これらの設定を新しい住所に正しく「住所変更」してあげる必要があります。

---

### ステップ1：ライブラリの一覧を保存する（引越し前の準備）

まず、引越し前のプロジェクトフォルダで、インストール済みのライブラリ一覧をファイルに保存します。

1.  ターミナル（PowerShellなど）を開き、**現在のプロジェクトフォルダ**に移動します。
    *（引越し前のデスクトップにあるプロジェクトフォルダで実行します）*
    ```powershell
    cd C:\Users\mitsu\Desktop\X-Grid-Package
    ```

2.  仮想環境を有効化します。
    ```powershell
    .\venv\Scripts\activate
    ```

3.  現在インストールされているライブラリの一覧を、`requirements.txt` というファイルに書き出します。
    ```powershell
    pip freeze > requirements.txt
    ```

4.  `X-Grid-Package` フォルダの中に `requirements.txt` というファイルが作成されたことを確認します。

---

### ステップ2：フォルダを移動する

エクスプローラーで、`X-Grid-Package` フォルダを、デスクトップから**ドキュメントフォルダに丸ごと移動**させます。

---

### ステップ3：仮想環境を再構築する（新しい住所で）

1.  移動先の `X-Grid-Package` フォルダの中にある、**古い `venv` フォルダを、手動で削除**してください。

2.  ターミナルを開き、**新しい場所**のプロジェクトフォルダに移動します。
    ```powershell
    cd C:\Users\mitsu\Documents\X-Grid-Package
    ```

3.  新しい場所で、仮想環境を**作り直し**ます。
    ```powershell
    python -m venv venv
    ```

4.  新しい仮想環境を**有効化**します。
    ```powershell
    .\venv\Scripts\activate
    ```

5.  ステップ1で作成した `requirements.txt` を使って、必要なライブラリを**一括で再インストール**します。
    ```powershell
    pip install -r requirements.txt
    ```
    *これで、仮想環境の引越しは完了です。*

---

### ステップ4：Inno Setupの設計図を修正する

1.  ドキュメントフォルダに移動させた `X-Grid-Package` フォルダの中から、`インストーラー.iss` （または `X_Grid.iss` など）のファイルを探し、テキストエディタ（Cursorなど）で開きます。

2.  ファイルの中に、`C:\Users\mitsu\Desktop\X-Grid-Package` という古い住所が書かれている行がいくつかあるはずです。
    -   `OutputDir=...`
    -   `SetupIconFile=...`
    -   `Source: ...`

3.  これらの行にある `Desktop` の部分を、すべて **`Documents`** に、手動で書き換えてください。

    **修正前:**
    `OutputDir=C:\Users\mitsu\Desktop\X-Grid-Package`
    **修正後:**
    `OutputDir=C:\Users\mitsu\Documents\X-Grid-Package`

4.  ファイルを上書き保存します。

---

これで、すべての「住所変更」手続きは完了です！
あなたのプロジェクトは、ドキュメントフォルダに完全に引越し、以前と全く同じようにビルドや開発を続けることができます。