# X-Grid & X-Grid Styler 開発・リリース作業マニュアル【超完全版・改】

このマニュアルは、`X_Grid`アプリ本体と`X-Grid Styler` QGISプラグインの開発から、新しいバージョンの配布ファイルをGitHubにリリースするまでの一連の作業手順を、詳細にまとめたものです。よくあるトラブルシューティングも含まれています。

---

## Part 0: 初期セットアップ（PCごとに最初の一回だけ）

新しい開発環境を整える際に、最初に行うべき設定です。

### 1. Git LFS (Large File Storage) のインストール

-   **目的:** 25MBを超える大きなファイル（インストーラーなど）をGitHubにアップロードできるようにするためのGit拡張機能です。
-   **手順:**
    1.  [Git LFSの公式サイト](https://git-lfs.com/)からインストーラーをダウンロードし、インストールします。
    2.  ターミナル（PowerShellなど）を開き、以下のコマンドを一度だけ実行します。
        ```bash
        git lfs install
        ```

### 2. `.gitignore` ファイルの作成

-   **目的:** 不要なファイル（ビルド時の中間ファイル、仮想環境フォルダなど）をGitの管理対象から除外し、リポジトリを常にクリーンな状態に保つための設定ファイルです。
-   **手順:**
    1.  プロジェクトフォルダの**一番上の階層**に、`.gitignore` という名前のテキストファイルを作成します。（`.txt`などの拡張子が付かないように注意）
    2.  以下の内容をコピー＆ペーストして保存します。

        ```
        # Python
        __pycache__/
        *.pyc
        venv/
        
        # PyInstaller
        build/
        dist/
        *.spec
        
        # Inno Setup
        Output/
        *.iss
        
        # OS generated files
        .DS_Store
        Thumbs.db
        ```

### 3. Inno Setup のインストール

-   **目的:** PyInstallerで生成したファイル群を、単一で高圧縮な `.exe` 形式のインストーラーにまとめるためのツールです。
-   **手順:**
    1.  [Inno Setupの公式サイト](https://jrsoftware.org/isinfo.php)から安定版をダウンロードし、インストールします。

---

## Part 1: 開発とリリースのサイクル（更新ごとに行う作業）

ここからの手順は、アプリやプラグインを更新して、新しいバージョンをリリースするたびに繰り返し行います。

### ステップ1: コードの編集とテスト

1.  **仮想環境の有効化**
    -   ターミナルを開き、プロジェクトフォルダに移動します。
        ```powershell
        cd path\to\your\X-Grid-Package
        ```
    -   以下のコマンドで、このプロジェクト専用のPython環境を有効化します。
        ```powershell
        .\venv\Scripts\activate
        ```
        *(プロンプトの先頭に `(venv)` と表示されれば成功です)*

2.  **ソースコードの編集**
    -   `X_Grid.py` や `x_grid_styler.py` をエディタで開き、必要な修正や機能追加を行います。
    -   `python X_Grid.py` を実行するなどして、ローカルで十分に動作をテストします。

### ステップ2: 配布ファイルのビルド

1.  **QGISプラグインのzipファイル作成**
    -   `x_grid_styler` フォルダ（`__init__.py`, `x_grid_styler.py`などが入っている）を丸ごとzipで圧縮します。
    -   ファイル名を `x_grid_styler.zip` として、プロジェクトフォルダのトップに保存します。

2.  **本体アプリのビルド (PyInstaller)**
    -   **古いビルドファイルを削除:** `build` フォルダと `dist` フォルダがあれば、手動で削除します。
    -   ターミナルで以下のコマンドを実行します。
        ```powershell
        pyinstaller X_Grid.spec
        ```
    -   `dist` フォルダ内に `X_Grid` という名前のフォルダが生成されます。

### ステップ3: インストーラーの作成 (Inno Setup)

1.  **Inno Setupの起動**
    -   Inno Setupを起動します。
    -   **💡ヒント:** 前回作成したインストーラーの設計図 (`.iss`ファイル) があれば、それを開いて**バージョン番号だけを修正**するのが一番簡単です。

2.  **ウィザードの実行（新規作成の場合）**
    -   `Application version`: 新しいバージョン番号 (例: `1.1.0`) に更新します。
    -   `Application main executable file`: `dist\X_Grid\X_Grid.exe` を指定します。
    -   `Other application files`: **[Add folder...]** ボタンで `dist\X_Grid` フォルダ全体を指定します。
    -   `Setup Install Mode`: **"Non administrative install mode"** を選択します。
    -   `Compiler output base file name`: 新しいバージョン番号を含めた名前にします (例: `X_Grid_v1.1.0_setup`)。

3.  **インストーラーのコンパイル**
    -   メニューの `Build` > `Compile` を実行します。
    -   単一の `.exe` 形式のインストーラーが生成されます。
    -   **重要:** このインストーラーを、プロジェクトフォルダ内の `installer` フォルダに配置します。

### ステップ4: GitHubへのアップロード (Push)

1.  **変更内容の確認 (`git status`)**
    -   ターミナルで `git status` を実行し、変更・追加されたファイルを確認します。

2.  **すべての変更を追加 (`git add`)**
    -   `.gitignore` があるので、安全にすべての変更をステージングできます。
        ```powershell
        git add .
        ```

3.  **変更を記録 (`git commit`)**
    -   変更内容が分かるように、メモを付けてPC内に記録します。
        ```powershell
        git commit -m "feat: バージョン1.1.0の更新 (ここに具体的な変更内容を書く)"
        ```

4.  **GitHubに送信 (`git push`)**
    -   PC内の記録を、インターネット上のGitHubにアップロードします。
        ```powershell
        git push origin main
        ```

    -   **🚨【重要】`push`が失敗 (rejected) した場合の対処法**
        -   `error: failed to push...` というメッセージが表示されたら、それはあなたのPCが知らない変更が先にGitHubに存在するためです。慌てずに以下の手順に進んでください。

    -   **手順A: GitHubの最新の変更を取り込む (`git pull`)**
        ```powershell
        git pull origin main
        ```
        -   **成功した場合:** `Successfully merged...` のようなメッセージが出たら、競合はありませんでした。もう一度 `git push origin main` を実行して、ステップ5に進んでください。
        -   **コンフリクトが発生した場合:** `CONFLICT ...` `Automatic merge failed;` というメッセージが出たら、**手順B**に進んでください。

    -   **手順B: コンフリクト (競合) の解決**
        -   Gitが自動で統合できず、あなたの判断が必要な状態です。
        -   `git status` を実行すると、どのファイルが競合しているか (`Unmerged paths`) がわかります。
        -   競合したファイルごとに、**自分の変更を残すか、相手(GitHub)の変更を残すか**を決定し、以下のコマンドを実行します。

            -   **自分の変更を優先する場合 (例: 自分が削除したなら、削除を正とする)**
                ```powershell
                git checkout --ours "ファイル名"
                ```
            -   **相手(GitHub)の変更を優先する場合 (例: GitHub上の更新版を採用する)**
                ```powershell
                git checkout --theirs "ファイル名"
                ```
            *(例: `取扱説明書.pdf`が競合し、自分の「削除」を優先するなら `git rm 取扱説明書.pdf` を実行します)*

        -   競合ファイルをすべて解決したら、それをGitに伝えます。
            ```powershell
            # 変更を適用したファイルをaddする
            git add . 
            ```

        -   最後に、マージを完了させるためのコミットを実行します。
            ```powershell
            # このコマンドを実行するとエディタが開くので、何も変更せずに保存して閉じる
            git commit
            ```

        -   これでローカルの準備が整いました。満を持して、もう一度 `push` します。
            ```powershell
            git push origin main
            ```

### ステップ5: GitHubでのリリース作成

1.  ブラウザであなたのGitHubリポジトリを開き、右側の **[Releases]** > **[Draft a new release]** をクリックします。

2.  **リリース情報を入力**
    -   **Tag version**: 新しいバージョン番号を入力します (例: `v1.1.0`)。
    -   **Release title**: リリースのタイトルを入力します (例: `X_Grid v1.1.0`)。
    -   **Describe this release**: このバージョンでの変更内容や、ユーザーへのメッセージを記述します。

3.  **配布ファイルを添付**
    -   **"Attach binaries..."** のエリアに、以下の**2つのファイル**をドラッグ＆ドロップします。
        1.  **本体アプリのインストーラー:** `installer` フォルダの中の `X_Grid_v1.1.0_setup.exe`
        2.  **QGISプラグインのzip:** プロジェクトフォルダ直下の `x_grid_styler.zip`

4.  **公開**
    -   **(推奨)** `Set as a pre-release` にチェックを入れます。
    -   **[Publish release]** ボタンを押して、リリースを公開します。

---

**これで、すべての更新作業は完了です！** 🎉