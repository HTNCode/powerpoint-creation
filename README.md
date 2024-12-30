# powerpoint-creation

    - gpt-researcher と marp ライブラリを使って PowerPoint 資料を作成するサンプル。
    - とりあえず遊びで生成 AI が吐き出したもので動作検証したものであるため、いまいちな結果に。
    - きちんと各機能を読み解いて作りこめば、ちまたでよく使われているような PowerPoint 形式での資料作成自動化が作れるかも

# 使い方

    1. .envファイルに以下設定
    ```
    OPENAI_API_KEY=your-api-key
    TAVILY_API_KEY=your-api-key
    ```
    2. ライブラリのインストール
    3. python main.pyで実行すると、何を調べるか聞かれるのでターミナル上で答える
    4. outputディレクトリにパワーポイント資料と生成した画像が保存される
