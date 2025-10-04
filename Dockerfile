FROM python:3.12-slim

WORKDIR /app

# 必要なシステムパッケージをインストール（Camelot用）
RUN apt-get update && apt-get install -y \
    ghostscript \
    python3-tk \
    && rm -rf /var/lib/apt/lists/*

# Poetry のインストール
RUN pip install poetry

# 依存関係ファイルをコピー
COPY pyproject.toml poetry.lock ./

# 依存関係をインストール
RUN poetry config virtualenvs.create false \
    && poetry install --no-interaction --no-ansi --no-root

# アプリケーションコードをコピー
COPY . .

# スクリプトを実行
CMD ["python", "main.py"]