@echo off
REM venvディレクトリが存在しなければ作成
if not exist venv (
    python -m venv venv
)

REM 仮想環境を有効化してpipアップグレード、ライブラリインストール
call venv\Scripts\activate
pip install --upgrade pip
pip install pynput

echo Library installation completed
pause