echo off

echo 外字フォントの確認

dir %systemroot%\fonts\eudc.*

ECHO "ボタンを押して外字管理画面が表示されたら成功です！"

pause

%SystemRoot%\system32\eudcedit.exe

