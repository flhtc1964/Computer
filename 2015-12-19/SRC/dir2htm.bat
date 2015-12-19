echo off

rem =================================
rem =                               =
rem =  findstr　で正規表現を体験    =
rem =  onigsed　で正規表現を体験    =
rem =  2015-12-19                   =
rem =================================


echo Windowsで使用するBATファイルです
echo 【重要】処理出来る行数に制限がありますので注意！
echo 2015-10-27

echo -


echo 作者名　SKojima@kuhen.jp


echo -


echo このBATファイルをマウス右クリック　→　「管理者として実行」を選択！

echo -

echo 起動ファイル名【 %0 】

echo -

echo 現在の居場所【 %~dp0 】

echo -


set /P fpath="検索フォルダ入力（未入力なら現在のフォルダを使用）"


rem もしも【fpath】が==空白（""）ならば、現在の居場所をfpathにセット

IF "%fpath%" == "" (set fpath=%~dp0)


echo 検索対象フォルダ【 %fpath% 】


echo -


set /P fkey="検索ファイル名を入力（例：txt）"


echo -


echo ファイル一覧作成中（%~dp0foo.txt）


echo -


rem dir /B →　ファイル名のみ表示
rem dir /I →　指定ディレクトリとその下のサブディレクトリを表示

dir /B /S %fpath% > %~dp0foo.txt


echo -


echo %~dp0foo.txtから%fkey%を検索し（%~dp0foo2.txt）を作成中


echo -


rem findstr /R →　検索文字を正規表現として扱う
rem findstr /I →　大文字・小文字区別しない
rem findstr /L →　検索文字列をリテラルとして使用します。

findstr /I /L "%fkey%" %~dp0foo.txt > %~dp0foo2.txt


echo -


echo 次行の最後に【：検索ヒット数】が表示されます


echo -


rem ====================================================
rem 【引用元】指定ファイルの行数を取得
rem http://d.hatena.ne.jp/necoyama3/20090716/1247752451
rem ====================================================


FOR /F "DELIMS=" %%A IN ('FIND /C /V "" %~dp0foo2.txt') DO SET LINECOUNT=%%A
ECHO %LINECOUNT% 


echo -

rem %LINECOUNT%の後ろから３文字取得　→　"%LINECOUNT:~-3%"

IF "%LINECOUNT:~-3%" == ": 0" goto zero

IF not "%LINECOUNT:~-3%" == ": 0" goto htm

:zero

echo 該当件数が無いので処理を終了します

pause

goto end


:htm

echo -

echo 今から上記検索結果参照用HTMLを自動作成します


pause


rem =====================================
rem =onigsed.exeの【取得元】            =
rem = http://www.kt.rim.or.jp/~kbk/sed/ =
rem =                                   =
rem = onigsed.exe を使用                =
rem =                                   =
rem =====================================


rem 新規作成 >

echo ^<HTML^> > %~dp0foo3.txt


rem 既存ファイルに追加 >>

echo ^<HEAD^> >> %~dp0foo3.txt


echo 【 %fkey% 】一覧 ^<P^> >> %~dp0foo3.txt


echo ^</HEAD^> >> %~dp0foo3.txt


echo ^<BODY^> >> %~dp0foo3.txt


%~dp0onigsed.exe -f %~dp0onig1.txt  %~dp0foo2.txt >> %~dp0foo3.txt


echo ^</BODY^> >> %~dp0foo3.txt
echo ^</HTML^> >> %~dp0foo3.txt

%~dp0onigsed.exe -f %~dp0onig2.txt %~dp0foo3.txt > %~dp0file_list.htm


rem 作業ファイルを削除

del %~dp0foo.txt
del %~dp0foo2.txt
del %~dp0foo3.txt

rem 検索結果をHTMLファイル形式で起動

%~dp0file_list.htm



:end


