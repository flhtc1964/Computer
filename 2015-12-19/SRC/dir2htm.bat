echo off

rem =================================
rem =                               =
rem =  findstr�@�Ő��K�\����̌�    =
rem =  onigsed�@�Ő��K�\����̌�    =
rem =  2015-12-19                   =
rem =================================


echo Windows�Ŏg�p����BAT�t�@�C���ł�
echo �y�d�v�z�����o����s���ɐ���������܂��̂Œ��ӁI
echo 2015-10-27

echo -


echo ��Җ��@SKojima@kuhen.jp


echo -


echo ����BAT�t�@�C�����}�E�X�E�N���b�N�@���@�u�Ǘ��҂Ƃ��Ď��s�v��I���I

echo -

echo �N���t�@�C�����y %0 �z

echo -

echo ���݂̋��ꏊ�y %~dp0 �z

echo -


set /P fpath="�����t�H���_���́i�����͂Ȃ猻�݂̃t�H���_���g�p�j"


rem �������yfpath�z��==�󔒁i""�j�Ȃ�΁A���݂̋��ꏊ��fpath�ɃZ�b�g

IF "%fpath%" == "" (set fpath=%~dp0)


echo �����Ώۃt�H���_�y %fpath% �z


echo -


set /P fkey="�����t�@�C��������́i��Ftxt�j"


echo -


echo �t�@�C���ꗗ�쐬���i%~dp0foo.txt�j


echo -


rem dir /B ���@�t�@�C�����̂ݕ\��
rem dir /I ���@�w��f�B���N�g���Ƃ��̉��̃T�u�f�B���N�g����\��

dir /B /S %fpath% > %~dp0foo.txt


echo -


echo %~dp0foo.txt����%fkey%���������i%~dp0foo2.txt�j���쐬��


echo -


rem findstr /R ���@���������𐳋K�\���Ƃ��Ĉ���
rem findstr /I ���@�啶���E��������ʂ��Ȃ�
rem findstr /L ���@��������������e�����Ƃ��Ďg�p���܂��B

findstr /I /L "%fkey%" %~dp0foo.txt > %~dp0foo2.txt


echo -


echo ���s�̍Ō�Ɂy�F�����q�b�g���z���\������܂�


echo -


rem ====================================================
rem �y���p���z�w��t�@�C���̍s�����擾
rem http://d.hatena.ne.jp/necoyama3/20090716/1247752451
rem ====================================================


FOR /F "DELIMS=" %%A IN ('FIND /C /V "" %~dp0foo2.txt') DO SET LINECOUNT=%%A
ECHO %LINECOUNT% 


echo -

rem %LINECOUNT%�̌�납��R�����擾�@���@"%LINECOUNT:~-3%"

IF "%LINECOUNT:~-3%" == ": 0" goto zero

IF not "%LINECOUNT:~-3%" == ": 0" goto htm

:zero

echo �Y�������������̂ŏ������I�����܂�

pause

goto end


:htm

echo -

echo �������L�������ʎQ�ƗpHTML�������쐬���܂�


pause


rem =====================================
rem =onigsed.exe�́y�擾���z            =
rem = http://www.kt.rim.or.jp/~kbk/sed/ =
rem =                                   =
rem = onigsed.exe ���g�p                =
rem =                                   =
rem =====================================


rem �V�K�쐬 >

echo ^<HTML^> > %~dp0foo3.txt


rem �����t�@�C���ɒǉ� >>

echo ^<HEAD^> >> %~dp0foo3.txt


echo �y %fkey% �z�ꗗ ^<P^> >> %~dp0foo3.txt


echo ^</HEAD^> >> %~dp0foo3.txt


echo ^<BODY^> >> %~dp0foo3.txt


%~dp0onigsed.exe -f %~dp0onig1.txt  %~dp0foo2.txt >> %~dp0foo3.txt


echo ^</BODY^> >> %~dp0foo3.txt
echo ^</HTML^> >> %~dp0foo3.txt

%~dp0onigsed.exe -f %~dp0onig2.txt %~dp0foo3.txt > %~dp0file_list.htm


rem ��ƃt�@�C�����폜

del %~dp0foo.txt
del %~dp0foo2.txt
del %~dp0foo3.txt

rem �������ʂ�HTML�t�@�C���`���ŋN��

%~dp0file_list.htm



:end


