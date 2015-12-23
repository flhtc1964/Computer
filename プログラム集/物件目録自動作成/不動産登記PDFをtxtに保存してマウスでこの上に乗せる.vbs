'=================================
'(C)skojima@kitahama.or.jp
'=================================

Option Explicit


'========================

Dim strVER

strVER = "Ver. 2015-02-26 16:01"

'========================



Dim args, arg
Dim objFileSys
Dim objFileSys2
Dim objFileSys3

Dim objTextStream
Dim objTextStream2 
Dim objTextStream3
Dim objWriteStream3

Dim strText

Dim strWRFN

Dim strWriteFile

Dim strShurui

Dim I,ichi,ichi2,ichi3

Dim strWK1

Dim strWK_ALL

Dim strWRSAVE




msgbox "【  " & strVER & " 】" & vbCrLf & "不動産登記情報PDFを開き、編集→すべてを選択→コピー→メモ帳に貼り付けて保存した物を" & vbCrLf & "このアイコン上に乗せてください！" & vbCrLf & "By Skojima@kitahama.or.jp"
msgbox "【重要】このプログラムは無保証です！"

Set args = WScript.Arguments
For Each arg In args
    Call MsgBox(arg,,"【今から物件目録を作成します】")

    
Dim objFile    ' 対象ファイル
Dim oldText    ' 置換前テキスト
Dim newText    ' 置換後テキスト
Dim objFSO     ' ファイルシステムオブジェクト
Dim objRep     ' 正規表現オブジェクト
Dim repText    ' 置換対象文字列

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(arg)

' テキストデータ読込
oldText = objFile.ReadAll

' 置換対象文字列
repText = ""

Set objRep = New RegExp

'一行毎に処理→True
objRep.Multiline = True

'大文字小文字区別しない→True
objRep.IgnoreCase = True

'最初にマッチした部分だけ→False
objRep.Global = True


repText = "+"


' 正規表現パターンを指定する
objRep.Pattern = "^.*┏━.*┓"
newText = objRep.replace(oldText, repText)

oldText = newText
objRep.Pattern = "^.*┗━.*┛"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*┠─.*┨"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*┠─.*┨"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*┃.*┨"
newText = objRep.replace(oldText, repText)

oldText = newText
objRep.Pattern = "　| "
repText = ""
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^[+]┃"
repText = "+" & vbCrLf &"┃"
newText = objRep.replace(oldText, repText)

objFile.Close

'========================================================

if right(arg,4) <> ".txt" then
msgbox "エラー：この拡張子は、半角の「.txt」ではないので処理を中止します！" & vbCrLf & "例：　○○.TXT　←大文字のTXTだと処理出来ません！"& vbCrLf & "例：　○○.txt　←半角にすること"
                        wscript.quit
end if


strWRFN = replace(arg,".txt","_作業用.txt")


' 書き込み
Set objFile = objFSO.CreateTextFile(strWRFN)
objFile.WriteLine (newText)
objFile.Close

Set objFSO = Nothing

'=============================================================
'これで、罫線は消えて、データ無い空白行は、「+」に置換された
'=============================================================




'------------------------------------
Set objFileSys2 = CreateObject("Scripting.FileSystemObject")
Set objTextStream2 = objFileSys2.OpenTextFile(strWRFN, 1)

ichi = 0

'******************************************************************
Do while objTextStream2.AtEndOfStream <> True
   strText = objTextStream2.ReadLine
   
   I = I +1

	if instr(strText,"表題部（一棟の建物の表示）") > 0 then
            strShurui = 1

    end if
    
    
    if instr(strText,"表題部（主である建物の表示）") > 0 then
            strShurui = 2


    end if
    
    
    if instr(strText,"表題部（土地の表示）") > 0 then
            strShurui = 3

    end if
   
Loop
'******************************************************************

objTextStream2.Close
Set objFileSys2 = Nothing

'=============================================================='

'　　　　　　　　　　読込作業ファイル名


strWK_ALL = "　　　　　　　　　　　物　件　目　録" & vbCrLf

Select case  strShurui

case 1

            msgbox "表題部（一棟の建物の表示）"

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12(strWRFN ,"表題部（一棟の建物の表示）",2) 


strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123(strWRFN ,"表題部（専有部分の建物の表示）",4)



strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_1234(strWRFN ,"表題部（敷地権の目的である土地の表示）",4)




strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_123(strWRFN ,"表題部（敷地権の表示）",4)

case 2

            msgbox "表題部（主である建物の表示）"

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123(strWRFN ,"表題部（主である建物の表示）",4)

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123_FUZOKU(strWRFN ,"表題部（附属建物の表示）",4)


case 3
            msgbox "表題部（土地の表示）"


strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123_B(strWRFN ,"表題部（土地の表示）",4)



case else
msgbox "このファイルには、該当処理項目が無いので、処理を中止します！"
                        wscript.quit

End Select

'------------------------------------


'■目録保存用
strWRSAVE = replace(arg,".txt","_変換後の物件目録.txt")

msgbox strWRSAVE

'msgbox strWK_ALL

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strWRSAVE)
objFile.WriteLine (strWK_ALL)
objFile.Close

Set objFSO = Nothing




strWK_ALL = ""

'msgbox strWRSAVE,"書き込み完了"
    
'----------------------
Next

WScript.echo "処理が完了しました！"



'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２列処理用
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_12(tmpFILE, tmpWORD,tmpTASU)

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi


Dim strWK1
Dim strWK2


strWK1 = ""
strWK2 = ""


strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0

'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
						'+が２つ連続したら処理を止める
						if  instr(strText,"┃①構造") > 0  then
							ichi = 1
						end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"│")
						
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						
						ichi = ichi + 1
						
						else
						
						strRT_WK = strRT_WK & strWK1 & "　" & strWK2 & vbCrLf

						
						strWK1 = ""
						strWK2 = ""

						
						ichi = ichi + 1
						
					end if
					
				end if
				
				
				
		end if
		
		

Loop
'******************************************************************

objTextStream4.Close


	if strRT_WK = tmpWORD & vbCrLf then
		strRT_WK = ""
	end if


		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)
		
		



    '結果返す
    DAT_PUT_12 = strRT_WK

End Function




'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２と１、２、３列処理用
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_12_123(tmpFILE, tmpWORD, tmpTASU )

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi,ichi2

Dim strWK1
Dim strWK2

Dim strWK3
Dim strWK4
Dim strWK5

strWK1 = ""
strWK2 = ""

strWK3 = ""
strWK4 = ""
strWK5 = ""

strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		

						'+が２つ連続したら処理を止める  「┃表題部（附属建物の表示）」は→別関数で処理
						if  instr(strText,"┃表題部（附属建物の表示）") or instr(strText,"┃権利部") > 0 or instr(strText,"┃所有者") > 0 or instr(strText,"┃表題部（敷地権の表示）") > 0 then
							ichi2 = 1
							ichi = 0
						end if
		
		
		if ichi > 0 then
		
		
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+"  then
					
						'msgbox I & "→1→" & strText
					
						strSPL_WK = split(strText,"│")
						
						
						'msgbox "2→" &strSPL_WK(0) & "|" & strSPL_WK(1)
						
						if strSPL_WK(0) <> "┃①種類" and strSPL_WK(1) <> "②構造" then
						
						
							'┃　があるから > 1
							if instr(strSPL_WK(0), "") = 0 and len(strSPL_WK(0)) > 1 then
								strWK1 = strSPL_WK(0)
							end if
						
						
							if instr(strSPL_WK(1), "") = 0 and len(strSPL_WK(1)) > 0 then
								strWK2 = strWK2 & strSPL_WK(1)
							end if
							
							'msgbox "3→" & strWK1 & strWK2
						
						end if
						
						
						ichi = ichi + 1
						
						else
						
						ichi = ichi + 1
						
						if strSPL_WK(0) = "┃①種類" and strSPL_WK(1) = "②構造" then
																		
							ichi2 = I + 1
							ichi = 1
						else
						
							'msgbox "4→" & strWK1 & strWK2
												
							strRT_WK = strRT_WK & strWK1 & "　" & strWK2 & vbCrLf
							
							strWK1 = ""
							strWK2 = ""
						
						end if
						
						
					end if
					
					
				end if
				
		end if
		
		
		if ( I > 0 ) and ( I = ichi2 ) then
		
				if left(strText,1) <> "+" and instr(strText,"│") > 0 then
				strSPL_WK = split(strText,"│")
				
				'msgbox("--" & strText)
				

				
						strWK3 = strWK3 & strSPL_WK(0)
						strWK4 = strWK4 & strSPL_WK(1)
						
						if strSPL_WK(2) <> "：" then
							strWK5 = strWK5 & strSPL_WK(2) & "　"
						end if
						

						
						ichi2 = ichi2 + 1
						
					else
					
									'空白は登録しない
									
						'msgbox ("|" & strWK5 & "|" & len(strWK5))
						
						if instr(strWK3, "") = 0 and len(strWK5) > 1 then
					
							strRT_WK = strRT_WK & "　種　　　　類　" & strWK3 & vbCrLf
							strRT_WK = strRT_WK & "　構　　　　造　" & strWK4 & vbCrLf
							strRT_WK = strRT_WK & "　床　 面　 積　" & replace(replace(strWK5,"：","㎡"),"分","分　") & vbCrLf 
					
						end if
						
						strWK3 = ""
						strWK4 = ""
						strWK5 = ""
						
						ichi2 = ichi2 + 1
						
					
				end if
		
		end if
		

Loop
'******************************************************************

objTextStream4.Close


if strRT_WK = tmpWORD & vbCrLf then
	strRT_WK = ""
end if


		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)



    '結果返す
    DAT_PUT_12_123 = strRT_WK

End Function




'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２と１、２、３列処理用　付属建物
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_12_123_FUZOKU(tmpFILE, tmpWORD, tmpTASU )

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi,ichi2

Dim strWK1
Dim strWK2
Dim strWK3
Dim strWK4


strWK1 = ""
strWK2 = ""
strWK3 = ""
strWK4 = ""


strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		

						'+が２つ連続したら処理を止める 
						if  instr(strText,"┃権利部") > 0 or instr(strText,"┃所有者") then
							ichi = 0
						end if
		
		
	if ichi > 0 then
		
		if ( I > 0 ) and ( I = ichi + tmpTASU ) then
		
				if left(strText,1) <> "+" and instr(strText,"│") > 0 then
				strSPL_WK = split(strText,"│")
				
				'msgbox("--" & strText)
				

				
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						strWK3 = strWK3 & strSPL_WK(2)
						
						if strSPL_WK(3) <> "：" then
							strWK4 = strWK4 & strSPL_WK(3) & "　"
						end if
						

						
						ichi = ichi + 1
						
					else
					
									'空白は登録しない
									
						'msgbox ("|" & strWK1 & "|" & len(strWK4))
						
						if instr(strWK1, "") = 0 and len(strWK4) > 1 then
							strRT_WK = strRT_WK & "　符　　　　号　" & strWK1 & vbCrLf
							strRT_WK = strRT_WK & "　種　　　　類　" & strWK2 & vbCrLf
							strRT_WK = strRT_WK & "　構　　　　造　" & strWK3 & vbCrLf
							strRT_WK = strRT_WK & "　床　 面　 積　" & replace(replace(strWK4,"：","㎡"),"分","分　") & vbCrLf 
					
						end if
						
						strWK1 = ""
						strWK2 = ""
						strWK3 = ""
						strWK4 = ""
						
						ichi = ichi + 1
						
					
				end if
		
		end if
		
	end if
		

Loop
'******************************************************************

objTextStream4.Close


if strRT_WK = tmpWORD & vbCrLf then
	strRT_WK = ""
end if


		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)



    '結果返す
    DAT_PUT_12_123_FUZOKU = strRT_WK

End Function



'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２と１、２、３列処理用 表題部（土地の表示）
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_12_123_B(tmpFILE, tmpWORD,tmpTASU)

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi,ichi2

Dim strWK1
Dim strWK2

Dim strWK3
Dim strWK4
Dim strWK5

Dim WK_OK

strWK1 = ""
strWK2 = ""

strWK3 = ""
strWK4 = ""
strWK5 = ""

strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
		if ichi > 0 then
		


		
						'+が２つ連続したら処理を止める
						if  instr(strText,"┃権利部") > 0  then
							ichi2 = 1
						end if
		
				if  I = ichi + tmpTASU then
				

		
					if left(strText,1) <> "+"  then
					
						'msgbox I & "→" & strText
					
						strSPL_WK = split(strText,"│")
						

						
						if strSPL_WK(0) <> "┃①地番" and strSPL_WK(1) <> "②地目" then
						
						
							'┃　があるから > 1
							if instr(strSPL_WK(0), "") = 0 and len(strSPL_WK(0)) > 1 then
								strWK1 = strSPL_WK(0)
							end if
						
						
							if instr(strSPL_WK(1), "") = 0 and len(strSPL_WK(1)) > 0 then
								strWK2 = strSPL_WK(1)
							end if
							
							'msgbox strWK1 & strWK2
						
						end if
						
						
						ichi = ichi + 1
						
						else
						
						
						ichi = ichi + 1
						
						if strSPL_WK(0) = "┃①地番" and strSPL_WK(1) = "②地目" then
						
							strRT_WK = strRT_WK & strWK1 & "　" & strWK2 & vbCrLf
												
							ichi2 = I + 1
							ichi = 1
						end if
						
						
					end if
					
					
					
				end if
				
				
				
		end if
		
		
		if ( I > 0 ) and ( I = ichi2 ) then
		
				if left(strText,1) <> "+" then
				
						strSPL_WK = split(strText,"│")
						
						WK_OK = 0
						WK_OK = instr(strSPL_WK(2),"０")
						WK_OK = WK_OK + instr(strSPL_WK(2),"１")
						WK_OK = WK_OK + instr(strSPL_WK(2),"２")
						WK_OK = WK_OK + instr(strSPL_WK(2),"３")
						WK_OK = WK_OK + instr(strSPL_WK(2),"４")
						WK_OK = WK_OK + instr(strSPL_WK(2),"５")
						WK_OK = WK_OK + instr(strSPL_WK(2),"６")
						WK_OK = WK_OK + instr(strSPL_WK(2),"７")
						WK_OK = WK_OK + instr(strSPL_WK(2),"８")
						WK_OK = WK_OK + instr(strSPL_WK(2),"９")
						
						
						'msgbox I & "|" & strText &  "|" & WK_OK 
						
						

						
						if strSPL_WK(2) <> "：" and WK_OK > 0 then
						
						'msgbox "分割｜" & strSPL_WK(0) & strSPL_WK(1) & strSPL_WK(2)
				
				
						'msgbox len(strSPL_WK(0))
				
						'余白は登録しない
						if instr(strSPL_WK(0), "") = 0 and len(strSPL_WK(0)) > 1 then
							strWK3 = strSPL_WK(0)
						end if

						
						'余白は登録しない
						if instr(strSPL_WK(1), "") = 0 and len(strSPL_WK(1)) > 0 then
							strWK4 = strSPL_WK(1)
						end if
						
						
						'余白は登録しない
						if instr(strSPL_WK(2), "") = 0 and len(strSPL_WK(2)) > 0 then
							strWK5 = strSPL_WK(2)
						end if


						'msgbox "登録｜" & strWK3 & strWK4 & strWK5
						
						end if
						
												
						ichi2 = ichi2 + 1
						
					else
					
						ichi2 = ichi2 + 1
					
				end if
		
		end if
		

Loop
'******************************************************************



objTextStream4.Close



if strRT_WK = tmpWORD & vbCrLf then
	strRT_WK = ""
end if


			strRT_WK = strRT_WK & "　地　　　　番　" & strWK3 & vbCrLf
			strRT_WK = strRT_WK & "　地　　　　目　" & strWK4 & vbCrLf
			strRT_WK = strRT_WK & "　地　　　　積　" & replace(replace(strWK5,"：","㎡"),"分","分　") & vbCrLf 


		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)



    '結果返す
    DAT_PUT_12_123_B = strRT_WK

End Function




'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２、３、４列処理用
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_1234(tmpFILE, tmpWORD,tmpTASU)

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi

Dim strMAE

strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
strMAE = ""
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"│")
						
						strRT_WK = strRT_WK & "　種　　　　類　" & strSPL_WK(0) & vbCrLf
						strRT_WK = strRT_WK & "　所在及び地番　" & strSPL_WK(1) & vbCrLf
						strRT_WK = strRT_WK & "　地　　　　目　" & strSPL_WK(2) & vbCrLf
						strRT_WK = strRT_WK & "　地　　　　積　" & replace(replace(strSPL_WK(3),"：","㎡"),"分","分　") & vbCrLf & vbCrLf
						
						ichi = ichi + 2
						
					end if
					
				end if
				
						'+が２つ連続したら処理を止める
						if  left(strText,1) = "+" and strMAE = "+"  then
							ichi = 1
						end if
				
						strMAE = left(strText,1)
				
		end if
		
		

Loop
'******************************************************************

objTextStream4.Close

if strRT_WK = tmpWORD & vbCrLf then
	strRT_WK = ""
end if



		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)


    '結果返す
    DAT_PUT_1234 = strRT_WK

End Function






'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'= １、２、３列処理用
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_PUT_123(tmpFILE, tmpWORD,tmpTASU)

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi


Dim strWK1
Dim strWK2
Dim strWK3

strWK1 = ""
strWK2 = ""
strWK3 = ""

strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' ファイルシステムオブジェクト

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0

'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'最初の位置を発見
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
						'+が２つ連続したら処理を止める
						if  instr(strText,"┃所有者") > 0  then
							ichi = 1
						end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"│")
						
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						strWK3 = strWK3 & strSPL_WK(2)
						
						ichi = ichi + 1
						
						else
						
						strRT_WK = strRT_WK & "　土地の符号　　" & strWK1 & vbCrLf
						strRT_WK = strRT_WK & "　敷地件の種類　" & strWK2 & vbCrLf
						strRT_WK = strRT_WK & "　敷地権の割合　" & strWK3 & vbCrLf & vbCrLf
						
						strWK1 = ""
						strWK2 = ""
						strWK3 = ""
						
						ichi = ichi + 1
						
					end if
					
				end if
				
				
				
		end if
		
		

Loop
'******************************************************************

objTextStream4.Close


if strRT_WK = tmpWORD & vbCrLf then
	strRT_WK = ""
end if

		'Ex. "┃所在"→"所　　　在"
		strRT_WK = DAT_CONV(strRT_WK)


    '結果返す
    DAT_PUT_123 = strRT_WK

End Function







'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
'=  "┃所在"→"所　　　在"
'■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
Function DAT_CONV(tmpWORD)

	tmpWORD = replace(tmpWORD,"表題部（一棟の建物の表示）","１（一棟の建物の表示）")
		tmpWORD = replace(tmpWORD,"表題部（主である建物の表示）","１（主である建物の表示）")
	
	tmpWORD = replace(tmpWORD,"表題部（","（")
	
	tmpWORD = replace(tmpWORD,"┃所在","　所　　　　在")
	tmpWORD = replace(tmpWORD,"┃家屋番号","　家 屋　番 号")
	tmpWORD = replace(tmpWORD,"┃建物の名称","　建物の名称　")
	tmpWORD = replace(tmpWORD,"┃","")


    '結果返す
     DAT_CONV = tmpWORD

End Function




