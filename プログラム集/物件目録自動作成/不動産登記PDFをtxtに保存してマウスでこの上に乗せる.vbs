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




msgbox "�y  " & strVER & " �z" & vbCrLf & "�s���Y�o�L���PDF���J���A�ҏW�����ׂĂ�I�����R�s�[���������ɓ\��t���ĕۑ���������" & vbCrLf & "���̃A�C�R����ɏ悹�Ă��������I" & vbCrLf & "By Skojima@kitahama.or.jp"
msgbox "�y�d�v�z���̃v���O�����͖��ۏ؂ł��I"

Set args = WScript.Arguments
For Each arg In args
    Call MsgBox(arg,,"�y�����畨���ژ^���쐬���܂��z")

    
Dim objFile    ' �Ώۃt�@�C��
Dim oldText    ' �u���O�e�L�X�g
Dim newText    ' �u����e�L�X�g
Dim objFSO     ' �t�@�C���V�X�e���I�u�W�F�N�g
Dim objRep     ' ���K�\���I�u�W�F�N�g
Dim repText    ' �u���Ώە�����

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(arg)

' �e�L�X�g�f�[�^�Ǎ�
oldText = objFile.ReadAll

' �u���Ώە�����
repText = ""

Set objRep = New RegExp

'��s���ɏ�����True
objRep.Multiline = True

'�啶����������ʂ��Ȃ���True
objRep.IgnoreCase = True

'�ŏ��Ƀ}�b�`��������������False
objRep.Global = True


repText = "+"


' ���K�\���p�^�[�����w�肷��
objRep.Pattern = "^.*����.*��"
newText = objRep.replace(oldText, repText)

oldText = newText
objRep.Pattern = "^.*����.*��"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*����.*��"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*����.*��"
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^.*��.*��"
newText = objRep.replace(oldText, repText)

oldText = newText
objRep.Pattern = "�@| "
repText = ""
newText = objRep.replace(oldText, repText)


oldText = newText
objRep.Pattern = "^[+]��"
repText = "+" & vbCrLf &"��"
newText = objRep.replace(oldText, repText)

objFile.Close

'========================================================

if right(arg,4) <> ".txt" then
msgbox "�G���[�F���̊g���q�́A���p�́u.txt�v�ł͂Ȃ��̂ŏ����𒆎~���܂��I" & vbCrLf & "��F�@����.TXT�@���啶����TXT���Ə����o���܂���I"& vbCrLf & "��F�@����.txt�@�����p�ɂ��邱��"
                        wscript.quit
end if


strWRFN = replace(arg,".txt","_��Ɨp.txt")


' ��������
Set objFile = objFSO.CreateTextFile(strWRFN)
objFile.WriteLine (newText)
objFile.Close

Set objFSO = Nothing

'=============================================================
'����ŁA�r���͏����āA�f�[�^�����󔒍s�́A�u+�v�ɒu�����ꂽ
'=============================================================




'------------------------------------
Set objFileSys2 = CreateObject("Scripting.FileSystemObject")
Set objTextStream2 = objFileSys2.OpenTextFile(strWRFN, 1)

ichi = 0

'******************************************************************
Do while objTextStream2.AtEndOfStream <> True
   strText = objTextStream2.ReadLine
   
   I = I +1

	if instr(strText,"�\�蕔�i�꓏�̌����̕\���j") > 0 then
            strShurui = 1

    end if
    
    
    if instr(strText,"�\�蕔�i��ł��錚���̕\���j") > 0 then
            strShurui = 2


    end if
    
    
    if instr(strText,"�\�蕔�i�y�n�̕\���j") > 0 then
            strShurui = 3

    end if
   
Loop
'******************************************************************

objTextStream2.Close
Set objFileSys2 = Nothing

'=============================================================='

'�@�@�@�@�@�@�@�@�@�@�Ǎ���ƃt�@�C����


strWK_ALL = "�@�@�@�@�@�@�@�@�@�@�@���@���@�ځ@�^" & vbCrLf

Select case  strShurui

case 1

            msgbox "�\�蕔�i�꓏�̌����̕\���j"

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12(strWRFN ,"�\�蕔�i�꓏�̌����̕\���j",2) 


strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123(strWRFN ,"�\�蕔�i��L�����̌����̕\���j",4)



strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_1234(strWRFN ,"�\�蕔�i�~�n���̖ړI�ł���y�n�̕\���j",4)




strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_123(strWRFN ,"�\�蕔�i�~�n���̕\���j",4)

case 2

            msgbox "�\�蕔�i��ł��錚���̕\���j"

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123(strWRFN ,"�\�蕔�i��ł��錚���̕\���j",4)

strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123_FUZOKU(strWRFN ,"�\�蕔�i���������̕\���j",4)


case 3
            msgbox "�\�蕔�i�y�n�̕\���j"


strWK_ALL = strWK_ALL & vbCrLf & DAT_PUT_12_123_B(strWRFN ,"�\�蕔�i�y�n�̕\���j",4)



case else
msgbox "���̃t�@�C���ɂ́A�Y���������ڂ������̂ŁA�����𒆎~���܂��I"
                        wscript.quit

End Select

'------------------------------------


'���ژ^�ۑ��p
strWRSAVE = replace(arg,".txt","_�ϊ���̕����ژ^.txt")

msgbox strWRSAVE

'msgbox strWK_ALL

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(strWRSAVE)
objFile.WriteLine (strWK_ALL)
objFile.Close

Set objFSO = Nothing




strWK_ALL = ""

'msgbox strWRSAVE,"�������݊���"
    
'----------------------
Next

WScript.echo "�������������܂����I"



'������������������������������������������������������������
'= �P�A�Q�񏈗��p
'������������������������������������������������������������
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

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0

'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
						'+���Q�A�������珈�����~�߂�
						if  instr(strText,"���@�\��") > 0  then
							ichi = 1
						end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"��")
						
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						
						ichi = ichi + 1
						
						else
						
						strRT_WK = strRT_WK & strWK1 & "�@" & strWK2 & vbCrLf

						
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


		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)
		
		



    '���ʕԂ�
    DAT_PUT_12 = strRT_WK

End Function




'������������������������������������������������������������
'= �P�A�Q�ƂP�A�Q�A�R�񏈗��p
'������������������������������������������������������������
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

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		

						'+���Q�A�������珈�����~�߂�  �u���\�蕔�i���������̕\���j�v�́��ʊ֐��ŏ���
						if  instr(strText,"���\�蕔�i���������̕\���j") or instr(strText,"��������") > 0 or instr(strText,"�����L��") > 0 or instr(strText,"���\�蕔�i�~�n���̕\���j") > 0 then
							ichi2 = 1
							ichi = 0
						end if
		
		
		if ichi > 0 then
		
		
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+"  then
					
						'msgbox I & "��1��" & strText
					
						strSPL_WK = split(strText,"��")
						
						
						'msgbox "2��" &strSPL_WK(0) & "|" & strSPL_WK(1)
						
						if strSPL_WK(0) <> "���@���" and strSPL_WK(1) <> "�A�\��" then
						
						
							'���@�����邩�� > 1
							if instr(strSPL_WK(0), "������") = 0 and len(strSPL_WK(0)) > 1 then
								strWK1 = strSPL_WK(0)
							end if
						
						
							if instr(strSPL_WK(1), "������") = 0 and len(strSPL_WK(1)) > 0 then
								strWK2 = strWK2 & strSPL_WK(1)
							end if
							
							'msgbox "3��" & strWK1 & strWK2
						
						end if
						
						
						ichi = ichi + 1
						
						else
						
						ichi = ichi + 1
						
						if strSPL_WK(0) = "���@���" and strSPL_WK(1) = "�A�\��" then
																		
							ichi2 = I + 1
							ichi = 1
						else
						
							'msgbox "4��" & strWK1 & strWK2
												
							strRT_WK = strRT_WK & strWK1 & "�@" & strWK2 & vbCrLf
							
							strWK1 = ""
							strWK2 = ""
						
						end if
						
						
					end if
					
					
				end if
				
		end if
		
		
		if ( I > 0 ) and ( I = ichi2 ) then
		
				if left(strText,1) <> "+" and instr(strText,"��") > 0 then
				strSPL_WK = split(strText,"��")
				
				'msgbox("--" & strText)
				

				
						strWK3 = strWK3 & strSPL_WK(0)
						strWK4 = strWK4 & strSPL_WK(1)
						
						if strSPL_WK(2) <> "�F" then
							strWK5 = strWK5 & strSPL_WK(2) & "�@"
						end if
						

						
						ichi2 = ichi2 + 1
						
					else
					
									'�󔒂͓o�^���Ȃ�
									
						'msgbox ("|" & strWK5 & "|" & len(strWK5))
						
						if instr(strWK3, "������") = 0 and len(strWK5) > 1 then
					
							strRT_WK = strRT_WK & "�@��@�@�@�@�ށ@" & strWK3 & vbCrLf
							strRT_WK = strRT_WK & "�@�\�@�@�@�@���@" & strWK4 & vbCrLf
							strRT_WK = strRT_WK & "�@���@ �ʁ@ �ρ@" & replace(replace(strWK5,"�F","�u"),"��","���@") & vbCrLf 
					
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


		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)



    '���ʕԂ�
    DAT_PUT_12_123 = strRT_WK

End Function




'������������������������������������������������������������
'= �P�A�Q�ƂP�A�Q�A�R�񏈗��p�@�t������
'������������������������������������������������������������
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

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		

						'+���Q�A�������珈�����~�߂� 
						if  instr(strText,"��������") > 0 or instr(strText,"�����L��") then
							ichi = 0
						end if
		
		
	if ichi > 0 then
		
		if ( I > 0 ) and ( I = ichi + tmpTASU ) then
		
				if left(strText,1) <> "+" and instr(strText,"��") > 0 then
				strSPL_WK = split(strText,"��")
				
				'msgbox("--" & strText)
				

				
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						strWK3 = strWK3 & strSPL_WK(2)
						
						if strSPL_WK(3) <> "�F" then
							strWK4 = strWK4 & strSPL_WK(3) & "�@"
						end if
						

						
						ichi = ichi + 1
						
					else
					
									'�󔒂͓o�^���Ȃ�
									
						'msgbox ("|" & strWK1 & "|" & len(strWK4))
						
						if instr(strWK1, "������") = 0 and len(strWK4) > 1 then
							strRT_WK = strRT_WK & "�@���@�@�@�@���@" & strWK1 & vbCrLf
							strRT_WK = strRT_WK & "�@��@�@�@�@�ށ@" & strWK2 & vbCrLf
							strRT_WK = strRT_WK & "�@�\�@�@�@�@���@" & strWK3 & vbCrLf
							strRT_WK = strRT_WK & "�@���@ �ʁ@ �ρ@" & replace(replace(strWK4,"�F","�u"),"��","���@") & vbCrLf 
					
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


		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)



    '���ʕԂ�
    DAT_PUT_12_123_FUZOKU = strRT_WK

End Function



'������������������������������������������������������������
'= �P�A�Q�ƂP�A�Q�A�R�񏈗��p �\�蕔�i�y�n�̕\���j
'������������������������������������������������������������
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

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
ichi2 = 0
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
		if ichi > 0 then
		


		
						'+���Q�A�������珈�����~�߂�
						if  instr(strText,"��������") > 0  then
							ichi2 = 1
						end if
		
				if  I = ichi + tmpTASU then
				

		
					if left(strText,1) <> "+"  then
					
						'msgbox I & "��" & strText
					
						strSPL_WK = split(strText,"��")
						

						
						if strSPL_WK(0) <> "���@�n��" and strSPL_WK(1) <> "�A�n��" then
						
						
							'���@�����邩�� > 1
							if instr(strSPL_WK(0), "������") = 0 and len(strSPL_WK(0)) > 1 then
								strWK1 = strSPL_WK(0)
							end if
						
						
							if instr(strSPL_WK(1), "������") = 0 and len(strSPL_WK(1)) > 0 then
								strWK2 = strSPL_WK(1)
							end if
							
							'msgbox strWK1 & strWK2
						
						end if
						
						
						ichi = ichi + 1
						
						else
						
						
						ichi = ichi + 1
						
						if strSPL_WK(0) = "���@�n��" and strSPL_WK(1) = "�A�n��" then
						
							strRT_WK = strRT_WK & strWK1 & "�@" & strWK2 & vbCrLf
												
							ichi2 = I + 1
							ichi = 1
						end if
						
						
					end if
					
					
					
				end if
				
				
				
		end if
		
		
		if ( I > 0 ) and ( I = ichi2 ) then
		
				if left(strText,1) <> "+" then
				
						strSPL_WK = split(strText,"��")
						
						WK_OK = 0
						WK_OK = instr(strSPL_WK(2),"�O")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�P")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�Q")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�R")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�S")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�T")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�U")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�V")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�W")
						WK_OK = WK_OK + instr(strSPL_WK(2),"�X")
						
						
						'msgbox I & "|" & strText &  "|" & WK_OK 
						
						

						
						if strSPL_WK(2) <> "�F" and WK_OK > 0 then
						
						'msgbox "�����b" & strSPL_WK(0) & strSPL_WK(1) & strSPL_WK(2)
				
				
						'msgbox len(strSPL_WK(0))
				
						'�]���͓o�^���Ȃ�
						if instr(strSPL_WK(0), "������") = 0 and len(strSPL_WK(0)) > 1 then
							strWK3 = strSPL_WK(0)
						end if

						
						'�]���͓o�^���Ȃ�
						if instr(strSPL_WK(1), "������") = 0 and len(strSPL_WK(1)) > 0 then
							strWK4 = strSPL_WK(1)
						end if
						
						
						'�]���͓o�^���Ȃ�
						if instr(strSPL_WK(2), "������") = 0 and len(strSPL_WK(2)) > 0 then
							strWK5 = strSPL_WK(2)
						end if


						'msgbox "�o�^�b" & strWK3 & strWK4 & strWK5
						
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


			strRT_WK = strRT_WK & "�@�n�@�@�@�@�ԁ@" & strWK3 & vbCrLf
			strRT_WK = strRT_WK & "�@�n�@�@�@�@�ځ@" & strWK4 & vbCrLf
			strRT_WK = strRT_WK & "�@�n�@�@�@�@�ρ@" & replace(replace(strWK5,"�F","�u"),"��","���@") & vbCrLf 


		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)



    '���ʕԂ�
    DAT_PUT_12_123_B = strRT_WK

End Function




'������������������������������������������������������������
'= �P�A�Q�A�R�A�S�񏈗��p
'������������������������������������������������������������
Function DAT_PUT_1234(tmpFILE, tmpWORD,tmpTASU)

Dim strRT_WK
Dim strSPL_WK
Dim objTextStream4

Dim I,ichi

Dim strMAE

strRT_WK = tmpWORD & vbCrLf
strSPL_WK = ""

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0
strMAE = ""
'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"��")
						
						strRT_WK = strRT_WK & "�@��@�@�@�@�ށ@" & strSPL_WK(0) & vbCrLf
						strRT_WK = strRT_WK & "�@���݋y�ђn�ԁ@" & strSPL_WK(1) & vbCrLf
						strRT_WK = strRT_WK & "�@�n�@�@�@�@�ځ@" & strSPL_WK(2) & vbCrLf
						strRT_WK = strRT_WK & "�@�n�@�@�@�@�ρ@" & replace(replace(strSPL_WK(3),"�F","�u"),"��","���@") & vbCrLf & vbCrLf
						
						ichi = ichi + 2
						
					end if
					
				end if
				
						'+���Q�A�������珈�����~�߂�
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



		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)


    '���ʕԂ�
    DAT_PUT_1234 = strRT_WK

End Function






'������������������������������������������������������������
'= �P�A�Q�A�R�񏈗��p
'������������������������������������������������������������
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

Dim objFSO4     ' �t�@�C���V�X�e���I�u�W�F�N�g

Set objFSO4 = CreateObject("Scripting.FileSystemObject")
Set objTextStream4 = objFSO4.OpenTextFile(tmpFILE, 1)


I = 0
ichi = 0

'******************************************************************
Do while objTextStream4.AtEndOfStream <> True
   strText = objTextStream4.ReadLine
   
      I = I + 1

		'�ŏ��̈ʒu�𔭌�
		if instr(strText,tmpWORD) > 0 then
			ichi = I
		end if
		
						'+���Q�A�������珈�����~�߂�
						if  instr(strText,"�����L��") > 0  then
							ichi = 1
						end if
		
		if ichi > 0 then
		
				if  I = ichi + tmpTASU then
		
					if left(strText,1) <> "+" then
						strSPL_WK = split(strText,"��")
						
						strWK1 = strWK1 & strSPL_WK(0)
						strWK2 = strWK2 & strSPL_WK(1)
						strWK3 = strWK3 & strSPL_WK(2)
						
						ichi = ichi + 1
						
						else
						
						strRT_WK = strRT_WK & "�@�y�n�̕����@�@" & strWK1 & vbCrLf
						strRT_WK = strRT_WK & "�@�~�n���̎�ށ@" & strWK2 & vbCrLf
						strRT_WK = strRT_WK & "�@�~�n���̊����@" & strWK3 & vbCrLf & vbCrLf
						
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

		'Ex. "������"��"���@�@�@��"
		strRT_WK = DAT_CONV(strRT_WK)


    '���ʕԂ�
    DAT_PUT_123 = strRT_WK

End Function







'������������������������������������������������������������
'=  "������"��"���@�@�@��"
'������������������������������������������������������������
Function DAT_CONV(tmpWORD)

	tmpWORD = replace(tmpWORD,"�\�蕔�i�꓏�̌����̕\���j","�P�i�꓏�̌����̕\���j")
		tmpWORD = replace(tmpWORD,"�\�蕔�i��ł��錚���̕\���j","�P�i��ł��錚���̕\���j")
	
	tmpWORD = replace(tmpWORD,"�\�蕔�i","�i")
	
	tmpWORD = replace(tmpWORD,"������","�@���@�@�@�@��")
	tmpWORD = replace(tmpWORD,"���Ɖ��ԍ�","�@�� ���@�� ��")
	tmpWORD = replace(tmpWORD,"�������̖���","�@�����̖��́@")
	tmpWORD = replace(tmpWORD,"��","")


    '���ʕԂ�
     DAT_CONV = tmpWORD

End Function




