'===============================================
' �t�@�C���� hide_spyware.vbs
' Shift-JIS�ŕۑ�����
'
' �E�B���h�E�Y�̃X�p�C�p�b�`���\���ɂ���
' Hidden Spy Patch Windows Update programs
' http://www.neko.ne.jp/~freewing/
' Copyright (c)2016 FREE WING, Y.Sakamoto
'
' �G�N�X�v���[���[����
' hide_spyware.vbs
' ���E�N���b�N�Łu�R�}���h�v�����v�g�ŊJ���v��I���A
' ���_�C�A���O���o����u�͂��v�������Ď��s����B
'
' Base program Searching, Downloading, and Installing Updates
' https://msdn.microsoft.com/en-us/library/aa387102.aspx
'===============================================
' UAC Permission elevation from VBScript
' http://stackoverflow.com/questions/13296281/permission-elevation-from-vbscript
'===============================================
Dim OSList, OS, UAC
UAC = False
If WScript.Arguments.Count >= 1 Then
    If WScript.Arguments.Item(0) = "elevated" Then UAC = True
End If

If Not(UAC) Then
    Set OSList = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")
    For Each OS In OSList
        If InStr(1, OS.Caption, "XP") = 0 And InStr(1, OS.Caption, "Server 2003") = 0 Then
            CreateObject("Shell.Application").ShellExecute "cscript.exe", """" & WScript.ScriptFullName & """ elevated" , "", "runas", 1
            WScript.Quit
        End If
    Next
End If

' ���L�̍X�V�v���O�����͔�\���ɂ��Ȃ��B

' ���L�̍X�V�v���O�������\���ɂ���B
spyListStr = "" _
  & "2876229:Microsoft Update �p Skype" _
  & ",2506928:Outlook �ŊJ���� .html �t�@�C���̃����N���@�\���Ȃ�" _
  & ",2545698:�ꕔ�̃R�A �t�H���g�̃e�L�X�g���ڂ₯�ĕ\�������" _
  & ",2660075:the time zone is set to Samoa (UTC+13:00) and KB 2657025" _
  & ",2726535:���̃��X�g�ɓ�X�[�_����ǉ�����" _
  & ",2970228:���V�A ���[�u���̐V�����ʉ݋L�����T�|�[�g" _
  & ",2592687:�����[�g �f�X�N�g�b�v�֌W RDP 8.0�p" _
  & ",2923545:�����[�g �f�X�N�g�b�v�֌W RDP 8.1�p" _
  & ",2994023:�����[�g �f�X�N�g�b�v�֌W RDP 8.1�p�C��" _
  & ",2952664:Win10�A�b�v�O���[�h�֌W" _
  & ",2990214:Win10�A�b�v�O���[�h�֌W" _
  & ",3035583:Win10�A�b�v�O���[�h�֌W" _
  & ",3123862:GWX 2016/02 Get Windows 10 Win10�A�b�v�O���[�h�֌W" _
  & ",3021917:CEIP���e�����g��(���u�����W)�֘A" _
  & ",3022345:�e�����g���֘A" _
  & ",3068708:�e�����g���֘A" _
  & ",3075249:�e�����g���֘A" _
  & ",3080149:�e�����g���֘A" _
  & ",3050265:WUC 2015/06 Windows Update Client" _
  & ",3065987:WUC 2015/07 Windows Update Client" _
  & ",3075851:WUC 2015/08 Windows Update Client" _
  & ",3083324:WUC 2015/09 Windows Update Client" _
  & ",3083710:WUC 2015/10 Windows Update Client" _
  & ",3102810:WUC 2015/11 Windows Update Client 7.6.7601.19046" _
  & ",3112343:WUC 2015/12 Windows Update Client 7.6.7601.19077" _
  & ",3135445:WUC 2016/02 Windows Update Client 7.6.7601.19116" _
  & ",2977759:Windows 7 �� rtm�� Windows CEIP �J�X�^�}�[�G�N�X�y���G���X����v���O����" _
  & ",3008273:Windows 8 to 8.1 Update" _
  & ",3065988:Windows 8.1 Windows Server 2012 R2 �p Windows Update �N���C�A���g" _
  & ",2976978:Windows 8.1 Windows CEIP �J�X�^�}�[ �G�N�X�y���G���X����v���O����" _
  & ",3044374:Windows 8.1 to Windows 10 Update that enables you to upgrade"

spySplit = Split(spyListStr, ",")

Set objShell = WScript.CreateObject("WScript.Shell")

For J = 0 To UBound(spySplit)
    spyDatas = Split(spySplit(J), ":")
    ' �A���C���X�g�[������
    cmdStr = "wusa.exe /uninstall /kb:" & spyDatas(0) & " /quiet /norestart"
    WScript.Echo " Uninstall > KB" & spyDatas(0) & " " & cmdStr
    objShell.Run cmdStr,,True
Next

Set objShell = Nothing

Set updateSession = CreateObject("Microsoft.Update.Session")
updateSession.ClientApplicationID = "MSDN Sample Script"

Set updateSearcher = updateSession.CreateUpdateSearcher()

WScript.Echo "Searching for updates..." & vbCRLF

Set searchResult = _
updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

WScript.Echo "List of applicable items on the machine:"

For I = 0 To searchResult.Updates.Count-1
    Set update = searchResult.Updates.Item(I)
    WScript.Echo I + 1 & "> " & update.Title

    For J = 0 To UBound(spySplit)
        spyDatas = Split(spySplit(J), ":")
        If InStr(update.Title, spyDatas(0)) > 0 Then
            ' �X�V�v���O�������\���ɂ���
            update.IsHidden = true
            WScript.Echo " IsHidden > " & update.Title
        End If
    Next
Next

If searchResult.Updates.Count = 0 Then
    WScript.Echo "There are no applicable updates."
    WScript.Quit
End If
