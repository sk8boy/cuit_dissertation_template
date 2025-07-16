VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BaseInfoForm 
   Caption         =   "���Ļ�����Ϣ"
   ClientHeight    =   6570
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8010
   OleObjectBlob   =   "BaseInfoForm.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "BaseInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub OkBtn_Click() ' ȷ����ť
    Dim ur  As UndoRecord

    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "���»�����Ϣ"

    titleCN = tbTitleCN.Value
    titleEN = tbTitleEN.Value
    studentName = tbName.Value
    studentNo = tbStudentNo.Value
    firstTeacherName = tbFirstTeacherName.Value
    firstTeacherTitle = tbFirstTeacherTitle.Value
    otherTeacherName = tbOtherTeacherName.Value
    otherTeacherTitle = tbOtherTeacherTitle.Value

    UpdateContentControl "����������Ŀ", Trim(titleCN)
    UpdateContentControl "������Ŀ", Trim(titleCN)
    UpdateContentControl "ժҪ����������Ŀ", Trim(titleCN)

    UpdateContentControl "����Ӣ����Ŀ", Trim(titleEN)
    UpdateContentControl "ժҪ����Ӣ����Ŀ", Trim(titleEN)


    UpdateContentControl "ѧ������", Trim(studentName)

    UpdateContentControl "ѧ��ѧ��", Trim(studentNo)
    UpdateContentControl "���", Trim(studentNo)

    UpdateContentControl "��һ��ʦ����", Trim(firstTeacherName)
    UpdateContentControl "ָ����ʦ", Trim(firstTeacherName)

    UpdateContentControl "��һ��ʦְ��", Trim(firstTeacherTitle)

    UpdateContentControl "������ʦ����", Trim(otherTeacherName)

    UpdateContentControl "������ʦְ��", Trim(otherTeacherTitle)

    Unload Me ' �رմ���

    ur.EndCustomRecord
    Exit Sub

ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "���»�����Ϣʱ��������: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub CancelBtn_Click() ' ȡ����ť
    Unload Me
End Sub

Private Function GetContentControl(title As String) As String
    Dim cc As ContentControl
    
    ' ͨ������(Title)���Ҳ��������ݿؼ�
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    GetContentControl = cc.Range.text
End Function

Private Sub UpdateContentControl(title As String, val As String)
    Dim cc As ContentControl
    
    ' ͨ������(Title)���Ҳ��������ݿؼ�
    On Error Resume Next
    Set cc = ActiveDocument.SelectContentControlsByTitle(title).item(1)
    On Error GoTo 0
    
    If Not cc Is Nothing Then
        ' ����ʹ�����·�ʽ���ô��ı����ݿؼ���ֵ
        cc.LockContents = False ' �Ƚ���(�����Ҫ)
        cc.Range.text = val
        cc.LockContents = True ' ��������(�����Ҫ)
        
        'MsgBox "���ݿؼ��Ѹ���!", vbInformation
    Else
        MsgBox "δ�ҵ�ָ����������ݿؼ�!", vbExclamation
    End If
End Sub


Private Sub UserForm_Initialize()
    titleCN = GetContentControl("����������Ŀ")
    titleEN = GetContentControl("����Ӣ����Ŀ")
    studentName = GetContentControl("ѧ������")
    studentNo = GetContentControl("ѧ��ѧ��")
    firstTeacherName = GetContentControl("��һ��ʦ����")
    firstTeacherTitle = GetContentControl("��һ��ʦְ��")
    otherTeacherName = GetContentControl("������ʦ����")
    otherTeacherTitle = GetContentControl("������ʦְ��")
    
    If titleCN <> "����������Ŀ" Then
        tbTitleCN.Value = titleCN
    End If
    
    If titleEN <> "����Ӣ����Ŀ" Then
        tbTitleEN.Value = titleEN
    End If
    
    If studentName <> "ѧ������" Then
        tbName.Value = studentName
    End If
    
    If studentNo <> "ѧ��" Then
        tbStudentNo.Value = studentNo
    End If
    
    If firstTeacherName <> "��ʦ����" Then
        tbFirstTeacherName.Value = firstTeacherName
    End If
    
    If firstTeacherTitle <> "ְ��" Then
        tbFirstTeacherTitle.Value = firstTeacherTitle
    End If
    
    If otherTeacherName <> "��ʦ����" Then
        tbOtherTeacherName.Value = otherTeacherName
    End If
    
    If otherTeacherTitle <> "ְ��" Then
        tbOtherTeacherTitle.Value = otherTeacherTitle
    End If
End Sub

