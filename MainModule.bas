Attribute VB_Name = "MainModule"
Option Explicit

' �萔�̐ݒ�
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const conTempFileName = "NCView._$$" ' �e���|�����t�@�C����
Public Const intRow As Integer = 50 ' ��
Public Const int1mm As Integer = 100 ' 1mm�̒l

Public Type NCInfo
    dblMin(1) As Double ' �ŏ��l X/Y
    dblMax(1) As Double ' �ő�l X/Y
    lngOffSet(1) As Long
End Type

'Public Type WBInfo
'    intSosu As Integer ' �w��
'    lngWBS(1) As Long ' WBS X/Y
'    lngStack(1) As Long ' Stack X/Y
'    strStart As String
'End Type

Public Type ToolInfo
    intTNo As Integer ' T�R�[�h
    sngDrill As Single ' �h�����a
    lngColor As Long ' �F
End Type

'*********************************************************
' �p  �r: NCView�̃X�^�[�g�A�b�v
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Sub Main()

    Dim strNC As String
    Dim strNCFileName As String
'    Dim frmNewMain As Form
'    Dim frmNewToolInfo As Form

    ' 2�d�N�����`�F�b�N
'    If App.PrevInstance Then
'        MsgBox "���łɋN������Ă��܂��I"
'        End
'    End If

    ' ����������
    Call sInitialize

'    Set frmNewMain = New frmMain
'    Set frmNewToolInfo = New frmToolInfo
'    frmNewMain.Show
'    frmNewToolInfo.Show vbModal

'    Load frmMain
    Load frmNCInfo
'    frmMain.Show
    frmNCInfo.Show

End Sub

'*********************************************************
' �p  �r: �ϐ�������������
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Public Sub sInitialize()

    Dim i As Integer

    frmMain.ScaleFactor = 1

    ' �l�����͂���Ă��邩�ۂ��𔻒f����ׂɗL�蓾�Ȃ��l�ŏ���������


'    With gudtNCInfo(frmMain.THNT)
'        .strFileName = ""
'        .dblMax(X) = 0
'        .dblMax(Y) = 0
'        .dblMin(X) = 0
'        .dblMin(Y) = 0
'    End With

'    With gudtWBInfo
'        .intSosu = 0
'        .lngWBS(X) = 0
'        .lngWBS(Y) = 0
'        .lngStack(X) = 0
'        .lngStack(Y) = 0
'        .strStart = ""
'    End With

End Sub

'*********************************************************
' �p  �r: ���ϐ�TEMP�̒l���擾����
' ��  ��: ����
' �߂�l: ���ϐ�TEMP�̒l��Ԃ�
'*********************************************************

Public Function fTempPath() As String

    ' �v���O�����I���܂�TempPath�̓��e��ێ�
    Static TempPath As String

    ' �r���Ńf�B���N�g��-���ύX����Ă�Temp�f�B���N�g��-���m��
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP") ' �f�B���N�g��-���擾
        ' ���[�g�f�B���N�g���[���̔��f
        If right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function
