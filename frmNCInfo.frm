VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNCInfo 
   BorderStyle     =   1  '�Œ�(����)
   Caption         =   "NCMerge"
   ClientHeight    =   5310
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmNCInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6645
   Begin VB.ComboBox cmbOffSet 
      Height          =   300
      Index           =   1
      Left            =   4920
      TabIndex        =   21
      Text            =   "Combo1"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.ComboBox cmbOffSet 
      Height          =   300
      Index           =   0
      Left            =   4920
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox txtOutFile 
      Height          =   270
      Left            =   4920
      TabIndex        =   23
      Text            =   "Text2"
      Top             =   3600
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5400
      TabIndex        =   24
      Top             =   4200
      Width           =   1095
   End
   Begin VB.Frame frmTHNT 
      Caption         =   "NT"
      Height          =   4935
      Index           =   1
      Left            =   2520
      TabIndex        =   6
      Top             =   240
      Width           =   2175
      Begin VB.TextBox txtInFile 
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   1440
      End
      Begin VB.CommandButton cmdSee 
         Caption         =   "..."
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtInput 
         Height          =   264
         Index           =   1
         Left            =   360
         TabIndex        =   10
         Text            =   "000,000"
         Top             =   1800
         Width           =   732
      End
      Begin MSFlexGridLib.MSFlexGrid msgDrill 
         Height          =   3360
         Index           =   1
         Left            =   240
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   5927
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "̧�ٖ�"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame frmTHNT 
      Caption         =   "TH"
      Height          =   4935
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton cmdSee 
         Caption         =   "..."
         Height          =   255
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
      Begin VB.TextBox txtInFile 
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   600
         Width           =   1440
      End
      Begin VB.TextBox txtInput 
         Height          =   264
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Text            =   "000,000"
         Top             =   1800
         Width           =   672
      End
      Begin MSFlexGridLib.MSFlexGrid msgDrill 
         Height          =   3360
         Index           =   0
         Left            =   240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1080
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   5927
         _Version        =   393216
         Enabled         =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "̧�ٖ�"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "�ر"
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      Top             =   4800
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   4800
      _ExtentX        =   688
      _ExtentY        =   688
      _Version        =   393216
   End
   Begin VB.Label Label5 
      Caption         =   "NT�̈ړ���"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "TH�̈ړ���"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblOutFile 
      Caption         =   "̧�ٖ�(&N)"
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblMax 
      BorderStyle     =   1  '����
      Caption         =   "-999.99"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "�ő�l"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblMin 
      BorderStyle     =   1  '����
      Caption         =   "-999.99"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMax 
      BorderStyle     =   1  '����
      Caption         =   "-999.99"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblMin 
      BorderStyle     =   1  '����
      Caption         =   "-999.99"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "�ŏ��l"
      Height          =   255
      Left            =   4920
      TabIndex        =   12
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmNCInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrNC(1) As String
Private mudtTHTool(1 To intRow) As ToolInfo ' TH�̃c�[�����
Private mudtNTTool(1 To intRow) As ToolInfo
Private mudtNCInfo(1) As NCInfo ' NC���
Private mblnKeyFlag As Boolean

Private Sub cmbOffSet_LostFocus(Index As Integer)

    With cmbOffSet(Index)
        .Text = UCase(.Text)
    End With

End Sub

'*********************************************************
' �p  �r: OK�{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdOK_Click()

    Dim i As Integer
    Dim intTHNT As Integer
    Dim blnRet As Boolean
    Dim intF0 As Integer
    Dim intF1 As Integer
    Dim strTempFile As String
    Dim strOutFile As String
    Dim udtNCInfo As NCInfo
    Dim strXY() As String

    On Error GoTo ErrMsg

    ' frmMain�̃��[�h
    Load frmMain
    With frmMain
        .Show
        .ProgressBar1.Value = 0
        .ProgressBar1.Visible = True
    End With
    DoEvents

    For i = 1 To intRow
        With msgDrill(TH)
            .Row = i
            .Col = 1 ' �h�����a
            If .Text <> "" Then
                .Col = 0 ' TNo
                mudtTHTool(i).intTNo = CInt(.Text)
                .Col = 1 ' �h�����a
                mudtTHTool(i).sngDrill = CSng(.Text) / 2
                mudtTHTool(i).lngColor = RGB(255, 0, 0) ' TH�͐Ԃŕ`��
            End If
        End With
        With msgDrill(NT)
            .Row = i
            .Col = 1 ' �h�����a
            If .Text <> "" Then
                .Col = 0 ' TNo
                mudtNTTool(i).intTNo = CInt(.Text)
                .Col = 1 ' �h�����a
                mudtNTTool(i).sngDrill = CSng(.Text) / 2
                mudtNTTool(i).lngColor = RGB(0, 0, 255) ' NT�͐ŕ`��
            End If
        End With
    Next

    ' �I�t�Z�b�g�̐ݒ�
    If cmbOffSet(TH).Text Like "X*Y*" = True Then
        strXY = Split(Mid(cmbOffSet(TH).Text, 2), "Y", -1, vbTextCompare)
        With mudtNCInfo(TH)
            .lngOffSet(X) = CLng(strXY(X))
            .lngOffSet(Y) = CLng(strXY(Y))
        End With
    End If
    If cmbOffSet(NT).Text Like "X*Y*" = True Then
        strXY = Split(Mid(cmbOffSet(NT).Text, 2), "Y", -1, vbTextCompare)
        With mudtNCInfo(NT)
            .lngOffSet(X) = CLng(strXY(X))
            .lngOffSet(Y) = CLng(strXY(Y))
        End With
    End If

    ' �e���|�����t�@�C���̃I�[�v��
    strTempFile = fTempPath & conTempFileName
    intF0 = FreeFile
    Open strTempFile For Output As #intF0
    ' TH�̓W�J
    frmMain.StatusBar1.Panels(1).Text = "TH��W�J��..."
    DoEvents
    blnRet = fConvertNC(mstrNC(TH), _
                        mudtTHTool, _
                        mudtNCInfo(TH), _
                        intF0, _
                        frmMain.ProgressBar1)
    ' NT�̓W�J
    If txtInFile(NT).Text <> "" Then
        frmMain.StatusBar1.Panels(1).Text = "NT��W�J��..."
        DoEvents
        blnRet = fConvertNC(mstrNC(NT), _
                            mudtNTTool, _
                            mudtNCInfo(NT), _
                            intF0, _
                            frmMain.ProgressBar1)
    End If
    Close #intF0
    If blnRet = False Then Exit Sub

    udtNCInfo.dblMin(X) = fSmall(mudtNCInfo(TH).dblMin(X), mudtNCInfo(NT).dblMin(X))
    udtNCInfo.dblMin(Y) = fSmall(mudtNCInfo(TH).dblMin(Y), mudtNCInfo(NT).dblMin(Y))
    udtNCInfo.dblMax(X) = fLarge(mudtNCInfo(TH).dblMax(X), mudtNCInfo(NT).dblMax(X))
    udtNCInfo.dblMax(Y) = fLarge(mudtNCInfo(TH).dblMax(Y), mudtNCInfo(NT).dblMax(Y))

    ' �o�͗pNC�t�@�C���̃I�[�v��
    frmMain.StatusBar1.Panels(1).Text = "��撆..."
    strOutFile = txtOutFile.Text
    intF1 = FreeFile
    Open strOutFile For Output As #intF1
    Call sDispNC(frmMain.picDraw, _
                 udtNCInfo, _
                 intF1, _
                 frmMain.ProgressBar1)
    Close #intF1
    With frmMain
        .ProgressBar1.Visible = False
        .StatusBar1.Panels(1).Text = ""
    End With

    lblMin(X).Caption = Format(udtNCInfo.dblMin(X), "##0.00")
    lblMin(Y).Caption = Format(udtNCInfo.dblMin(Y), "##0.00")
    lblMax(X).Caption = Format(udtNCInfo.dblMax(X), "##0.00")
    lblMax(Y).Caption = Format(udtNCInfo.dblMax(Y), "##0.00")

    Exit Sub

ErrMsg:
    Close #intF0
    Close #intF1
    frmMain.ProgressBar1.Visible = False
    Unload frmMain
    MsgBox "̧�ق���ݏo���܂���B"

End Sub

'*********************************************************
' �p  �r: �N���A�{�^���̃N���b�N�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub cmdClear_Click()

    Dim i As Integer

    Call sInit(TH)
    Call sInit(NT)
    cmbOffSet(TH).Text = ""
    cmbOffSet(NT).Text = ""
    txtOutFile.Text = ""
    
    mstrNC(TH) = ""
    mstrNC(NT) = ""

    With mudtNCInfo(TH)
        .dblMax(X) = 0
        .dblMax(Y) = 0
        .dblMin(X) = 0
        .dblMin(Y) = 0
        .lngOffSet(X) = 0
        .lngOffSet(Y) = 0
    End With
    With mudtNCInfo(NT)
        .dblMax(X) = 0
        .dblMax(Y) = 0
        .dblMin(X) = 0
        .dblMin(Y) = 0
        .lngOffSet(X) = 0
        .lngOffSet(Y) = 0
    End With

    For i = 1 To intRow
        With mudtTHTool(i)
            .intTNo = -1
            .sngDrill = -1
            .lngColor = -1
        End With
        With mudtNTTool(i)
            .intTNo = -1
            .sngDrill = -1
            .lngColor = -1
        End With
    Next

End Sub

Private Sub cmdSee_Click(Index As Integer)

    Dim strFileName As String

    Call sInit(Index)
    strFileName = fGetInputFile()
    If strFileName = "" Then Exit Sub

    txtInFile(Index).Text = Dir(strFileName)
    mstrNC(Index) = fReadNC(strFileName)
    Call sSetUsedTool(mstrNC(Index), Index)
    cmbOffSet(Index).Text = fGetNTIdou(mstrNC(Index))

End Sub

'*********************************************************
' �p  �r: frmToolInfo��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    ' �O��I������Form�̈ʒu�𕜌�
    top = GetSetting("NCMRG", "Info", "Top", 0)
    left = GetSetting("NCMRG", "Info", "Left", 0)

    mblnKeyFlag = True

    ' �R���{�{�b�N�X�̐ݒ�
    With cmbOffSet(TH)
        .AddItem "X0Y-2500"
        .AddItem "X0Y-25200"
    End With
    With cmbOffSet(NT)
        .AddItem "X100Y16600"
        .AddItem "X100Y20000"
        .AddItem "X100Y20500"
        .AddItem "X100Y25200"
    End With

    ' �N���b�N�C�x���g�𔭐������ď���������
    cmdClear.Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Form�̈ʒu�����W�X�g���ɕۑ�
    SaveSetting "NCMRG", "Info", "Top", top
    SaveSetting "NCMRG", "Info", "Left", left

End Sub

'*********************************************************
' �p  �r: �t���L�V�u���O���b�h��Click�C�x���g
' ��  ��: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub msgDrill_Click(Index As Integer)

    With msgDrill(Index)
        If .Col = 0 Then
            txtInput(Index).MaxLength = 3
        ElseIf .Col = 1 Then
            txtInput(Index).MaxLength = 5
        End If
    End With

    Select Case msgDrill(Index).Col
        Case 1
            With txtInput(Index)
                .Width = msgDrill(Index).CellWidth
                .Height = msgDrill(Index).CellHeight
                .Text = msgDrill(Index).Text
                .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
                      msgDrill(Index).CellTop + msgDrill(Index).top
                .SelStart = 0
                .SelLength = Len(.Text)
                .Visible = True
                .SetFocus
            End With
    End Select

End Sub

'*********************************************************
' �p  �r: �t���L�V�u���O���b�h��Scroll�C�x���g
' ��  ��: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub msgDrill_Scroll(Index As Integer)

    ' �R���g���[�����\�������O�ɃC�x���g����������ƃG���[�ɂȂ�̂Ńg���b�v����(-_-;
    On Error GoTo bye

    msgDrill(Index).SetFocus ' TextBox��Focus�����鎞��Scroll�����Focus���R�}���h�{�^���ɔ��ł��܂���
    txtInput(Index).Visible = False

bye:

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��Change�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
' �߂�l: ����
'*********************************************************

Private Sub txtInput_Change(Index As Integer)

    With msgDrill(Index)
        .CellAlignment = 1
        .Text = txtInput(Index).Text
    End With

End Sub

'*********************************************************
' �p  �r: ���͗p�e�L�X�g�{�b�N�X��KeyDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index�v���p�e�B
'         KeyCode: �L�[ �R�[�h�������萔
'         Shift: �C�x���g��������Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    ' �������͂���Ă��Ȃ����͎��ɐi�܂Ȃ�(�߂�̂͋�����)
    If txtInput(Index).Text = "" And KeyCode <> vbKeyUp Then Exit Sub

    ' Enter����, Ctrl-M����, �����L�[
    With msgDrill(Index)
        If KeyCode = vbKeyReturn Or _
           (Shift = 2 And KeyCode = vbKeyM) Or _
           KeyCode = vbKeyDown Then
                mblnKeyFlag = False
                .Text = txtInput(Index).Text
                If .Row < intRow - 1 Then
                    .Row = .Row + 1
                End If
        ElseIf KeyCode = vbKeyUp Then
            With msgDrill(Index)
                If .Row > 1 Then
                    .Row = .Row - 1
                End If
            End With
        Else
            Exit Sub
        End If

        If .Col = 0 Then ' TNo.
            txtInput(Index).MaxLength = 3
        Else
            txtInput(Index).MaxLength = 5
        End If

        With txtInput(Index)
            With txtInput(Index)
                .Width = msgDrill(Index).CellWidth
                .Height = msgDrill(Index).CellHeight
                .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
                      msgDrill(Index).CellTop + msgDrill(Index).top
                    .Visible = True
            End With
            .SetFocus
            .Text = msgDrill(Index).Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End With

End Sub

'*********************************************************
' �p  �r: �g�p����Ă���T�R�[�h�𒲂ׂăO���b�h�R���g���[����
'         �Z�b�g����
' ��  ��: strNC: NC�f�[�^
' �߂�l: ����
'*********************************************************

Private Sub sSetUsedTool(ByVal strNC As String, _
                         ByVal Index As Integer)

    Dim i As Integer
    Dim objReg As New RegExp
    Dim objMatches As Object
    Dim objMatch As Object

    objReg.Global = True
    objReg.IgnoreCase = False ' �啶������������ʂ���
    objReg.Pattern = "T[0-9]+"
    Set objMatches = objReg.Execute(strNC)

    ' T�R�[�h���H����ɃZ�b�g����
    i = 1
    With msgDrill(Index)
        For Each objMatch In objMatches
            .Row = i
            .Col = 0
            .Text = Mid(objMatch.Value, 2)
            .Col = 1 ' �h�����a
            .Text = "1.000" ' �f�t�H���g�̃h�����a
            i = i + 1
        Next
        ' �f�t�H���g�̈ʒu�ɃZ�b�g
        .Row = 1
        .Col = 1
        With Me.txtInput(Index)
            .SetFocus
            .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
                  msgDrill(Index).CellTop + msgDrill(Index).top
            .Text = Me.msgDrill(Index).Text
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End With

End Sub

'*********************************************************
' �p  �r: �R���g���[���̏�����
' ��  ��: TH/NT�������l(TH - 1, NT - 2)
' �߂�l: ����
'*********************************************************

Private Sub sInit(Index As Integer)

    Dim i As Integer

    With msgDrill(Index) ' �O���b�h�̏�����
        .Cols = 2
        .Rows = intRow + 1 ' +1���Ă���̂͌Œ�s�������
        .FixedCols = 0 ' �Œ��Ȃ�
        .FixedRows = 1 ' �Œ�s1
        .Width = 1440
        .Height = 3600
        .RowHeight(-1) = 264 ' �S��̍���
        .RowHeight(0) = 240 ' �Œ��̍���
        .ColWidth(0) = 456 ' TNo.�̌���
        .ColWidth(1) = 624 ' �h�����a�̌���
        .FocusRect = flexFocusNone ' �t�H�[�J�X����������\�����Ȃ�
        .HighLight = flexHighlightNever ' �����\�����Ȃ�
        .Row = 0 ' ���
        .Col = 0
        .Text = "TNo."
        .Col = 1
        .Text = "���ٌa"
        For i = 1 To intRow
            .Row = i
            .Col = 0 ' TNo�̌�
            .CellAlignment = 1 ' �����̒���
            .Text = ""
            .Col = 1 ' �h�����a�̌�
            .CellAlignment = 1 ' �����̒���
            .Text = ""
        Next
    End With
    msgDrill(Index).Row = 1

    With txtInput(Index) ' �e�L�X�g�{�b�N�X�̏�����
        .ZOrder 0 ' �őO�ʂֈړ�
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Width = msgDrill(Index).CellWidth
        .Height = msgDrill(Index).CellHeight
        .Appearance = 0 ' �t���b�g
        .Alignment = vbLeftJustify ' ����
        If msgDrill(Index).Text <> "" Then
            .Text = msgDrill(Index).Text
        Else
            .Text = ""
        End If
        .MaxLength = 5
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    With lblMin(X)
        .Caption = ""
        .Alignment = vbRightJustify
    End With
    With lblMin(Y)
        .Caption = ""
        .Alignment = vbRightJustify
    End With
    With lblMax(X)
        .Caption = ""
        .Alignment = vbRightJustify
    End With
    With lblMax(Y)
        .Caption = ""
        .Alignment = vbRightJustify
    End With

    txtInFile(Index).Text = ""
'    txtOutFile.Text = ""
'    cmbOffSet.Text = ""

End Sub

Private Function fSmall(ByVal dblA As Double, _
                        ByVal dblB As Double) As Double

    If dblA < dblB Then
        fSmall = dblA
    Else
        fSmall = dblB
    End If

End Function

Private Function fLarge(ByVal dblA As Double, _
                        ByVal dblB As Double) As Double

    If dblA > dblB Then
        fLarge = dblA
    Else
        fLarge = dblB
    End If

End Function

'*********************************************************
' �p  �r: �t�@�C�����J���_�C�A���O��\������
' ��  ��: ����
' �߂�l: �I�������t�@�C����
'*********************************************************

Public Function fGetInputFile() As String

'    Dim strPathName() As String

    ' CancelError �v���p�e�B��^ (True) �ɐݒ肵�܂��B
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler

    ' �t�@�C���̑I����@��ݒ肵�܂��B
    CommonDialog1.Filter = "���ׂẴt�@�C�� (*.*)|*.*|�e�L�X�g" & "�t�@�C�� (*.txt)|*.txt|�f�[�^�t�@�C�� (*.dat)|*.dat"

    ' ����̑I����@���w�肵�܂��B
    CommonDialog1.FilterIndex = 1

    ' [�t�@�C�����J��] �_�C�A���O �{�b�N�X��\�����܂��B
    CommonDialog1.ShowOpen

    ' �t�@�C���̗L�����`�F�b�N����
    If Dir(CommonDialog1.FileName) = "" Then
        MsgBox "�t�@�C����������܂���B"
        Exit Function
    End If

'    strPathName = Split(CommonDialog1.FileName, "\", -1)
'    ' �t�@�C�������폜����
'    strPathName(UBound(strPathName)) = ""
'    ' �J�����g�f�B���N�g�����ړ�����
'    ChDir (Join(strPathName, "\"))
    fGetInputFile = CommonDialog1.FileName
    Exit Function

ErrHandler:
    fGetInputFile = ""

End Function

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)

    If mblnKeyFlag = False Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtInput_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    mblnKeyFlag = True

End Sub

Private Sub txtOutFile_LostFocus()

    With txtOutFile
        .Text = UCase(.Text)
    End With

End Sub
