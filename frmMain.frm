VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   Caption         =   "Form1"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5415
   ForeColor       =   &H80000007&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":030A
   ScaleHeight     =   4215
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   1920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   3600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '������
      Height          =   270
      Left            =   0
      TabIndex        =   4
      Top             =   3945
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFrame 
      BackColor       =   &H8000000C&
      Height          =   732
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   1275
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.PictureBox picDraw 
         Height          =   372
         Index           =   1
         Left            =   720
         MouseIcon       =   "frmMain.frx":0614
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   3
         Top             =   120
         Width           =   372
      End
      Begin VB.PictureBox picDraw 
         Height          =   372
         Index           =   0
         Left            =   120
         MouseIcon       =   "frmMain.frx":091E
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   1
         Top             =   120
         Width           =   372
      End
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1320
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C28
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0D3A
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuView 
      Caption         =   "�\��"
      Visible         =   0   'False
      Begin VB.Menu mnuSeisun 
         Caption         =   "�����\��"
      End
      Begin VB.Menu mnuStandard 
         Caption         =   "�S�̕\��"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conCaption As String = "NCMerge"

Private Type tagRECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

' �E�B���h�E�̋�`�T�C�Y���擾
Private Declare Function GetWindowRect Lib "user32.dll" _
    (ByVal hwnd As Long, _
     lpRect As tagRECT) As Long

' �}�E�X�J�[�\���̈ړ��͈͂��w�肷��֐�
Private Declare Function ClipCursor Lib "user32.dll" _
    (lpRect As Any) As Long
' �V�X�e���̐ݒ��V�X�e�����g���b�N�̒l���擾����֐�
Private Declare Function GetSystemMetrics Lib "user32.dll" _
    (ByVal nIntex As Long) As Long

Const SM_CYCAPTION = 4

Private mblnDisp As Boolean ' ��ʕ\���ォ�ۂ�������Flag
Private mblnPanMode As Boolean
Private msngDragDistX As Single
Private msngDragDistY As Single
Private msngCurrentTop As Single
Private msngCurrentLeft As Single
Private mblnHMove As Boolean
Private mblnVMove As Boolean
Private mstrStart As String
Private mdblScaleFactor As Double ' �s�N�`���{�b�N�X�ɕ\�����鎞�̃t�@�N�^(�v���p�e�B�p)

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�ɕ\������ׂ̃X�P�[���t�@�N�^
'         (ScaleFactor�v���p�e�B)�̎擾
' ��  ��: ����
' �߂�l: ScaleFactor�v���p�e�B�̒l
'*********************************************************

Public Property Get ScaleFactor() As Double

    ScaleFactor = mdblScaleFactor

End Property

'*********************************************************
' �p  �r: ScaleFactor�v���p�e�B�Ƀs�N�`���{�b�N�X�ɕ\������
'         �ׂ̃X�P�[���t�@�N�^���Z�b�g
' ��  ��: dblScaleFactor: �X�P�[���t�@�N�^
' �߂�l: ����
'*********************************************************

Public Property Let ScaleFactor(ByVal dblScaleFactor As Double)

    mdblScaleFactor = dblScaleFactor

End Property

'*********************************************************
' �p  �r: frmMain��Unload�C�x���g
' ��  ��: Cancel: �t�H�[������ʂ���������邩�ǂ������w�肷��
'                 �����l(0�ŏ���, ���̑��͏������Ȃ�)
' �߂�l: ����
'*********************************************************

Private Sub Form_Unload(Cancel As Integer)

'    Unload Me

    If Dir(fTempPath & conTempFileName) <> "" Then
        Kill fTempPath & conTempFileName ' �e���|�����t�@�C�����폜
    End If

    ' Form�̈ʒu�Ƒ傫�������W�X�g���ɕۑ�
    SaveSetting "NCMRG", "Viewer", "Top", top
    SaveSetting "NCMRG", "Viewer", "Left", left
    SaveSetting "NCMRG", "Viewer", "Height", Height
    SaveSetting "NCMRG", "Viewer", "Width", Width

End Sub

Private Sub mnuSeisun_Click()

    picDraw(0).Visible = False ' �S�̕\��
    picDraw(1).Visible = True ' �����\��

End Sub

Private Sub mnuStandard_Click()

    picDraw(0).Visible = True ' �S�̕\��
    picDraw(1).Visible = False ' �����\��

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw()��KeyDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         KeyCode: �L�[ �R�[�h�������萔
'         Shift: �C�x���g��������Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyUp
            picDraw(Index).top = picDraw(Index).top - 200
        Case vbKeyDown
            picDraw(Index).top = picDraw(Index).top + 200
        Case vbKeyLeft
            picDraw(Index).left = picDraw(Index).left - 200
        Case vbKeyRight
            picDraw(Index).left = picDraw(Index).left + 200
    End Select

End Sub

'*********************************************************
' �p  �r: frmMain��Load�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Load()

    ' �O��I������Form�̈ʒu�Ƒ傫���𕜌�
    top = GetSetting("NCMRG", "Viewer", "Top", 0)
    left = GetSetting("NCMRG", "Viewer", "Left", 0)
    Height = GetSetting("NCMRG", "Viewer", "Height", Height)
    Width = GetSetting("NCMRG", "Viewer", "Width", Width)

    ' �v���O���X�o�[���\���ɂ���
    ProgressBar1.Visible = False

    ' �^�C�g��
    Caption = conCaption

    ' �t�H�[���̏�����
    Call sInit

    ' �s�N�`���[�{�b�N�X�̕\��/��\���̐ݒ�
    picDraw(0).Visible = True
    picDraw(1).Visible = False

End Sub

'*********************************************************
' �p  �r: frmMain��Resize�C�x���g
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub Form_Resize()

    With picFrame
        .Align = 3 ' ������
        .Align = 1 ' �㑵��
    End With

    If mblnDisp = False Then
        With picDraw(0)
'            .Width = picFrame.Width
'            .Height = picFrame.Height
'            .top = -24
'            .left = -24
        End With
    End If

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseDown�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: �{�^���������ꂽ����Shift, Ctrl, Alt�L�[��
'                ��Ԃ����������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim udtRect As tagRECT

    If Button = 1 Then
        MousePointer = vbCustom '�}�E�X�J�[�\����ύX
        mblnPanMode = True
        mblnHMove = True
        mblnVMove = True
        msngDragDistX = X
        msngDragDistY = Y
        msngCurrentTop = picDraw(Index).top
        msngCurrentLeft = picDraw(Index).left

        ' �s�N�`���{�b�N�X�̋�`�̈���擾
        GetWindowRect picFrame.hwnd, udtRect
        ' �擾�����̈�Ƀ}�E�X�̈ړ��͈͂𐧌�
        ClipCursor udtRect
    ElseIf Button = 2 Then
        PopupMenu mnuView
    End If

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseMove�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: Shift, Ctrl, Alt�L�[�̏�Ԃ����������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If mblnPanMode = False Then Exit Sub

    Dim udtRect As tagRECT
    Dim dblFactor As Double

    If Index = 1 Then
        dblFactor = 1
    Else
        dblFactor = Me.ScaleFactor
    End If

    ' �s�N�`���[�{�b�N�X���R���e�i����͂ݏo���Ȃ��l�ɂ���
    With picDraw(Index)
        If .Width >= picFrame.Width Then
            If .left > -24 Then
                mblnHMove = False
                .left = -24
            ElseIf .left < picFrame.Width - .Width - 24 Then
                mblnHMove = False
                .left = picFrame.Width - .Width - 24
            End If
        ElseIf .Width < picFrame.Width Then
            If .left < -24 Then ' ����
                mblnHMove = False
                .left = -24
            ElseIf .left > picFrame.Width - .Width - 24 Then
                mblnHMove = False
                .left = picFrame.Width - .Width - 24
            End If
        End If
        If .Height >= picFrame.Height Then
            If .top > -24 Then
                mblnVMove = False
                .top = -24
            ElseIf .top < picFrame.Height - .Height - 24 Then
                mblnVMove = False
                .top = picFrame.Height - .Height - 24
            End If
        ElseIf .Height < picFrame.Height Then
            If .top < -24 Then
                mblnVMove = False
                .top = -24
            ElseIf .top > picFrame.Height - .Height - 24 Then
                mblnVMove = False
                .top = picFrame.Height - .Height - 24
            End If
        End If
    End With

    ' left, top�v���p�e�B��twip�P�ʂł��鎖�ɒ���!
    If mblnHMove = True Then
    picDraw(Index).left = _
        -(msngDragDistX - X) * 56.7 / dblFactor + msngCurrentLeft
    End If
    If mblnVMove = True Then
    picDraw(Index).top = _
        (msngDragDistY - Y) * 56.7 / dblFactor + msngCurrentTop
    End If
    msngCurrentLeft = picDraw(Index).left
    msngCurrentTop = picDraw(Index).top

End Sub

'*********************************************************
' �p  �r: frmMain.picDraw��MouseUp�C�x���g
' ��  ��: Index: �R���g���[���z���Index
'         Button: �����ꂽ�{�^�������������l
'         Shift: �����ꂽ����Shift, Ctrl, Alt�L�[�̏�Ԃ�
'                ���������l
'         X, Y: �}�E�X�|�C���^�̌��݈ʒu��\�����l
' �߂�l: ����
'*********************************************************

Private Sub picDraw_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    mblnPanMode = False
    mblnHMove = False
    mblnVMove = False

    If Button = 1 Then ' ���{�^��
        MousePointer = vbDefault ' �}�E�X�J�[�\�����f�t�H���g�ɖ߂�

        ' ������NULL���w�肷�邱�Ƃ�
        ' �}�E�X�J�[�\���̈ړ�����������
        ClipCursor ByVal 0
    End If

End Sub

'*********************************************************
' �p  �r: �s�N�`���{�b�N�X�̏�����
' ��  ��: ����
' �߂�l: ����
'*********************************************************

Private Sub sInit()

    mblnDisp = False

    With picFrame
        .Align = 3 ' ������
        .Align = 1 ' �㑵��
    End With

    With picDraw(0) ' �S�̕\���p�s�N�`���[�{�b�N�X
        ' �w�i����
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' ��
        picDraw(0).Width = SysInfo1.WorkAreaWidth
        picDraw(0).Height = SysInfo1.WorkAreaHeight - (GetSystemMetrics(SM_CYCAPTION) _
                                                        * Screen.TwipsPerPixelY) _
                                                        - StatusBar1.Height
        .top = -24
        .left = -24
        .ScaleMode = 6
        .AutoRedraw = True
        .ScaleHeight = -Abs(.ScaleHeight)
'        .Visible = True
    End With

    With picDraw(1) ' �����\���p�s�N�`���{�b�N�X
        ' �w�i����
        .BackColor = RGB(250, 250, 250)
        .ForeColor = QBColor(0) ' ��
        .top = -24
        .left = -24
        .ScaleMode = 6
        .AutoRedraw = True
        .ScaleHeight = -Abs(.ScaleHeight)
        .Width = 50000
        .Height = 50000
        .Visible = False ' �N�����͔�\��
    End With

    ' �v���O���X�o�[�̃v���p�e�B�̐ݒ�
    With ProgressBar1
        .Width = 3000
        .Height = StatusBar1.Height - 12
        .top = StatusBar1.top + 12
        .left = StatusBar1.left + StatusBar1.Panels(1).Width
    End With

End Sub
