VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmNCInfo 
   BorderStyle     =   1  '固定(実線)
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
         Caption         =   "ﾌｧｲﾙ名"
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
         Caption         =   "ﾌｧｲﾙ名"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "ｸﾘｱ"
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
      Caption         =   "NTの移動量"
      Height          =   255
      Left            =   4920
      TabIndex        =   20
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "THの移動量"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblOutFile 
      Caption         =   "ﾌｧｲﾙ名(&N)"
      Height          =   255
      Left            =   4920
      TabIndex        =   22
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label lblMax 
      BorderStyle     =   1  '実線
      Caption         =   "-999.99"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   16
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "最大値"
      Height          =   255
      Left            =   4920
      TabIndex        =   15
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblMin 
      BorderStyle     =   1  '実線
      Caption         =   "-999.99"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblMax 
      BorderStyle     =   1  '実線
      Caption         =   "-999.99"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   17
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label lblMin 
      BorderStyle     =   1  '実線
      Caption         =   "-999.99"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "最小値"
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
Private mudtTHTool(1 To intRow) As ToolInfo ' THのツール情報
Private mudtNTTool(1 To intRow) As ToolInfo
Private mudtNCInfo(1) As NCInfo ' NC情報
Private mblnKeyFlag As Boolean

Private Sub cmbOffSet_LostFocus(Index As Integer)

    With cmbOffSet(Index)
        .Text = UCase(.Text)
    End With

End Sub

'*********************************************************
' 用  途: OKボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
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

    ' frmMainのロード
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
            .Col = 1 ' ドリル径
            If .Text <> "" Then
                .Col = 0 ' TNo
                mudtTHTool(i).intTNo = CInt(.Text)
                .Col = 1 ' ドリル径
                mudtTHTool(i).sngDrill = CSng(.Text) / 2
                mudtTHTool(i).lngColor = RGB(255, 0, 0) ' THは赤で描く
            End If
        End With
        With msgDrill(NT)
            .Row = i
            .Col = 1 ' ドリル径
            If .Text <> "" Then
                .Col = 0 ' TNo
                mudtNTTool(i).intTNo = CInt(.Text)
                .Col = 1 ' ドリル径
                mudtNTTool(i).sngDrill = CSng(.Text) / 2
                mudtNTTool(i).lngColor = RGB(0, 0, 255) ' NTは青で描く
            End If
        End With
    Next

    ' オフセットの設定
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

    ' テンポラリファイルのオープン
    strTempFile = fTempPath & conTempFileName
    intF0 = FreeFile
    Open strTempFile For Output As #intF0
    ' THの展開
    frmMain.StatusBar1.Panels(1).Text = "THを展開中..."
    DoEvents
    blnRet = fConvertNC(mstrNC(TH), _
                        mudtTHTool, _
                        mudtNCInfo(TH), _
                        intF0, _
                        frmMain.ProgressBar1)
    ' NTの展開
    If txtInFile(NT).Text <> "" Then
        frmMain.StatusBar1.Panels(1).Text = "NTを展開中..."
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

    ' 出力用NCファイルのオープン
    frmMain.StatusBar1.Panels(1).Text = "作画中..."
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
    MsgBox "ﾌｧｲﾙをｵｰﾌﾟﾝ出来ません。"

End Sub

'*********************************************************
' 用  途: クリアボタンのクリックイベント
' 引  数: 無し
' 戻り値: 無し
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
' 用  途: frmToolInfoのLoadイベント
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Private Sub Form_Load()

    ' 前回終了時のFormの位置を復元
    top = GetSetting("NCMRG", "Info", "Top", 0)
    left = GetSetting("NCMRG", "Info", "Left", 0)

    mblnKeyFlag = True

    ' コンボボックスの設定
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

    ' クリックイベントを発生させて初期化する
    cmdClear.Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' Formの位置をレジストリに保存
    SaveSetting "NCMRG", "Info", "Top", top
    SaveSetting "NCMRG", "Info", "Left", left

End Sub

'*********************************************************
' 用  途: フレキシブルグリッドのClickイベント
' 引  数: コントロール配列のIndexプロパティ
' 戻り値: 無し
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
' 用  途: フレキシブルグリッドのScrollイベント
' 引  数: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub msgDrill_Scroll(Index As Integer)

    ' コントロールが表示される前にイベントが発生するとエラーになるのでトラップする(-_-;
    On Error GoTo bye

    msgDrill(Index).SetFocus ' TextBoxにFocusがある時にScrollするとFocusがコマンドボタンに飛んでしまう為
    txtInput(Index).Visible = False

bye:

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのChangeイベント
' 引  数: Index: コントロール配列のIndexプロパティ
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_Change(Index As Integer)

    With msgDrill(Index)
        .CellAlignment = 1
        .Text = txtInput(Index).Text
    End With

End Sub

'*********************************************************
' 用  途: 入力用テキストボックスのKeyDownイベント
' 引  数: Index: コントロール配列のIndexプロパティ
'         KeyCode: キー コードを示す定数
'         Shift: イベント発生時のShift, Ctrl, Altキーの
'                状態を示す整数値
' 戻り値: 無し
'*********************************************************

Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

    ' 何も入力されていない時は次に進まない(戻るのは許可する)
    If txtInput(Index).Text = "" And KeyCode <> vbKeyUp Then Exit Sub

    ' Enter又は, Ctrl-M又は, 下矢印キー
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
' 用  途: 使用されているTコードを調べてグリッドコントロールに
'         セットする
' 引  数: strNC: NCデータ
' 戻り値: 無し
'*********************************************************

Private Sub sSetUsedTool(ByVal strNC As String, _
                         ByVal Index As Integer)

    Dim i As Integer
    Dim objReg As New RegExp
    Dim objMatches As Object
    Dim objMatch As Object

    objReg.Global = True
    objReg.IgnoreCase = False ' 大文字小文字を区別する
    objReg.Pattern = "T[0-9]+"
    Set objMatches = objReg.Execute(strNC)

    ' Tコードを工具情報にセットする
    i = 1
    With msgDrill(Index)
        For Each objMatch In objMatches
            .Row = i
            .Col = 0
            .Text = Mid(objMatch.Value, 2)
            .Col = 1 ' ドリル径
            .Text = "1.000" ' デフォルトのドリル径
            i = i + 1
        Next
        ' デフォルトの位置にセット
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
' 用  途: コントロールの初期化
' 引  数: TH/NTを示す値(TH - 1, NT - 2)
' 戻り値: 無し
'*********************************************************

Private Sub sInit(Index As Integer)

    Dim i As Integer

    With msgDrill(Index) ' グリッドの初期化
        .Cols = 2
        .Rows = intRow + 1 ' +1しているのは固定行がある為
        .FixedCols = 0 ' 固定列なし
        .FixedRows = 1 ' 固定行1
        .Width = 1440
        .Height = 3600
        .RowHeight(-1) = 264 ' 全列の高さ
        .RowHeight(0) = 240 ' 固定列の高さ
        .ColWidth(0) = 456 ' TNo.の桁幅
        .ColWidth(1) = 624 ' ドリル径の桁幅
        .FocusRect = flexFocusNone ' フォーカスを示す線を表示しない
        .HighLight = flexHighlightNever ' 強調表示しない
        .Row = 0 ' 個定列
        .Col = 0
        .Text = "TNo."
        .Col = 1
        .Text = "ﾄﾞﾘﾙ径"
        For i = 1 To intRow
            .Row = i
            .Col = 0 ' TNoの桁
            .CellAlignment = 1 ' 左側の中央
            .Text = ""
            .Col = 1 ' ドリル径の桁
            .CellAlignment = 1 ' 左側の中央
            .Text = ""
        Next
    End With
    msgDrill(Index).Row = 1

    With txtInput(Index) ' テキストボックスの初期化
        .ZOrder 0 ' 最前面へ移動
        .Move msgDrill(Index).CellLeft + msgDrill(Index).left, _
              msgDrill(Index).CellTop + msgDrill(Index).top
        .Width = msgDrill(Index).CellWidth
        .Height = msgDrill(Index).CellHeight
        .Appearance = 0 ' フラット
        .Alignment = vbLeftJustify ' 左寄せ
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
' 用  途: ファイルを開くダイアログを表示する
' 引  数: 無し
' 戻り値: 選択したファイル名
'*********************************************************

Public Function fGetInputFile() As String

'    Dim strPathName() As String

    ' CancelError プロパティを真 (True) に設定します。
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler

    ' ファイルの選択方法を設定します。
    CommonDialog1.Filter = "すべてのファイル (*.*)|*.*|テキスト" & "ファイル (*.txt)|*.txt|データファイル (*.dat)|*.dat"

    ' 既定の選択方法を指定します。
    CommonDialog1.FilterIndex = 1

    ' [ファイルを開く] ダイアログ ボックスを表示します。
    CommonDialog1.ShowOpen

    ' ファイルの有無をチェックする
    If Dir(CommonDialog1.FileName) = "" Then
        MsgBox "ファイルが見つかりません。"
        Exit Function
    End If

'    strPathName = Split(CommonDialog1.FileName, "\", -1)
'    ' ファイル名を削除する
'    strPathName(UBound(strPathName)) = ""
'    ' カレントディレクトリを移動する
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
