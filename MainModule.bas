Attribute VB_Name = "MainModule"
Option Explicit

' 定数の設定
Public Const X As Integer = 0
Public Const Y As Integer = 1
Public Const R As Integer = 2
Public Const TH As Integer = 0
Public Const NT As Integer = 1
Public Const conTempFileName = "NCView._$$" ' テンポラリファイル名
Public Const intRow As Integer = 50 ' 列数
Public Const int1mm As Integer = 100 ' 1mmの値

Public Type NCInfo
    dblMin(1) As Double ' 最小値 X/Y
    dblMax(1) As Double ' 最大値 X/Y
    lngOffSet(1) As Long
End Type

'Public Type WBInfo
'    intSosu As Integer ' 層数
'    lngWBS(1) As Long ' WBS X/Y
'    lngStack(1) As Long ' Stack X/Y
'    strStart As String
'End Type

Public Type ToolInfo
    intTNo As Integer ' Tコード
    sngDrill As Single ' ドリル径
    lngColor As Long ' 色
End Type

'*********************************************************
' 用  途: NCViewのスタートアップ
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Sub Main()

    Dim strNC As String
    Dim strNCFileName As String
'    Dim frmNewMain As Form
'    Dim frmNewToolInfo As Form

    ' 2重起動をチェック
'    If App.PrevInstance Then
'        MsgBox "すでに起動されています！"
'        End
'    End If

    ' 初期化する
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
' 用  途: 変数を初期化する
' 引  数: 無し
' 戻り値: 無し
'*********************************************************

Public Sub sInitialize()

    Dim i As Integer

    frmMain.ScaleFactor = 1

    ' 値が入力されているか否かを判断する為に有り得ない値で初期化する


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
' 用  途: 環境変数TEMPの値を取得する
' 引  数: 無し
' 戻り値: 環境変数TEMPの値を返す
'*********************************************************

Public Function fTempPath() As String

    ' プログラム終了までTempPathの内容を保持
    Static TempPath As String

    ' 途中でディレクトリ-が変更されてもTempディレクトリ-を確保
    If Len(TempPath) = 0& Then
        TempPath = Environ("TEMP") ' ディレクトリ-を取得
        ' ルートディレクトリーかの判断
        If right$(TempPath, 1&) <> "\" Then
            TempPath = TempPath & "\"
        End If
    End If
    fTempPath = TempPath

End Function
