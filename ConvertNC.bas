Attribute VB_Name = "ConvertNC"
Option Explicit

Dim mlngABS(1) As Long
Dim mlngMax(1) As Long
Dim mlngMin(1) As Long
Dim mblnDrillHit As Boolean
Dim mintF0 As Integer
Dim msngDrl As Single
Dim mintTool As Integer
Dim mlngColor As Long
Dim mstrEnter As String

'*********************************************************
' 用  途: NCファイルを変数に一気読みする
' 引  数: strNCFileName: NCファイル名
' 戻り値: NCデータを丸ごと返す
'*********************************************************

Public Function fReadNC(ByVal strNCFileName As String) As String

    Dim intF0 As Integer
    Dim bytBuf() As Byte
    Dim strNC As String

    ' NCを読み込む
    intF0 = FreeFile
    Open strNCFileName For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)
    Erase bytBuf ' 配列のメモリを開放する

    ' 改行コードを調べる
    If InStr(strNC, vbCrLf) > 0 Then
        mstrEnter = vbCrLf
    ElseIf InStr(strNC, vbLf) > 0 Then
        mstrEnter = vbLf
    ElseIf InStr(strNC, vbCr) > 0 Then
        mstrEnter = vbCr
    End If

    fReadNC = strNC

End Function

'*********************************************************
' 用  途: NCデータを全面追いに展開, 穴径, 色情報を追加する
' 引  数: strNC: NCデータ
'         udtNCInfo: NCファイル名, 最大値/最小値を格納する構造体
'         udtToolInfo(): ドリル径情報を格納する構造体の配列
'         udtWBInfo: 層数, WBS etc...を格納する構造体
'         objBar: プログレスバーのオブジェクト変数
' 戻り値: 正常終了すればTrue
'*********************************************************

Public Function fConvertNC(ByVal strNC As String, _
                           ByRef udtToolInfo() As ToolInfo, _
                           ByRef udtNCInfo As NCInfo, _
                           ByRef intF0 As Integer, _
                           ByRef objBar As Object) As Boolean

    Dim strMainSub() As String
    Dim varSub(44 To 97) As Variant
    Dim strMain() As String
    Dim strSubTmp() As String
    Dim intN As Integer
    Dim i As Long
    Dim j As Long
    Dim intIndex As Integer
    Dim strXY() As String
    Dim strOutFile As String
    Dim intSubNo As Integer
    Dim lngCount As Long 'プログレスバーのカウンタ用
    Dim lngNTIdou(1) As Long

    mlngABS(X) = 0& - udtNCInfo.lngOffSet(X)
    mlngABS(Y) = 0& - udtNCInfo.lngOffSet(Y)
    mintF0 = intF0
    mlngMax(X) = -2147483647
    mlngMax(Y) = -2147483647
    mlngMin(X) = 2147483647
    mlngMin(Y) = 2147483647
    mblnDrillHit = False
    mintTool = -32767
    With objBar
        .Max = 100
        .Min = 0
        .Value = .Min
    End With

    ' 削除する文字列を処理する
    strNC = Replace(strNC, " ", "")
    ' メイン,サブに分割する
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' 変数のメモリを開放する
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        For i = 1 To UBound(strSubTmp)
            intN = left(strSubTmp(i), 2) ' サブメモリの番号を取得
            varSub(intN) = Split(strSubTmp(i), mstrEnter, -1, vbBinaryCompare)
        Next
        strMain = Split(strMainSub(1), mstrEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), mstrEnter, -1, vbBinaryCompare)
    End If
    ' 配列のメモリを開放する
    Erase strMainSub
    Erase strSubTmp

    ' 出力する
    objBar.Visible = True
    lngCount = UBound(strMain)
    For i = 0 To lngCount
        If strMain(i) Like "X*Y*" = True Then
            strXY = Split(Mid(strMain(i), 2), "Y", -1, vbTextCompare)
            mlngABS(X) = mlngABS(X) + CLng(strXY(X)) ' 現在のX座標
            mlngABS(Y) = mlngABS(Y) + CLng(strXY(Y)) ' 現在のY座標
            If mblnDrillHit = True Then
                Write #mintF0, mlngABS(X), mlngABS(Y), msngDrl, mlngColor, mintTool
                ' 最小値/最大値をセットする
                Call sSetMinMax
            End If
        ElseIf strMain(i) Like "M89" = True Then ' 逆セットチェック用コード
            ' 何もしない
        ElseIf strMain(i) Like "G81" = True Then
            mblnDrillHit = True
        ElseIf strMain(i) Like "G80" = True Then
            mblnDrillHit = False
        ElseIf strMain(i) Like "M##" = True Then
            Call sSubMemo(strMain(i), varSub)
        ElseIf strMain(i) Like "T*" = True Then
            mintTool = CInt(Mid(strMain(i), 2))
            For intIndex = 1 To intRow
                With udtToolInfo(intIndex)
                    If mintTool = CInt(.intTNo) Then
                        msngDrl = .sngDrill
                        mlngColor = .lngColor
                        Exit For
                    End If
                End With
            Next
            If intIndex > intRow Then ' 一致するツールが見つからなかった時
                MsgBox "工具情報を見直して下さい"
                objBar.Visible = False
                fConvertNC = False ' 一致するツールが見つからない時はFalseを返す
                Exit Function
            End If
        End If
        ' プログレスバーの更新
        objBar.Value = Int(i / lngCount * 100)
    Next
    Erase strMain ' 配列のメモリを開放する

    ' NCデータの最大/最小値をセット
    With udtNCInfo
        .dblMin(X) = mlngMin(X) / int1mm
        .dblMin(Y) = mlngMin(Y) / int1mm
        .dblMax(X) = mlngMax(X) / int1mm
        .dblMax(Y) = mlngMax(Y) / int1mm
    End With

    fConvertNC = True ' 正常終了時はTrueを返す

End Function

Private Sub sSubMemo(ByRef strM As String, _
                     ByRef varSub() As Variant)

    Dim intSubNo As Integer
    Dim j As Long
    Dim strXY() As String

    intSubNo = CInt(Mid(strM, 2))
    ' サブメモリーの範囲はN44～N97である
    If intSubNo >= 44 And intSubNo <= 97 Then
        For j = 0 To UBound(varSub(intSubNo))
            If varSub(intSubNo)(j) Like "X*Y*" = True Then
                strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, vbTextCompare)
                mlngABS(X) = mlngABS(X) + CLng(strXY(X)) ' 現在のX座標
                mlngABS(Y) = mlngABS(Y) + CLng(strXY(Y)) ' 現在のY座標
                If mblnDrillHit = True Then
                    Write #mintF0, mlngABS(X), mlngABS(Y), msngDrl, mlngColor, mintTool
                    ' 最小値/最大値をセットする
                    Call sSetMinMax
                End If
            ElseIf varSub(intSubNo)(j) Like "G81" = True Then
                mblnDrillHit = True
            ElseIf varSub(intSubNo)(j) Like "G80" = True Then
                mblnDrillHit = False
            End If
        Next
    End If

End Sub

'*********************************************************
' 用  途: NTのデータの最小値/最大値を設定する
' 引  数: mlngMin(): 現在までの最小値X/Yの配列
'         mlngMax(): 現在までの最大値X/Yの配列
'         mlngABS(): 現在の座標X/Yの配列
' 戻り値: 無し
'*********************************************************

Private Sub sSetMinMax()

    If mlngMax(X) < mlngABS(X) Then mlngMax(X) = mlngABS(X)
    If mlngMin(X) > mlngABS(X) Then mlngMin(X) = mlngABS(X)
    If mlngMax(Y) < mlngABS(Y) Then mlngMax(Y) = mlngABS(Y)
    If mlngMin(Y) > mlngABS(Y) Then mlngMin(Y) = mlngABS(Y)

End Sub

'*********************************************************
' 用  途: NTのデータから移動量を取得する
' 引  数: strNC: NCデータ
' 戻り値: 移動量を"X～Y～"の形式で返す
'*********************************************************

Public Function fGetNTIdou(ByVal strNC As String) As String

    Dim strMainSub() As String
    Dim strMain() As String
    Dim i As Long

    ' メイン,サブに分割する
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
'    strNC = ""
    If UBound(strMainSub) = 1 Then
        strMain = Split(strMainSub(1), mstrEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), mstrEnter, -1, vbBinaryCompare)
    End If
    ' 配列のメモリを開放する
    Erase strMainSub

    ' NTの移動量を調べる
    For i = 0 To UBound(strMain)
        If strMain(i) Like "X*Y*" = True Then
            fGetNTIdou = strMain(i) ' 移動量を返す
            Exit For
        ElseIf strMain(i) Like "T*" = True Then
            Exit For
        ElseIf strMain(i) Like "G81" = True Then
            Exit For
        End If
    Next
    Erase strMain ' 配列のメモリを開放する

End Function
