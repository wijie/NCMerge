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
' �p  �r: NC�t�@�C����ϐ��Ɉ�C�ǂ݂���
' ��  ��: strNCFileName: NC�t�@�C����
' �߂�l: NC�f�[�^���ۂ��ƕԂ�
'*********************************************************

Public Function fReadNC(ByVal strNCFileName As String) As String

    Dim intF0 As Integer
    Dim bytBuf() As Byte
    Dim strNC As String

    ' NC��ǂݍ���
    intF0 = FreeFile
    Open strNCFileName For Binary As #intF0
    ReDim bytBuf(LOF(intF0))
    Get #intF0, , bytBuf
    Close #intF0
    strNC = StrConv(bytBuf, vbUnicode)
    Erase bytBuf ' �z��̃��������J������

    ' ���s�R�[�h�𒲂ׂ�
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
' �p  �r: NC�f�[�^��S�ʒǂ��ɓW�J, ���a, �F����ǉ�����
' ��  ��: strNC: NC�f�[�^
'         udtNCInfo: NC�t�@�C����, �ő�l/�ŏ��l���i�[����\����
'         udtToolInfo(): �h�����a�����i�[����\���̂̔z��
'         udtWBInfo: �w��, WBS etc...���i�[����\����
'         objBar: �v���O���X�o�[�̃I�u�W�F�N�g�ϐ�
' �߂�l: ����I�������True
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
    Dim lngCount As Long '�v���O���X�o�[�̃J�E���^�p
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

    ' �폜���镶�������������
    strNC = Replace(strNC, " ", "")
    ' ���C��,�T�u�ɕ�������
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
    strNC = "" ' �ϐ��̃��������J������
    If UBound(strMainSub) = 1 Then
        strSubTmp = Split(strMainSub(0), "N", -1, vbTextCompare)
        For i = 1 To UBound(strSubTmp)
            intN = left(strSubTmp(i), 2) ' �T�u�������̔ԍ����擾
            varSub(intN) = Split(strSubTmp(i), mstrEnter, -1, vbBinaryCompare)
        Next
        strMain = Split(strMainSub(1), mstrEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), mstrEnter, -1, vbBinaryCompare)
    End If
    ' �z��̃��������J������
    Erase strMainSub
    Erase strSubTmp

    ' �o�͂���
    objBar.Visible = True
    lngCount = UBound(strMain)
    For i = 0 To lngCount
        If strMain(i) Like "X*Y*" = True Then
            strXY = Split(Mid(strMain(i), 2), "Y", -1, vbTextCompare)
            mlngABS(X) = mlngABS(X) + CLng(strXY(X)) ' ���݂�X���W
            mlngABS(Y) = mlngABS(Y) + CLng(strXY(Y)) ' ���݂�Y���W
            If mblnDrillHit = True Then
                Write #mintF0, mlngABS(X), mlngABS(Y), msngDrl, mlngColor, mintTool
                ' �ŏ��l/�ő�l���Z�b�g����
                Call sSetMinMax
            End If
        ElseIf strMain(i) Like "M89" = True Then ' �t�Z�b�g�`�F�b�N�p�R�[�h
            ' �������Ȃ�
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
            If intIndex > intRow Then ' ��v����c�[����������Ȃ�������
                MsgBox "�H������������ĉ�����"
                objBar.Visible = False
                fConvertNC = False ' ��v����c�[����������Ȃ�����False��Ԃ�
                Exit Function
            End If
        End If
        ' �v���O���X�o�[�̍X�V
        objBar.Value = Int(i / lngCount * 100)
    Next
    Erase strMain ' �z��̃��������J������

    ' NC�f�[�^�̍ő�/�ŏ��l���Z�b�g
    With udtNCInfo
        .dblMin(X) = mlngMin(X) / int1mm
        .dblMin(Y) = mlngMin(Y) / int1mm
        .dblMax(X) = mlngMax(X) / int1mm
        .dblMax(Y) = mlngMax(Y) / int1mm
    End With

    fConvertNC = True ' ����I������True��Ԃ�

End Function

Private Sub sSubMemo(ByRef strM As String, _
                     ByRef varSub() As Variant)

    Dim intSubNo As Integer
    Dim j As Long
    Dim strXY() As String

    intSubNo = CInt(Mid(strM, 2))
    ' �T�u�������[�͈̔͂�N44�`N97�ł���
    If intSubNo >= 44 And intSubNo <= 97 Then
        For j = 0 To UBound(varSub(intSubNo))
            If varSub(intSubNo)(j) Like "X*Y*" = True Then
                strXY = Split(Mid(varSub(intSubNo)(j), 2), "Y", -1, vbTextCompare)
                mlngABS(X) = mlngABS(X) + CLng(strXY(X)) ' ���݂�X���W
                mlngABS(Y) = mlngABS(Y) + CLng(strXY(Y)) ' ���݂�Y���W
                If mblnDrillHit = True Then
                    Write #mintF0, mlngABS(X), mlngABS(Y), msngDrl, mlngColor, mintTool
                    ' �ŏ��l/�ő�l���Z�b�g����
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
' �p  �r: NT�̃f�[�^�̍ŏ��l/�ő�l��ݒ肷��
' ��  ��: mlngMin(): ���݂܂ł̍ŏ��lX/Y�̔z��
'         mlngMax(): ���݂܂ł̍ő�lX/Y�̔z��
'         mlngABS(): ���݂̍��WX/Y�̔z��
' �߂�l: ����
'*********************************************************

Private Sub sSetMinMax()

    If mlngMax(X) < mlngABS(X) Then mlngMax(X) = mlngABS(X)
    If mlngMin(X) > mlngABS(X) Then mlngMin(X) = mlngABS(X)
    If mlngMax(Y) < mlngABS(Y) Then mlngMax(Y) = mlngABS(Y)
    If mlngMin(Y) > mlngABS(Y) Then mlngMin(Y) = mlngABS(Y)

End Sub

'*********************************************************
' �p  �r: NT�̃f�[�^����ړ��ʂ��擾����
' ��  ��: strNC: NC�f�[�^
' �߂�l: �ړ��ʂ�"X�`Y�`"�̌`���ŕԂ�
'*********************************************************

Public Function fGetNTIdou(ByVal strNC As String) As String

    Dim strMainSub() As String
    Dim strMain() As String
    Dim i As Long

    ' ���C��,�T�u�ɕ�������
    strMainSub = Split(strNC, "G25", -1, vbTextCompare)
'    strNC = ""
    If UBound(strMainSub) = 1 Then
        strMain = Split(strMainSub(1), mstrEnter, -1, vbBinaryCompare)
    Else
        strMain = Split(strMainSub(0), mstrEnter, -1, vbBinaryCompare)
    End If
    ' �z��̃��������J������
    Erase strMainSub

    ' NT�̈ړ��ʂ𒲂ׂ�
    For i = 0 To UBound(strMain)
        If strMain(i) Like "X*Y*" = True Then
            fGetNTIdou = strMain(i) ' �ړ��ʂ�Ԃ�
            Exit For
        ElseIf strMain(i) Like "T*" = True Then
            Exit For
        ElseIf strMain(i) Like "G81" = True Then
            Exit For
        End If
    Next
    Erase strMain ' �z��̃��������J������

End Function
