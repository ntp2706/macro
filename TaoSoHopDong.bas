Attribute VB_Name = "Module2"

Option Explicit

'========================================================
' T?O CHU?I:
' yyyymmdd-01/VNPT-TNH VNPTCGC/ HÐ – VNPT + TEXT
'
' Cách s? d?ng:
' 1. Ch?y macro TaoMaVNPT
' 2. Ch?n vùng/c?t k?t qu?
' 3. Nh?p text tùy ý
' 4. Macro l?y ngày ? c?t bên ph?i vùng k?t qu?
'========================================================

Public Sub TaoSoHopDong()

    Dim rngKQ As Range
    Dim o As Range
    Dim ngayGoc As Variant
    Dim ngayCode As String
    Dim txt As String
    Dim cotNgay As Long
    Dim lastRow As Long

    On Error GoTo Loi

    '====================================================
    ' 1. CH?N VÙNG K?T QU?
    '====================================================
    On Error Resume Next
    Set rngKQ = Application.InputBox( _
        Prompt:="Chon vung ghi ket qua", _
        Title:="Chon vung ghi ket qua", _
        Type:=8)
    On Error GoTo Loi

    If rngKQ Is Nothing Then Exit Sub

    'Ch? x? lý vùng d?u tiên
    Set rngKQ = rngKQ.Areas(1)

    '====================================================
    ' 2. NH?P TEXT
    '====================================================
    txt = InputBox( _
        "Nhap loai dich vu:" & vbCrLf & vbCrLf & _
        "Vi du: CA", _
        "Nhap Text")

    'N?u b?m Cancel
    If StrPtr(txt) = 0 Then Exit Sub

    '====================================================
    ' 3. C?T BÊN PH?I VÙNG K?T QU?
    '====================================================
    cotNgay = rngKQ.Column + rngKQ.Columns.Count

    Application.ScreenUpdating = False

    '====================================================
    ' 4. X? LÝ T?NG HÀNG
    '====================================================
    For Each o In rngKQ.Cells

        'L?y d? li?u cùng hàng ? c?t bên ph?i
        ngayGoc = o.Worksheet.Cells(o.Row, cotNgay).Value

        'N?u có d? li?u
        If Trim(CStr(ngayGoc)) <> "" Then

            '--------------------------------------------
            ' Tru?ng h?p Excel nh?n dúng là Date/Time
            '--------------------------------------------
            If IsDate(ngayGoc) Then

                ngayCode = Format(CDate(ngayGoc), "yyyymmdd")

                o.Value = ngayCode & _
                          "-01/VNPT-TNH VNPTCGC/ HÐ – VNPT " & txt

            Else

                '----------------------------------------
                ' Tru?ng h?p d? li?u dang là TEXT
                '----------------------------------------
                On Error Resume Next

                ngayCode = Format( _
                    CDate(Trim(CStr(ngayGoc))), _
                    "yyyymmdd")

                If Err.Number = 0 Then

                    o.Value = ngayCode & _
                              "-01/VNPT-TNH VNPTCGC/ HÐ – VNPT " & txt

                Else

                    o.Value = "Loi ngay"

                    Err.Clear

                End If

                On Error GoTo Loi

            End If

        Else

            'Không có ngày
            o.Value = ""

        End If

    Next o

    Application.ScreenUpdating = True

    MsgBox "Da tao ma thanh cong!", _
           vbInformation, _
           "Hoan tat"

    Exit Sub

Loi:

    Application.ScreenUpdating = True

    MsgBox "Co loi xay ra:" & vbCrLf & _
           Err.Description, _
           vbExclamation, _
           "Loi"

End Sub



