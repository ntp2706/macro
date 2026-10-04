Attribute VB_Name = "Module1"

Option Explicit

'========================================
' MAIN: TAO USERNAME
'========================================
Public Sub TaoUsername()

    Dim rngKQ As Range
    Dim rngTen As Range
    Dim rngNgaySinh As Range

    Dim i As Long
    Dim hoTen As String
    Dim ngaySinh As Variant
    Dim username As String

    On Error GoTo Loi

    '1. Chon cot ket qua
    Set rngKQ = ChonVung( _
        "1/3: Chon vung luu ket qua")

    If rngKQ Is Nothing Then Exit Sub

    '2. Chon cot ho va ten
    Set rngTen = ChonVung( _
        "2/3: Chon vung chua Ho va ten")

    If rngTen Is Nothing Then Exit Sub

    '3. Chon cot ngay sinh
    Set rngNgaySinh = ChonVung( _
        "3/3: Chon vung chua Ngay sinh")

    If rngNgaySinh Is Nothing Then Exit Sub

    'Kiem tra cac vung
    If rngKQ.Columns.Count <> 1 _
        Or rngTen.Columns.Count <> 1 _
        Or rngNgaySinh.Columns.Count <> 1 Then

        MsgBox "Vui lòng ch? ch?n m?t c?t cho m?i bu?c.", vbExclamation
        Exit Sub
    End If

    If rngKQ.Rows.Count <> rngTen.Rows.Count _
        Or rngKQ.Rows.Count <> rngNgaySinh.Rows.Count Then

        MsgBox "S? dòng c?a 3 vùng du?c ch?n không b?ng nhau.", vbExclamation
        Exit Sub
    End If

    'Dam bao cac vung cung bat dau tu dong 3
    If rngKQ.Row <> rngTen.Row _
        Or rngKQ.Row <> rngNgaySinh.Row Then

        MsgBox "Ba vùng ph?i b?t d?u t?i cùng m?t dòng.", vbExclamation
        Exit Sub
    End If

    'Ghi ket qua dang text
    rngKQ.NumberFormat = "@"

    Application.ScreenUpdating = False

    For i = 1 To rngKQ.Rows.Count

        hoTen = Trim(CStr(rngTen.Cells(i, 1).Value))
        ngaySinh = rngNgaySinh.Cells(i, 1).Value

        If hoTen <> "" And IsDate(ngaySinh) Then

            username = BoDau(hoTen) & _
                       Format(CDate(ngaySinh), "ddmmyyyy")

            rngKQ.Cells(i, 1).Value = username

        Else

            rngKQ.Cells(i, 1).ClearContents

        End If

    Next i

    Application.ScreenUpdating = True

    MsgBox "Ðã t?o username thành công!" & vbCrLf & _
           "S? dòng x? lý: " & rngKQ.Rows.Count, vbInformation

    Exit Sub

Loi:
    Application.ScreenUpdating = True
    MsgBox "Có l?i x?y ra: " & Err.Description, vbCritical

End Sub


'========================================
' HAM CHON VUNG DU LIEU
'========================================
Private Function ChonVung(ByVal noiDung As String) As Range

    On Error Resume Next

    Set ChonVung = Application.InputBox( _
        Prompt:=noiDung, _
        Title:="T?O USERNAME", _
        Type:=8)

    On Error GoTo 0

End Function


'========================================
' HAM BO DAU TIENG VIET
'========================================
Function BoDau(ByVal s As String) As String

    Dim CoDau As Variant
    Dim KhongDau As Variant
    Dim i As Long
    Dim ch As String
    Dim code As Long
    Dim ketQua As String

    '========================================
    ' CHUY?N V? CH? THU?NG
    '========================================
    s = LCase(s)


    '========================================
    ' CÁC KÝ T? TI?NG VI?T D?NG S?N
    '========================================

    CoDau = Array( _
        ChrW(&HE1), ChrW(&HE0), ChrW(&H1EA3), ChrW(&HE3), ChrW(&H1EA1), _
        ChrW(&H103), ChrW(&H1EAF), ChrW(&H1EB1), ChrW(&H1EB3), ChrW(&H1EB5), ChrW(&H1EB7), _
        ChrW(&HE2), ChrW(&H1EA5), ChrW(&H1EA7), ChrW(&H1EA9), ChrW(&H1EAB), ChrW(&H1EAD), _
        ChrW(&HE9), ChrW(&HE8), ChrW(&H1EBB), ChrW(&H1EBD), ChrW(&H1EB9), _
        ChrW(&HEA), ChrW(&H1EBF), ChrW(&H1EC1), ChrW(&H1EC3), ChrW(&H1EC5), ChrW(&H1EC7), _
        ChrW(&HED), ChrW(&HEC), ChrW(&H1EC9), ChrW(&H129), ChrW(&H1ECB), _
        ChrW(&HF3), ChrW(&HF2), ChrW(&H1ECF), ChrW(&HF5), ChrW(&H1ECD), _
        ChrW(&HF4), ChrW(&H1ED1), ChrW(&H1ED3), ChrW(&H1ED5), ChrW(&H1ED7), ChrW(&H1ED9), _
        ChrW(&H1A1), ChrW(&H1EDB), ChrW(&H1EDD), ChrW(&H1EDF), ChrW(&H1EE1), ChrW(&H1EE3), _
        ChrW(&HFA), ChrW(&HF9), ChrW(&H1EE7), ChrW(&H169), ChrW(&H1EE5), _
        ChrW(&H1B0), ChrW(&H1EE9), ChrW(&H1EEB), ChrW(&H1EED), ChrW(&H1EEF), ChrW(&H1EF1), _
        ChrW(&HFD), ChrW(&H1EF3), ChrW(&H1EF7), ChrW(&H1EF9), ChrW(&H1EF5), _
        ChrW(&H111))

    KhongDau = Array( _
        "a", "a", "a", "a", "a", _
        "a", "a", "a", "a", "a", "a", _
        "a", "a", "a", "a", "a", "a", _
        "e", "e", "e", "e", "e", _
        "e", "e", "e", "e", "e", "e", _
        "i", "i", "i", "i", "i", _
        "o", "o", "o", "o", "o", _
        "o", "o", "o", "o", "o", "o", _
        "o", "o", "o", "o", "o", "o", _
        "u", "u", "u", "u", "u", _
        "u", "u", "u", "u", "u", "u", _
        "y", "y", "y", "y", "y", _
        "d")

    'Thay th? các ký t? d?ng s?n
    For i = LBound(CoDau) To UBound(CoDau)
        s = Replace(s, CoDau(i), KhongDau(i))
    Next i


    '========================================
    ' X? LÝ UNICODE T? H?P
    '========================================
    '
    ' Ví d?:
    ' o + d?u n?ng
    ' a + d?u huy?n
    ' e + d?u s?c
    '
    ' Các d?u này t?n t?i nhu ký t? riêng
    ' nên ph?i lo?i b? riêng.
    '========================================

    ketQua = ""

    For i = 1 To Len(s)

        ch = Mid$(s, i, 1)
        code = AscW(ch)

        Select Case code

            ' D?u huy?n `
            Case &H300

            ' D?u s?c ´
            Case &H301

            ' D?u mu ^
            Case &H302

            ' D?u ngã ~
            Case &H303

            ' D?u trang
            Case &H306

            ' D?u h?i
            Case &H309

            ' D?u móc
            Case &H31B

            ' D?u n?ng
            Case &H323

            ' B? qua d?u

            Case Else
                ketQua = ketQua & ch

        End Select

    Next i


    '========================================
    ' LO?I B? KHO?NG TR?NG
    '========================================

    ketQua = Replace(ketQua, " ", "")
    ketQua = Replace(ketQua, vbTab, "")
    ketQua = Replace(ketQua, vbCr, "")
    ketQua = Replace(ketQua, vbLf, "")

    BoDau = ketQua

End Function


