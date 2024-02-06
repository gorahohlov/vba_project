Attribute VB_Name = "Module1"

Sub highlight_codes_39()

'������ ������ ������������ ��� ��������� ����� � ������ �� ���
'(10 ������), ������� �������� � �������� �������, ������� ������
'��������������� ��������� ������� � �� �������� ������� ��������
'��� �39 �� 26.04.2012. "������, ���������� ��� ���������� ���������
'������� ��� ����������� �����������, ������������� ��� ���� �����,
'������ ���� ������ ����� ���� �������� ����, �����, ������, �������
'� �������� ����������� ������������ � (���) �������������
'����������������."

'������� ������.
'������������ �������� ������������ ������� �������� �����.
'� ���� ������� ����� ���� ����� ������: �����, �����, �������, ������
'������ � �.�. �������� ����� ������������ � ������ �� �������. ����
'������ �������� ��� ����� �� �������, ����� ������ ����������
'����������� ���������������. � ���������� ��������� ��� ��������������
'����� ����������� ������, ����������� �� � �.�. ��� ��������� ��������
'����������� ��������� ������������.

'����������� ������ �������, ������� ����� ����� � ����.
'������ ����� �������� ������ �� ���������� �������� ������ ������.
'�.�. ���� ������ �������� 4, 6, 9 � �.�. ������ - ������ ���� �����
'�� ���������� � �������� ����������� ���������������. �� ���� ���������
'� ���������� ���������� ����� ������ ��� ��� ����� ������������� ����!
'�.�. � ������� ���� ���������� � ���� �� ������ ������ ������.

'������� ����� � ����, ��� � ���������� ������� � �������� ������
'������� ����������� ��������� (���������) ������. �.�. ��� ��� ������
'�������� ��������, ��������, ����������, ������� ���� � ������
'���������������� ������� - ����� �������������� ��������� - ���������
'��������� ������ (�����). ���� ������ �� ���������!

'������ � ������� ���� "������������" � �����! ������� ������ ���� �����
'����� "������ ������".

'������ ������� ����������� ������ ��� ���� �� �������� ������� 8523.
'������ ����� �������, ��� �������������� ��������� ������� � ��
'�������� ������ "�������� ������� ������������" �� 8523!

'���� � ������ ���� �����-�� �������������� (�������, ����� �� ��� ����
'��� ������������� ������ - �����) ��� �����������! �.�. ���� ���
'�������������� � ������� ������� �39 ��� ������������ ���������� ������
'�� ����� ��� � ���� ������� - ��� ����������� � ������ ����� ��������
'��� ����� ��� �������� �������������� ��������� �������.
'������� ����� ������� �����, ����� ������� ���������� ������ ��������,
'������� �������������� ����� � ������.

'�������������� ����� �� ������� ������� �������� ��� �39
'(�������������� ����� �������) - ����� ������ �����, ������ ��� ������,
'���������� ������ ������� ������.

'�������������� ����� �� ������� ������������� ������������� �342
'(��������� ���� 30000) - ������ ��� ������, ���� ������ ����,
'���������� ������� ������.

'�������������� �����, ������� ����������� � ����� �������� - ����������
'���, ����� �����, ������� ����� ����� ������ ������.

'������ ��������� �� ������������ ������� (���������������) ������;
'����� ��������� �� "�����������" �� ������� ���������� ������.
'����� �� ������ � ������� ��������� �������������� ����������.
'������������ ����� �� ������ �������������� �������� �����, �� �
'������������ ����������� ��������� �����.


    Dim rng As Range, cell As Range
    Dim arr_one As Variant, _
        arr_except As Variant, _
        var_except As Variant, _
        arr_04 As Variant, _
        arr_06 As Variant, _
        arr_08 As Variant, _
        arr_09 As Variant, _
        arr_10 As Variant
    Dim cond_01 As Boolean, _
        cond_02 As Boolean, _
        cond_03 As Boolean, _
        cond_04 As Boolean, _
        cond_05 As Boolean, _
        cond_06 As Boolean, _
        cond_07 As Boolean, _
        bool_01 As Boolean
    Dim t As Single

    t = Timer

    arr_one = Array("2204", "2710", "3403", "4011", "7321", "8414", _
                    "8415", "8418", "8421", "8422", "8423", "8427", _
                    "8429", "8430", "8443", "8450", "8452", "8467", _
                    "8470", "8471", "8472", "8476", "8508", "8509", _
                    "8510", "8513", "8515", "8516", "8517", "8518", _
                    "8519", "8521", "8523", "8524", "8525", "8526", _
                    "8527", "8528", "8542", "8703", "8903", "9005", _
                    "9006", "9007", "9008", "9014", "9015", "9101", _
                    "9102", "9207", "9504")

    arr_except = Array("8476900000", "8515900000", "8516900000", _
                       "8518900003", "8518900005", "8518900008", _
                       "8526100001", "8526920001", "8527212001", _
                       "8527215201", "8527215901", "8527290001", _
                       "9005900000", "9008900000", "9014900000", _
                       "9015900000")

    var_except = "851680"

    arr_04 = Array("2204", "4011", "8423", "8427", "8429", "8470", "8471", _
                   "8472", "8476", "8515", "8516", "8518", "8519", "8521", _
                   "8523", "8525", "8526", "8527", "8528", "8542", "9005", _
                   "9008", "9014", "9015", "9101", "9102", "9207")

    arr_06 = Array("732111", "841459", "841510", "841810", "841821", _
                   "841830", "841840", "841850", "844331", "844332", _
                   "844339", "845011", "845210", "846711", "846721", _
                   "846722", "846729", "890332", "890333")

    arr_08 = Array("84158100", "85241100", "85249100")

    arr_09 = Array("841460000", "841520000", "850819000", _
                   "851761000", "851762000", "900659000")

    arr_10 = Array("2710197100", "2710197500", "2710198200", "2710198400", _
                   "2710198600", "2710198800", "2710199200", "2710199400", _
                   "2710199800", "2710209000", "3403191000", "3403199000", _
                   "3403990000", "7321810000", "8414510000", "8415820000", _
                   "8415830000", "8418290000", "8421120000", "8422110000", _
                   "8422190000", "8430200000", "8450120000", "8450190000", _
                   "8450200000", "8452210000", "8452290000", "8467190000", _
                   "8467810000", "8467890000", "8508110000", "8508600000", _
                   "8509400000", "8509800000", "8510100000", "8510200000", _
                   "8510300000", "8513100000", "8517110000", "8517130000", _
                   "8517140000", "8517180000", "8517691000", "8703101100", _
                   "8903210000", "8903220000", "8903230000", "8903310000", _
                   "8903939900", "8903990000", "9006300000", "9006400000", _
                   "9007100000", "9007200000", "9504301000", "9504500001", _
                   "9504500002")

'   With ActiveSheet.UsedRange
    For Each cell In Selection
        If Not cell.Rows.Hidden Then
            With cell
                a = .Value
                .NumberFormat = "General"
                .Value = a
                .NumberFormat = "@"
            End With
        End If
    Next
    
    Set rng = Selection

    For Each cell In rng

        If Not IsError(cell.Value) And _
           Not cell.Rows.Hidden Then
        If IsInArray(arr_one, Left(cell.Value, 4)) Then

            cond_01 = IsInArray(arr_04, Left(cell.Value, 4))
            cond_02 = IsInArray(arr_06, Left(cell.Value, 6))
            cond_03 = IsInArray(arr_08, Left(cell.Value, 8))
            cond_04 = IsInArray(arr_09, Left(cell.Value, 9))
            cond_05 = IsInArray(arr_10, Left(cell.Value, 10))
            cond_06 = IsInArray(arr_except, Left(cell.Value, 10))
            cond_07 = Left(cell.Value, 6) = var_except

            bool_01 = (cond_01 Or cond_02 Or cond_03 Or cond_04 Or cond_05) _
                      And _
                      (Not cond_06 And Not cond_07)

            If bool_01 Then

                With cell.Font
                    .Name = "Cambria"
                    .FontStyle = "�������"
                    .Size = 9
                    .ThemeColor = xlThemeColorAccent4
                End With
                With cell.Borders
                    .LineStyle = xlDouble
                    .ThemeColor = 8
                End With
                cell.Interior.ThemeColor = xlThemeColorLight1
                cell.HorizontalAlignment = xlCenter
                cell.VerticalAlignment = xlTop

            End If

        End If
        End If

    Next cell

    t = Timer - t
    MsgBox "������." & " ����� ����������: " & Round(t, 1) & " sec"

End Sub

Private Function IsInArray( _
                           arr As Variant, _
                           match_code As Variant _
                          ) As Boolean

    IsInArray = False

    For Each Item In arr
        If Item = match_code Then
            IsInArray = True
            Exit For
        End If
    Next

End Function

Sub highlight_codes_342()

    Dim arr_342_position As Variant, _
        var_342_04 As Variant, _
        arr_342_06 As Variant, _
        arr_342_09 As Variant, _
        arr_342_10 As Variant
    Dim cond_08 As Boolean, _
        cond_09 As Boolean, _
        cond_10 As Boolean, _
        cond_11 As Boolean, _
        bool_02 As Boolean
    Dim rng As Range, cell As Range
    Dim t As Single

    t = Timer

    arr_342_position = Array("8443", "8471", "8473", "8517", "8518", "8519", _
                             "8521", "8523", "8525", "8526", "8527", "8528", _
                             "8531", "8536", "8544", "9006", "9007", "9008", _
                             "9010", "9012", "9014", "9015", "9016", "9017", _
                             "9024", "9025", "9026", "9027", "9028", "9029", _
                             "9030", "9031", "9032", "9101", "9102", "9104", _
                             "9106", "9504")

    var_342_04 = "8471"

    arr_342_06 = Array("844331", "844332", "847330", "847350", "851769", _
                       "851771", "851810", "851840", "851920", "851981", _
                       "851989", "852110", "852351", "852691", "852712", _
                       "852713", "852721", "852791", "852792", "852842", _
                       "852849", "852859", "852869", "852871", "852872", _
                       "853110", "853620", "900653", "901210", "901420", _
                       "901510", "901520", "901540", "901580", "901600", _
                       "901710", "901720", "902410", "902480", "902511", _
                       "902519", "902580", "902610", "902620", "902680", _
                       "902710", "902790", "902830", "902920", "903020", _
                       "903033", "903089", "903149", "903180", "903210")

    arr_342_09 = Array("851761000", "851762000", "851822000", "852190000", _
                       "852560000", "852610000", "852692000", "852729000", _
                       "900659000", "900669000", "902910000", "903039000", _
                       "910400000", "950450000")

    arr_342_10 = Array( _
                      "8517130000", "8517140000", "8518500000", "8519300000", _
                      "8527190000", "8527990000", "8528730000", "8544700000", _
                      "9006300000", "9006400000", "9007100000", "9007200000", _
                      "9008500000", "9010500000", "9010600000", "9014100000", _
                      "9014800000", "9027200000", "9027300000", "9027500000", _
                      "9028100000", "9028200000", "9030100000", "9030310000", _
                      "9030400000", "9031200000", "9031410000", "9032200000", _
                      "9032810000", "9032890000", "9101910000", "9102120000", _
                      "9102190000", "9102910000", "9106100000", "9106900000" _
                      )

    For Each cell In Selection
        If Not cell.Rows.Hidden Then
            With cell
                a = .Value
                .NumberFormat = "General"
                .Value = a
                .NumberFormat = "@"
            End With
        End If
    Next

    Set rng = Selection

    For Each cell In rng

        If Not IsError(cell.Value) And _
           Not cell.Rows.Hidden Then
        If IsInArray(arr_342_position, Left(cell.Value, 4)) Then

            cond_08 = Left(cell.Value, 4) = var_342_04
            cond_09 = IsInArray(arr_342_06, Left(cell.Value, 6))
            cond_10 = IsInArray(arr_342_09, Left(cell.Value, 9))
            cond_11 = IsInArray(arr_342_10, Left(cell.Value, 10))

            bool_02 = cond_08 Or cond_09 Or cond_10 Or cond_11

            If bool_02 Then

                With cell.Font
                    .Name = "Cambria"
                    .FontStyle = "�������"
                    .Size = 9
                    .ThemeColor = xlThemeColorDark1
                End With
                With cell.Borders
                    .LineStyle = xlDouble
                    .ThemeColor = 1
                End With
                cell.Interior.Color = 10119167
                cell.HorizontalAlignment = xlCenter
                cell.VerticalAlignment = xlTop

            End If

        End If
        End If

    Next cell

    t = Timer - t
    MsgBox "������." & " ����� ����������: " & Round(t, 1) & " sec"

End Sub

Sub highlight_cells()
Attribute highlight_cells.VB_ProcData.VB_Invoke_Func = "q\n14"

    Dim rng As Range, cell As Range
    Dim arr_one As Variant, _
        arr_except As Variant, _
        var_except As Variant, _
        arr_04 As Variant, _
        arr_06 As Variant, _
        arr_08 As Variant, _
        arr_09 As Variant, _
        arr_10 As Variant, _
        arr_342_position As Variant, _
        var_342_04 As Variant, _
        arr_342_06 As Variant, _
        arr_342_09 As Variant, _
        arr_342_10 As Variant
    Dim cond_01 As Boolean, cond_02 As Boolean, cond_03 As Boolean, _
        cond_04 As Boolean, cond_05 As Boolean, cond_06 As Boolean, _
        cond_07 As Boolean, cond_08 As Boolean, cond_09 As Boolean, _
        cond_10 As Boolean, cond_11 As Boolean, _
        bool_01 As Boolean, bool_02 As Boolean
    Dim t As Single, FLG As Single


    t = Timer

    arr_one = Array("2204", "2710", "3403", "4011", "7321", "8414", _
                    "8415", "8418", "8421", "8422", "8423", "8427", _
                    "8429", "8430", "8443", "8450", "8452", "8467", _
                    "8470", "8471", "8472", "8476", "8508", "8509", _
                    "8510", "8513", "8515", "8516", "8517", "8518", _
                    "8519", "8521", "8523", "8524", "8525", "8526", _
                    "8527", "8528", "8542", "8703", "8903", "9005", _
                    "9006", "9007", "9008", "9014", "9015", "9101", _
                    "9102", "9207", "9504")

    arr_except = Array("8476900000", "8515900000", "8516900000", _
                       "8518900003", "8518900005", "8518900008", _
                       "8526100001", "8526920001", "8527212001", _
                       "8527215201", "8527215901", "8527290001", _
                       "9005900000", "9008900000", "9014900000", _
                       "9015900000")

    var_except = "851680"

    arr_04 = Array("2204", "4011", "8423", "8427", "8429", "8470", "8471", _
                   "8472", "8476", "8515", "8516", "8518", "8519", "8521", _
                   "8523", "8525", "8526", "8527", "8528", "8542", "9005", _
                   "9008", "9014", "9015", "9101", "9102", "9207")

    arr_06 = Array("732111", "841459", "841510", "841810", "841821", _
                   "841830", "841840", "841850", "844331", "844332", _
                   "844339", "845011", "845210", "846711", "846721", _
                   "846722", "846729", "890332", "890333")

    arr_08 = Array("84158100", "85241100", "85249100")

    arr_09 = Array("841460000", "841520000", "850819000", _
                   "851761000", "851762000", "900659000")

    arr_10 = Array("2710197100", "2710197500", "2710198200", "2710198400", _
                   "2710198600", "2710198800", "2710199200", "2710199400", _
                   "2710199800", "2710209000", "3403191000", "3403199000", _
                   "3403990000", "7321810000", "8414510000", "8415820000", _
                   "8415830000", "8418290000", "8421120000", "8422110000", _
                   "8422190000", "8430200000", "8450120000", "8450190000", _
                   "8450200000", "8452210000", "8452290000", "8467190000", _
                   "8467810000", "8467890000", "8508110000", "8508600000", _
                   "8509400000", "8509800000", "8510100000", "8510200000", _
                   "8510300000", "8513100000", "8517110000", "8517130000", _
                   "8517140000", "8517180000", "8517691000", "8703101100", _
                   "8903210000", "8903220000", "8903230000", "8903310000", _
                   "8903939900", "8903990000", "9006300000", "9006400000", _
                   "9007100000", "9007200000", "9504301000", "9504500001", _
                   "9504500002")

    arr_342_position = Array("8443", "8471", "8473", "8517", "8518", "8519", _
                             "8521", "8523", "8525", "8526", "8527", "8528", _
                             "8531", "8536", "8544", "9006", "9007", "9008", _
                             "9010", "9012", "9014", "9015", "9016", "9017", _
                             "9024", "9025", "9026", "9027", "9028", "9029", _
                             "9030", "9031", "9032", "9101", "9102", "9104", _
                             "9106", "9504")

    var_342_04 = "8471"

    arr_342_06 = Array("844331", "844332", "847330", "847350", "851769", _
                       "851771", "851810", "851840", "851920", "851981", _
                       "851989", "852110", "852351", "852691", "852712", _
                       "852713", "852721", "852791", "852792", "852842", _
                       "852849", "852859", "852869", "852871", "852872", _
                       "853110", "853620", "900653", "901210", "901420", _
                       "901510", "901520", "901540", "901580", "901600", _
                       "901710", "901720", "902410", "902480", "902511", _
                       "902519", "902580", "902610", "902620", "902680", _
                       "902710", "902790", "902830", "902920", "903020", _
                       "903033", "903089", "903149", "903180", "903210")

    arr_342_09 = Array("851761000", "851762000", "851822000", "852190000", _
                       "852560000", "852610000", "852692000", "852729000", _
                       "900659000", "900669000", "902910000", "903039000", _
                       "910400000", "950450000")

    arr_342_10 = Array( _
                      "8517130000", "8517140000", "8518500000", "8519300000", _
                      "8527190000", "8527990000", "8528730000", "8544700000", _
                      "9006300000", "9006400000", "9007100000", "9007200000", _
                      "9008500000", "9010500000", "9010600000", "9014100000", _
                      "9014800000", "9027200000", "9027300000", "9027500000", _
                      "9028100000", "9028200000", "9030100000", "9030310000", _
                      "9030400000", "9031200000", "9031410000", "9032200000", _
                      "9032810000", "9032890000", "9101910000", "9102120000", _
                      "9102190000", "9102910000", "9106100000", "9106900000" _
                      )

    For Each cell In Selection
        If Not cell.Rows.Hidden Then
            With cell
                a = .Value
                .NumberFormat = "General"
                .Value = a
                .NumberFormat = "@"
            End With
        End If
    Next

    Set rng = Selection

    For Each cell In rng

        If Not IsError(cell.Value) And _
           Not cell.Rows.Hidden Then
        If IsInArray(arr_one, Left(cell.Value, 4)) Or _
           IsInArray(arr_342_position, Left(cell.Value, 4)) Then

            cond_01 = IsInArray(arr_04, Left(cell.Value, 4))
            cond_02 = IsInArray(arr_06, Left(cell.Value, 6))
            cond_03 = IsInArray(arr_08, Left(cell.Value, 8))
            cond_04 = IsInArray(arr_09, Left(cell.Value, 9))
            cond_05 = IsInArray(arr_10, Left(cell.Value, 10))
            cond_06 = IsInArray(arr_except, Left(cell.Value, 10))
            cond_07 = Left(cell.Value, 6) = var_except

            bool_01 = (cond_01 Or cond_02 Or cond_03 Or cond_04 Or cond_05) _
                      And _
                      (Not cond_06 And Not cond_07)

            cond_08 = Left(cell.Value, 4) = var_342_04
            cond_09 = IsInArray(arr_342_06, Left(cell.Value, 6))
            cond_10 = IsInArray(arr_342_09, Left(cell.Value, 9))
            cond_11 = IsInArray(arr_342_10, Left(cell.Value, 10))

            bool_02 = cond_08 Or cond_09 Or cond_10 Or cond_11

            FLG = IIf(bool_01 = True And bool_02 = False, 1, _
                    IIf(bool_01 = False And bool_02 = True, 2, _
                        IIf(bool_01 = True And bool_02 = True, 3, 0)))

            Select Case FLG
                Case 1

                    With cell.Font
                        .Name = "Cambria"
                        .FontStyle = "�������"
                        .Size = 9
                        .Bold = True
                        .Color = vbWhite
                    End With
                    With cell.Borders
                        .LineStyle = xlDouble
                        .Color = vbWhite
                    End With
                    cell.Interior.ThemeColor = xlThemeColorLight1
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlTop

                Case 2

                    With cell.Font
                        .Name = "Cambria"
                        .FontStyle = "�������"
                        .Size = 9
                        .Bold = True
                        .ColorIndex = xlAutomatic
'                        .ThemeColor = xlThemeColorDark1
                    End With
                    With cell.Borders
                        .LineStyle = xlDouble
                        .Color = vbBlack
'                        .Weight = xlThick
                    End With
                    cell.Interior.Color = vbYellow
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlTop

                Case 3

                    With cell.Font
                        .Name = "Cambria"
                        .FontStyle = "�������"
                        .Size = 9
                        .Bold = True
                        .Color = vbWhite
                    End With
                    With cell.Borders
                        .LineStyle = xlDouble
                        .ThemeColor = 1
                    End With
                    cell.Interior.Color = RGB(147, 75, 201)
                    cell.HorizontalAlignment = xlCenter
                    cell.VerticalAlignment = xlTop

            End Select

        End If
        End If

    Next cell

    t = Timer - t
    MsgBox "������." & " ����� ����������: " & Round(t, 1) & " sec"

End Sub

Function VLookUp2( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������,  � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'match_num: ����� ���������� �������� (���� ���������� �������������);

    Dim FLG As Boolean
    Dim i As Integer
    Dim iCount As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If table_rng.Cells(i, search_col_num) = search_value Then
            iCount = iCount + 1
        End If

        If iCount = match_num Then
            VLookUp2 = table_rng.Cells(i, result_col_num)
            FLG = True
            Exit For
        End If

    Next i

    If FLG = False Then
        VLookUp2 = CVErr(xlErrNA)
    End If

End Function

Function VLookUp3( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������, � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'match_num: ����� ���������� �������� (���� ���������� �������������);

    Dim FLG As Boolean
    Dim i As Integer
    Dim iCount As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If search_value Like table_rng.Cells(i, search_col_num) Then
            iCount = iCount + 1
        End If

        If iCount = match_num Then
            VLookUp3 = table_rng.Cells(i, result_col_num)
            FLG = True
            Exit For
        End If

    Next i

    If FLG = False Then
        VLookUp3 = CVErr(xlErrNA)
    End If

End Function

Function VLookUp4( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  Optional symbols_num = 0 _
                 )

'search_value: ������� ��������;
'table_rng: �������, �������� �����, � ������� ������ ���������� �
'                   ����������;
'search_col_num: ����� �������, � ��������� ����� [table_rng],
'                   � ������� ������ ����������;
'result_col_num: ����� �������, � ��������� ����� [table_rng],
'                   �� ������� ����������� ������� ������;
'symbols_num: ���������� ������ ����� �������� ��������� ��������
'                   (��������), �� ������� ����� �������� ����������.

    Dim FLG As Boolean
    Dim i As Integer

    FLG = False

    For i = 1 To table_rng.Rows.Count

        If symbols_num = 0 Then
            If table_rng.Cells(i, search_col_num) = search_value Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                FLG = True
                Exit For
            End If
        Else
            If Left(table_rng.Cells(i, search_col_num), symbols_num) = _
                                    Left(search_value, symbols_num) Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                FLG = True
                Exit For
            End If
        End If

    Next i

    If FLG = False Then
        VLookUp4 = CVErr(xlErrNA)
    End If

End Function

Sub partial_match()

'������ ������� ���������. ���������� ������ � ����� PERSONAL.XLSB

'��������� ������ � ����������� �������� �� ���������� ����������
'��������. �� ���� ��� ��������� ���������� ���������������� �������
'��� (VLookUp). �.�. �������� ��� ��� ������ ��������� �������� ��������
'(�������� �� ������� ������� �����, ������������� ����������) ��
'������-�� ���������� ������.

'� ������������ ������������� ���������� �������� �������� ��������
'(�������� 9) � ��������������� ������ ���������� �� ����������
'�������� 9, 8, 7, 6 - �� ������� �� ����� 4 ��������.


    Dim rng  As Range, _
        cell As Range, _
        article_col_rng As Range, _
        vlookup_table_rng As Range
    Dim upper_interval As Integer, _
        cntr1 As Long, _
        cntr2 As Integer
'    Dim lower_interval As Integer
    Dim article_col_num As Integer, _
        processing_row_num As Integer
    Dim hscode As Variant, _
        description_31 As Variant, _
        description_art As Variant
    Dim t As Single
 
    t = Timer

'    Set rng = Selection
'    lower_interval = 4
    upper_interval = InputBox( _
                              "������� ������� ������� ���������� ��������:", _
                              "���� ���������� ��������", _
                               9 _
                             )
    
    
'   ������ (����� �������) ������ ����������� ��������� �����
'   �� ��������� ��������� �� ��������� ������� � ���������;
'   �������� ����� ���� �����; ������� ����� ����� ������� ������
'   (����� �������) ��������� �� ��������� ������ (��������).
    Set article_col_rng = Application.InputBox( _
        "������� ������� � ���������, �� �������� ����� ������������ �����", _
        "���� ������� � �������� ������� (����������)", _
         Type:=8 _
        )
                                                        
    Set vlookup_table_rng = Application.InputBox( _
        "������� �������� ������ ��������� ���������� ������", _
        "���� ��������� (������ ����� �������� ������)", _
         Type:=8 _
        )

    article_col_num = article_col_rng.Column
    processing_row_num = Selection.Row
    
                
'    For Each cell In Selection
'
'        hscode = VLookUp4( _
                           Cells(processing_row_num, _
                                 article_col_num), _
                           vlookup_table_rng, _
                           9, _
                           1, _
                           0 _
                          )
'
'        If IsError(hscode) Then
'
'            For cntr = upper_interval To upper_interval - 3 Step -1
'
'                hscode = VLookUp4( _
                            Cells(processing_row_num, _
                                  article_col_num), _
                            vlookup_table_rng, _
                            9, _
                            1, _
                            cntr _
                           )
'
'                If Not IsError(hscode) Then
'                    cell = hscode
'                    Select Case cntr
'                        Case upper_interval - 0
'                            cell.Interior.Color = 13819376
'                        Case upper_interval - 1
'                            cell.Interior.Color = 11321572
'                        Case upper_interval - 2
'                            cell.Interior.Color = 8823768
'                        Case upper_interval - 3
'                            cell.Interior.Color = 4025277
'                    End Select
'                    Exit For
'                End If
'
'            Next
'
'        End If
'
'        cell = hscode
'        processing_row_num = processing_row_num + 1
'
'    Next cell
    
    For cntr1 = 1 To Selection.Count Step Selection.Columns.Count

        hscode = VLookUp4( _
                          Cells(processing_row_num, _
                                article_col_num), _
                          vlookup_table_rng, _
                          9, _
                          1, _
                          0 _
                         )
        If Not IsError(hscode) Then ' hscode <> "#�/�"
            description_31 = VLookUp4( _
                                      Cells(processing_row_num, _
                                            article_col_num), _
                                      vlookup_table_rng, _
                                      9, _
                                      5, _
                                      0 _
                                     )
            description_art = VLookUp4( _
                                       Cells(processing_row_num, _
                                             article_col_num), _
                                       vlookup_table_rng, _
                                       9, _
                                       6, _
                                       0 _
                                      )
        Else
            description_31 = CVErr(xlErrNA) ' "#�/�"
            description_art = CVErr(xlErrNA) ' "#�/�"
        End If
        
        If IsError(hscode) Then ' hscode = "#�/�"

            For cntr2 = upper_interval To upper_interval - 3 Step -1

                hscode = VLookUp4( _
                                  Cells(processing_row_num, article_col_num), _
                                  vlookup_table_rng, _
                                  9, _
                                  1, _
                                  cntr2 _
                                 )
                
                If Not IsError(hscode) Then ' hscode <> "#�/�"
                    description_31 = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                9, _
                                5, _
                                cntr2 _
                               )
                    description_art = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                9, _
                                6, _
                                cntr2 _
                               )
                Else
                    description_31 = CVErr(xlErrNA)  ' "#�/�"
                    description_art = CVErr(xlErrNA) ' "#�/�"
                End If

                If Not IsError(hscode) Then ' hscode <> "#�/�"
                    
                    Select Case cntr2
                        Case upper_interval - 0
                            Selection.Cells(cntr1).Interior.Color = 13819376
                            
'                        ���� ������������� ������� ���� (�� 6 �����)
'                        ���� �� ������� �������� � ��������� ���������!
                            If Selection.Columns.Count = 2 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    13819376
                            ElseIf Selection.Columns.Count = 3 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    13819376
                                Selection.Cells(cntr1 + 2).Interior.Color = _
                                                                    13819376
                            End If
                            
                        Case upper_interval - 1
                            Selection.Cells(cntr1).Interior.Color = 11321572
                            
                            If Selection.Columns.Count = 2 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    11321572
                            ElseIf Selection.Columns.Count = 3 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    11321572
                                Selection.Cells(cntr1 + 2).Interior.Color = _
                                                                    11321572
                            End If
                            
                        Case upper_interval - 2
                            Selection.Cells(cntr1).Interior.Color = 8823768
                            
                            If Selection.Columns.Count = 2 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    8823768
                            ElseIf Selection.Columns.Count = 3 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    8823768
                                Selection.Cells(cntr1 + 2).Interior.Color = _
                                                                    8823768
                            End If
                            
                        Case upper_interval - 3
                            Selection.Cells(cntr1).Interior.Color = 4025277
                            
                            If Selection.Columns.Count = 2 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    4025277
                            ElseIf Selection.Columns.Count = 3 Then
                                Selection.Cells(cntr1 + 1).Interior.Color = _
                                                                    4025277
                                Selection.Cells(cntr1 + 2).Interior.Color = _
                                                                    4025277
                            End If
                            
                    End Select
                    
                    Exit For
                End If

            Next cntr2

        End If

        Selection.Cells(cntr1) = hscode
        If Selection.Columns.Count = 2 Then
            Selection.Cells(cntr1 + 1) = description_31
        ElseIf Selection.Columns.Count = 3 Then
            Selection.Cells(cntr1 + 1) = description_31
            Selection.Cells(cntr1 + 2) = description_art
        End If
        
        processing_row_num = processing_row_num + 1

    Next cntr1
    
    t = Timer - t
    MsgBox "������." & " ����� ����������: " & Round(t, 1) & " sec"


'������� � ��������� ���������

'������ �������: ������������� ������ Application.Inputbox
'��� ����� �������
'        ActiveCell = Application.InputBox(prompt:= _
'            "������� �������: VLookUp2(
'               search_value;table_rng;SearchColNum;ResultColNum;match_num
'                                       )", _
'            Title:="���� �������", Default:="=VLookUp2()", Type:=0)
'����� ����� ������� ��� ����������, ����� ����� � ������ � ������� �
'������� Shift+F3 - ���� "��������� �������" ��� ���������� ����������;

'������ �������: ������������ ����������� ���� ������� �������:
'            Application.CommandBars.ExecuteMso "FunctionWizard"

'������ �������: ��������� ��� ��������� ��� ������� VLookUp
'� ������� ���������� ������� Application.InputBox


'ActiveCell.FormulaLocal = "=VLookUp3()" ' ����� Shift+F3 - �����
'                           ����������� ���� "��������� �������"
'���
'ActiveCell.FormulaLocal = "=VLookUp3(" ' ����� Ctrl+A ��� ������
'                           ������� "��������� �������"
'���
'ActiveCell = Application.InputBox(prompt:="������� �������:
'VLookUp2(search_value;<...>)", Title:="���� �������",
'Default:="=VLookUp2()", Type:=0)
'����� Enter � Shift+F3 ��� ������ ������� "��������� �������"

'Set Rng1 = Application.InputBox("������� ������� � ��������� ��
'�������� ����� ������������ �����", "���� �������", Type:=8)
'Print Rng1.Column

'Set Rng2 = Application.InputBox("������� �������� ������ ���������
'���������� ������", "���� ���������", Type:=8)
'Print Rng2.Column

End Sub

Function custom_toll( _
                     custom_sum As Variant, _
                     Optional currency_rate As Single = 1, _
                     Optional msg_flag As Boolean = False _
                    ) As Variant


'��� ���������� ������ ������� (� ����� my_funcs.xlam).

'���������� �������.
'��� ������� ������������ �����  ���������� ������ � ����������� ��
'����� ���������� ��������� ������������ �� ������ (������ ��������),
'������ �������� - �������������� - ��� ����, �� ������� ����������!
'������ �������� �������, ����� �������� ����� ���������� ���������
'(� ������) ��� ������� ������. ���� ��������������� ��� ������ ��������
'(������������) ��� ���������� ��������� � ������, ����� ������ ��������
'(����) ����� �� ��������� ��� ��������� 1 (�������) - ����� ��������
'����� �� ���������, ���� ������ �������� ��������.

'��� ������ �������.
'���� ������� ��������� �� ����� � ������ (��������, ���� ���
'����� - �������) - ����� ������ ���������� ������� ����, �� �������
'����� ����������! ������ ��������, ����� �������� ����� � ������.
'����� ���������� ���������� ��������� � ������ "�����������" ��
'����������� "IF ElseIF Else EndIF" - ������� ������ ����� ����������
'������ � ������ � ����������� �� ���������� ���������. ���� ������
'�������� ������ � ������� �� 1 (�������) - ���� �� ����������� "IF
' ElseIF Else EndIF" ������� �� ���� ���� � ��������� �����������
'� ��� �� ������, � ������� ��������������� ������������ ����������
'��������� (������) ��������.

'�������� msg_flag (������ ��������) �������� ����� �� ����� ����������
'��������� � ����������� �� ������, ���� ��� �������� (���� ������
'������� ����������� ������ ������� ��������� �� ��������).
'������ ������ ����� � ��� �������� ����� ��� ���������; �� �����
'����� � ����, ��� ���� �� ����� ����� ����� � ���� �������� � ��
'�����-�� ������� ��������� ������ � ������ (������� ������ � ���������
'������� � �.�.) �� �������� ����� ����������� ����������� ����
'��������� ��� ������! ����������� ����� ��� �����.
'�� ��������� ������� ������������ ��� ��������� �� ������� ����, ����
'��� ��������� (�.�. ���� �������� "msg_string" �� ��������� ��������).
'������� �� ��������� ������ ���������� ��������� ���������
'��������������� ������ � ������ ����������� ������.

'����������� ������.
'������� ������������ ��������, ����� ������, �� ������� ���������
'�������, �������� �������� ������ ��������������� � ������� ����,
'��� ����������� �������� (True, False) ��� �������� �������� �
'������������� ���������, � ����� ���� ������ �������� ��������� ������.
'� ����� �������, ����� ������, �� ������� ��������� �������,
'������������� ��� ����, ������ ��� ���� ��� �������� �������������
'��������, ��������� �������� - ������ �������� ������: "#����!",
'"#�����!", "#���?", "#���/0!".
'����� � ������ ����� � ����� ������� ����������
'���������-�������������� "� �������� ������ ���������� � �������".

'���������� (��������).
'����� ���������� ������ ��������� � ���� ������� �� ������������
'��������� - �.�. � ����������� �� ����� ���������� ��������� - ���
'����� �����, �� ������� ����� ����������� 30000 ���. ������� �������
'� ������� ������ �� �������������� ����� ��������, ����� �����
'���������� ������ ������� �� ���� �����.


    Dim custom_sum_ru As Variant
    Dim bool_1 As Boolean, _
        bool_2 As Boolean
    Dim msg_string As String

'   On Error GoTo handler:

'   ���� ��������� ��� ������� ���� ����, ��������� ��� ������ ��
'   ����������� ��� ����������� ������ ����� � ���� �� ������
'   ���������� ���������� � ���������� ����������� �������� ������
'   "#����!" - �� ����� ������. ������� ���������, ������� ����������
'   ��������� ������� �� �����.
    On Error Resume Next

    If TypeName(custom_sum) = "Range" Then
        bool_1 = TypeName(custom_sum.Value) = "Boolean"
        bool_2 = TypeName(custom_sum.Value) = "Error"
    Else
        bool_1 = TypeName(custom_sum) = "Boolean"
        bool_2 = TypeName(custom_sum) = "Error"
    End If

    custom_sum_ru = custom_sum * currency_rate
    
    If IsDate(custom_sum) Then 'TypeName(custom_sum.Value) = "Date" Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ����!"
        custom_toll = CVErr(xlErrValue) ' "#DATE!"
    ElseIf bool_1 Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ���������� ��������!"
        custom_toll = CVErr(xlErrValue) ' "#����!"
    ElseIf Application.WorksheetFunction.IsText(custom_sum) Then
        msg_string = "������� ������� �������� ��� ������; " _
                      & vbCrLf & _
                     "�������� ��������� �� ��������� �������� (�����)!"
        custom_toll = CVErr(xlErrValue) ' "#����!"
    ElseIf bool_2 Then
        msg_string = "������� �������� ����������� ��� ��������� ���:" _
                      & vbCrLf & _
                     "�������� ����������� �������� ��� ������ �� ������;" _
                      & vbCrLf & _
                      "������ ����������. ��������� ��������� ������."
        custom_toll = CVErr(xlErrName)
    ElseIf custom_sum < 0 Or currency_rate < 0 Then
        msg_string = "���������� ��������� ��� ���� ������ �� ����� " & _
                     "���� ������������� ������. " & vbCrLf & _
                     "��������� ���������� ������� ���������."
        custom_toll = CVErr(xlErrNum) ' "#����!"
    ElseIf currency_rate = 0 Then
        msg_string = "� ������� ����������� ������� ������� �� ����." _
                      & vbCrLf & _
                      "��������� ��������� � ������ ���������� � �������!"
        custom_toll = CVErr(xlErrDiv0)
    ElseIf custom_sum_ru >= 0 And _
           IsNumeric(custom_sum_ru) And _
           custom_sum_ru <> "" Then
            If custom_sum_ru >= 0 And custom_sum_ru <= 200000 Then
                custom_toll = 775 / currency_rate
            ElseIf custom_sum_ru > 200000 And custom_sum_ru <= 450000 Then
                custom_toll = 1550 / currency_rate
            ElseIf custom_sum_ru > 450000 And custom_sum_ru <= 1200000 Then
                custom_toll = 3100 / currency_rate
            ElseIf custom_sum_ru > 1200000 And custom_sum_ru <= 2700000 Then
                custom_toll = 8530 / currency_rate
            ElseIf custom_sum_ru > 2700000 And custom_sum_ru <= 4200000 Then
                custom_toll = 12000 / currency_rate
            ElseIf custom_sum_ru > 4200000 And custom_sum_ru <= 5500000 Then
                custom_toll = 15500 / currency_rate
            ElseIf custom_sum_ru > 5500000 And custom_sum_ru <= 7000000 Then
                custom_toll = 20000 / currency_rate
            ElseIf custom_sum_ru > 7000000 And custom_sum_ru <= 8000000 Then
                custom_toll = 23000 / currency_rate
            ElseIf custom_sum_ru > 8000000 And custom_sum_ru <= 9000000 Then
                custom_toll = 25000 / currency_rate
            ElseIf custom_sum_ru > 9000000 And custom_sum_ru <= 10000000 Then
                custom_toll = 27000 / currency_rate
            ElseIf custom_sum_ru > 10000000 Then
                custom_toll = 30000 / currency_rate
            End If
    Else
        msg_string = "�������� ��� ������. " _
                      & vbCrLf & _
                     "������� ������� ������������ ��������. " _
                      & vbCrLf & _
                     "��������� ������, �� ������� ��������� �������!"
        custom_toll = CVErr(xlErrValue) ' "#����!"
    End If

    If TypeName(custom_toll) <> "Error" Then _
        custom_toll = Round(custom_toll, 2)
    
    If msg_string <> "" And msg_flag Then MsgBox msg_string

'   Exit Function
    On Error GoTo 0

'handler:
'        MsgBox "�������� ��� ������. " & _
                "������� ������� ������������ �������� " & _
                 vbCrLf & _
'               "(����, ���������� ��� ������������� ��������)." & _
'                vbCrLf & vbCrLf & _
'               "��������� ������, �� ������� ��������� �������."
'        MsgBox msg_string
'        custom_toll = CVErr(xlErrValue) ' "#����!"

End Function

Public Sub cells_numbering()

'��������� ��������� ���������� ������ ���������, � ������� ���������
'������ ������ (�������������).
'��������� ������������ ������ ������������ (���������) �����.
'������ ���������� ����� ����������� ��� ��������, �.�. ����� ��������
'������� ��� ����������� ���� ����� - ��������� ��������� �� ��������,
'�� ���������������. ��� ���������� �������� (feature) ��������� - �����
'��������� �� ��������� ��� ��������� �������, ������� ��� �����������
'�����. ���� ����� ���������� ������ - ����� ��� ����� �������
'������������ ������� (�������, �������������.�����(�Ш��;...) � �.�.)

'��������� ���������� �� ��������, ������ ���������� � ����� �������
'������ ����������� ���������; ���� � ���� ������ ���������� ���������
'�������� ��� �������� ���� ��� ������ �������� ��� �������� ������
'�������, ������������� �������� - ����� ��������� ����������
'� 1 (�������).

'������������� �������.
'����� �������� �������� � ������� ������. ��� ������������ (���������)
'������ ��������� ����� ������������� �� ������� - � ������� ��� �������
'� ��������� �������� � ����� ������� ������ ���������. �������������
'����� ������ ����� ������� ����������� ���������, ���� ����� ��������
'������� �� ���������� �������. ���� �������� ������ ���� ������ - �����
'������������� ������ ��� �� �������� ��������� ���� (�������� ������ �
'�������� ������������ � ������).


       
    Dim i As Long, n As Long
       
    If IsNumeric(Selection.Cells(1, 1)) And _
            Not IsEmpty(Selection.Cells(1, 1)) And _
            Not Selection.Cells(1, 1) < 1 Then
        n = Selection.Cells(1).Value
    Else: n = 1
    End If
       
    For i = 1 To Selection.Rows.Count
        If Not Selection.Rows(i).Hidden Then
            Selection.Cells(i, 1) = n: n = n + 1
        End If
    Next
    
End Sub
