Attribute VB_Name = "Module1"
Public vlookup_arg3 As Integer
Public vlookup_arg4 As Integer
Public upper_interval As Integer
Public lower_interval As Integer
Public processing_row_num As Integer
Public article_col_num As Integer
Public vlookup_table_rng As Range
Public cancel_flag As Boolean

Sub highlight_codes_39()

' Данный макрос предназначен для выделения ячеек с кодами ТН ВЭД
' (10 знаков), которые попадают в Перечень товаров, которые должны
' декларироваться отдельным товаром в ДТ согласно Решению Коллегии
' ЕЭК №39 от 26.04.2012. "Товары, помещаемые под таможенную процедуру
' выпуска для внутреннего потребления, декларируются как один товар,
' только если товары имеют один товарный знак, марку, модель, артикул
' и обладают одинаковыми техническими и (или) коммерческими
' характеристиками."

' ПРИНЦИП работы.
' Пользователь выделяет произвольную область рабочего листа.
' В этой области могут быть любые данные: текст, числа, формулы, пустые
' ячейки и т.д. Значение ячеек сравнивается с кодами из Перечня. Если
' ячейка содержит код ТНВЭД из Перечня, такая ячейка выделяется
' контрастным форматированием. В дальнейшем используя это форматирование
' можно фильтровать ячейки, обрабатыать их и т.д. Все остальные значения
'выделенного диапазона игнорируются.

' ОСОБЕННОСТИ работы макроса, которые нужно иметь в виду.
' Макрос может выделять ячейки со значениями меньшими десяти знаков.
' Т.е. если ячейки содержат 4, 6, 9 и т.п. знаков - макрос тоже может
' их обработать и отметить контрастным форматированием. Но быть
' уверенным в корректном результате можно только там где стоят
' десятизначные коды! Т.к. в Перечне есть исключения и коды на уровне
' десяти знаков.

' СЛЕДУЕТ ИМЕТЬ В ВИДУ, что к выделенной области в процессе работы
' макроса применяется текстовый (строковый) формат. Т.е. там где стояли
' числовые значения, денежный, финансовый, форматы даты и прочие
' пользовательские форматы - такое форматирование удаляется - останется
' строковый формат (текст). Сами данные НЕ удаляются!

' Данные в формате даты "превращаются" в числа! Вернуть формат даты можно
' через "Формат Ячейки".

' Макрос выделит контрастным цветом все коды из товарной позиции 8523.
' Однако нужно помнить, что декларированию отдельным товаром в ДТ
' подлежат только "Носители готовые незаписанные" из 8523!

' Если у ячейки было какое-то форматирование (похожее, такое же или
' иное чем устанавливает макрос - любое) оно сохраняется! Т.е. если
' код несодержащийся в Перечне Решения №39 был форматирован контрасным
' цветом на манер как в этом макросе - оно сохраниться и ячейка будут
' выделена как будто она подлежит декларированию отдельным товаром.
' Поэтому будет хорошим тоном, перед началом применения данных макросов,
' очищать форматирование ячеек с кодами.

' Форматирование кодов по Перечню Решения Коллегии ЕЭК №39 (декларирование
' одним товаром) - белый жирный шрифт, черный фон ячейки, оканктовка ячейки
' двойной рамкой.

' Форматирование кодов по Перечню Постановления Правительства №342
' (таможенны сбор 30000) - желтый фон ячейки, цвет шрифта авто,
' оканктовка двойной рамкой.

' Форматирование кодов, которы упоминаются в обоих перечнях - фиолетовый
' фон, белый шрифт, двойная белая рамка вокруг ячейки.
 
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
    With Selection
        a = .Value
        .NumberFormat = "General"
        .Value = a
        .NumberFormat = "@"
    End With
     
    Set rng = Selection
     
    For Each cell In rng
        
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
                    .FontStyle = "обычный"
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
    
    Next cell
 
    t = Timer - t
    MsgBox "Готово." & " Время выполенния: " & Round(t, 1) & " sec"
 
End Sub

Private Function IsInArray( _
                           arr As Variant, _
                           match_code As Variant _
                          ) As Boolean

'    IsInArray = (UBound(Filter(arr, match_code)) > -1)
  
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

    With Selection
        a = .Value
        .NumberFormat = "General"
        .Value = a
        .NumberFormat = "@"
    End With
 
    Set rng = Selection
    
    For Each cell In rng
        
        If IsInArray(arr_342_position, Left(cell.Value, 4)) Then
        
            cond_08 = Left(cell.Value, 4) = var_342_04
            cond_09 = IsInArray(arr_342_06, Left(cell.Value, 6))
            cond_10 = IsInArray(arr_342_09, Left(cell.Value, 9))
            cond_11 = IsInArray(arr_342_10, Left(cell.Value, 10))
            
            bool_02 = cond_08 Or cond_09 Or cond_10 Or cond_11
            
            If bool_02 Then
                
                With cell.Font
                    .Name = "Cambria"
                    .FontStyle = "обычный"
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
               
    Next cell
        
    t = Timer - t
    MsgBox "Готово." & " Время выполенния: " & Round(t, 1) & " sec"
    
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
    Dim t As Single, flg As Single
    
    
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
    
    With Selection
        a = .Value
        .NumberFormat = "General"
        .Value = a
        .NumberFormat = "@"
    End With
     
    Set rng = Selection
     
    For Each cell In rng
    
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
            
            flg = IIf(bool_01 = True And bool_02 = False, 1, _
                    IIf(bool_01 = False And bool_02 = True, 2, _
                        IIf(bool_01 = True And bool_02 = True, 3, 0)))
    
            Select Case flg
                Case 1
                    
                    With cell.Font
                        .Name = "Cambria"
                        .FontStyle = "обычный"
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
                        .FontStyle = "обычный"
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
                        .FontStyle = "обычный"
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
    
    Next cell
        
    t = Timer - t
    MsgBox "Готово." & " Время выполенния: " & Round(t, 1) & " sec"

End Sub

Function VLookUp2( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: искомое значенике;
'table_rng: таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которой ищутся совпадения;
'result_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'match_num: номер совпадения значения (если совпадения множественные);

    Dim flg As Boolean
    Dim i As Integer
    Dim iCount As Integer
    
    flg = False
    
    For i = 1 To table_rng.Rows.Count
        
        If table_rng.Cells(i, search_col_num) = search_value Then
            iCount = iCount + 1
        End If
        
        If iCount = match_num Then
            VLookUp2 = table_rng.Cells(i, result_col_num)
            flg = True
            Exit For
        End If
    
    Next i
    
    If flg = False Then
        VLookUp2 = CVErr(xlErrNA) ' "#Н/Д"
    End If

End Function

Function VLookUp3( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: искомое значение;
'table_rng: Таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которойищутся совпадения;
'result_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'match_num: номер совпадения значения (если совпадения множественные);

    Dim flg As Boolean
    Dim i As Integer
    Dim iCount As Integer
    
    flg = False
    
    For i = 1 To table_rng.Rows.Count
        
        If search_value Like table_rng.Cells(i, search_col_num) Then
            iCount = iCount + 1
        End If
    
        If iCount = match_num Then
            VLookUp3 = table_rng.Cells(i, result_col_num)
            flg = True
            Exit For
        End If
    
    Next i
    
    If flg = False Then
        VLookUp3 = CVErr(xlErrNA) ' "#Н/Д"
    End If

End Function

Function VLookUp4( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  Optional symbols_num As Integer = 0 _
                 )

'search_value: искомое значение;
'table_rng: таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которой ищутся совпадения;
'result_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'symbols_num: количество первых левых символов артикула (искомого
'                   значения), по которым будут искаться совпадения.

    Dim flg As Boolean
    Dim i As Integer
    
    flg = False
    
    For i = 1 To table_rng.Rows.Count
        
        If symbols_num = 0 Then
            If table_rng.Cells(i, search_col_num) = search_value Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                flg = True
                Exit For
            End If
        Else
            If Left(table_rng.Cells(i, search_col_num), symbols_num) = _
                                    Left(search_value, symbols_num) Then
                VLookUp4 = table_rng.Cells(i, result_col_num)
                flg = True
                Exit For
            End If
        End If
            
    
    Next i
    
    If flg = False Then
        VLookUp4 = CVErr(xlErrNA) ' "#Н/Д"
    End If

End Function

Sub partial_match()
Attribute partial_match.VB_ProcData.VB_Invoke_Func = "Q\n14"

'Процедура поиска и подтягиваня значений по частичному
'совпадению артикула. По сути это несколько измененная функциональность
'функции ВПР (VLookUp). Т.е. работает как ВПР только позволяет обрезать
'артикулы (значения по которым ведется поиск, сопоставление диапазонов)
'до какого-то количества знаков.

'Сначала процедура ищит совпадения (артикулов) по полному значению.
'Если не находит - запускается цикл поиска по частичному совпадению.

'У пользователя запрашивается количество значимых символов артикула
'(например 9) и последовательно ищутся совпадения по количеству
'символов 9, 8, 7, 6 - на глубину не более 4 символов.

'Позже сделал изменил функциональность. В пользовательской форме
'запрашивается, кроме всего прочего - количество значимых символов
'артикула по которым будет искться совпадения (число "от" и число "до").
'Пользователь может настраивать их с помощью элементов управления -
'spinbutton или непосредственно вводить числовые значения в текстовоые
'поля. Т.е. "глубина" поиска совпадений по неполному артикулу может быть
'больше 4 символов - количество символов "глубины" теперь настроиваемо,
'его можно указывать произвольно в пользовательской форме.

'Добавил функциональность чтобы процедура не обрабатывала скрытые или
'отфильтрованные строки. Это будет востребовано при множественном
'использовании процедуры. Сначала ищутся совпадения по полному артикулу.
'Далее найденные позиции фильтруются. Остальные строки обрабатываются
'по частичному поиску артикула (например с параметрами от 16 до 13).
'Далее отфильтровываются найденные позиции, а оставшиеся обрабатываются
'с параметрами допустим допустим "от 12 до 8". Скрытые (обработанные
'ранее) ячейки при этом не обрабатываются заново.

'Если процедура не находит в источнике данных совпадений в
'соответствующей ячейке сохраняется значение ошибки xlErrNA, с помощью
'функции CVErr(); Функция IsError() выдает значение True на любую
'ошибку xlErrValue, xlErrNA, xlErrRef, xlErrNull, xlErrName, xlErrDiv0,
'xlErrNum; если нужно чтобы функция реагировала исключительно на
'ошибку xlErrNA - нужно что-то придумывать, потому что включить в
'условие такую инструкцию: hscode = CVErr(2042) не получится - такая
'инструкция выдаст ошибку, если hscode не будет иметь значение ошибки,
'а будет, например содержать значение литерала, etc.
'Решение здесь будет очень простое. Сначала проверяем является ли
'переменная ошибкой (например, [IsError(hscode)] или
'конструкция [TypeName(hscode) = "Error"]), а потом, если убедились что
'переменная содержит ошибку - [IF hscode = CVErr(2042) Then etc].

'Еще одна фича, добавленная позже - процедура не обрабатывает скрытые
'(отфильтрованные) строки, ячейки. Что удобно, поскольку после того как
'один раз отработала процедура - можно отфильтровать данные, по которым
'не подтянулись, не нашлись данные и обработать их этой же процедурой,
'но с другими параметрами. При этом ранее обработанные ячейки не
'обрабатываются вновь.


'    Dim rng  As Range
    Dim article_col_rng As Range

'    переопределил эти переменные как глобальные чтобы передавать
'    их значения в модуль кода пользовательской формы UserForm1:
'    Dim vlookup_arg3 As Integer, _
         vlookup_arg4 As Integer, _
'        upper_interval As Integer, _
'        lower_interval As Integer
'    Dim processing_row_num As Integer, _
'        article_col_num As Integer
'    Dim vlookup_table_rng As Range
    
    Dim counter1 As Long
    Dim counter2 As Integer
    Dim hscode As Variant, _
        description_31 As Variant, _
        description_art As Variant
    Dim t As Single
    Dim working_wbook_name As Variant
    Dim working_sheet_name As Variant
    
    Dim article_length As Integer
    
    Dim upper_ As Integer, lower_ As Integer
'   контекстные значения верхнего и нижнего интервала символов для
'   каждой строки выделенного интервала (в зависимости от длины
'   артикула - т.е. для цикла "For counter2 <...>"), тогда как
'   upper_interval, lower_interval это глобальные настройки для всего
'   диапазона Selection, для всей процедуры;


    t = Timer

    working_wbook_name = ActiveWorkbook.Name
    working_sheet_name = ActiveSheet.Name

    
    On Error Resume Next
    
    Set article_col_rng = Application.InputBox("Введите колонку с признаком," _
                    & " по которому будем осуществлять поиск", _
                    "Ввод колонки с искомыми данными (артикулами)", Type:=8)
'    первая (левая верхняя) ячейка выделенного диапазона будет
'    по умолчанию указывать на требуемую колонку с артикулом;
'    диапазон может быть любым; главное чтобы левая верхняя ячейка
'    (левая колонка) указывала на требуемые данные (артикулы).
    If Err.Number <> 0 Then Set article_col_rng = ActiveCell.Offset(0, -1)
    On Error GoTo -1
    
    Set vlookup_table_rng = Application.InputBox("Введите диапазон откуда " _
                    & "требуется подгрузить данные", _
                    "Ввод диапазона (откуда будем получать данные)", Type:=8)
    If Err.Number <> 0 Then Set vlookup_table_rng = ActiveCell.Offset(0, -1)
    
    On Error GoTo 0

    Windows(working_wbook_name).Activate
        
    processing_row_num = Selection.Row
    article_col_num = article_col_rng.Column
    
'    Call UserForm1.UserForm_Initialize( _
                                       processing_row_num, _
                                       article_col_num, _
                                       vlookup_table_rng _
                                      )
    UserForm1.Show
    
    Workbooks(working_wbook_name).Sheets(working_sheet_name).Activate
    
'    For Each cell In Selection
'
'        hscode = VLookUp4( _
                    Cells(processing_row_num, article_col_num), _
                    vlookup_table_rng, _
                    9, _
                    1, _
                    0 _
                   )
'
'        If IsError(hscode) Then
'
'            For counter = upper_interval To upper_interval - 3 Step -1
'
'                hscode = VLookUp4( _
                           Cells(processing_row_num, article_col_num), _
                           vlookup_table_rng, _
                           9, _
                           1, _
                           counter _
                          )
'
'                If Not IsError(hscode) Then
'                    cell = hscode
'                    Select Case counter
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
    
    If cancel_flag Then
        MsgBox "Процедура прервана пользователем."
        Exit Sub
    End If
    
    For counter1 = 1 To Selection.Count Step Selection.Columns.Count
        rw_num = Application.WorksheetFunction.RoundUp(counter1 _
                                        / Selection.Columns.Count, 0)
        If Not Selection.Rows(rw_num).Hidden Then
        hscode = VLookUp4( _
                        Cells(processing_row_num, article_col_num), _
                        vlookup_table_rng, _
                        vlookup_arg3, _
                        vlookup_arg4, _
                        0 _
                        )
        If Not IsError(hscode) Then ' <> "#Н/Д" Then
            description_31 = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                vlookup_arg3, _
                                5, _
                                0 _
                                )
            description_art = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                vlookup_arg3, _
                                6, _
                                0 _
                                )
        Else
            description_31 = CVErr(xlErrNA)  ' "#Н/Д"
            description_art = CVErr(xlErrNA) ' "#Н/Д"
        End If
        
        If IsError(hscode) Then ' = "#Н/Д" Then
        
            article_length = Len(Cells(processing_row_num, article_col_num))
            If article_length >= upper_interval Then
                upper_ = upper_interval
                lower_ = lower_interval
            ElseIf article_length < upper_interval And _
               article_length >= lower_interval Then
               upper_ = article_length
               lower_ = lower_interval
            ElseIf article_length < lower_interval Then
                upper_ = article_length
                lower_ = article_length
            End If

            For counter2 = upper_ To lower_ Step -1
        
                hscode = VLookUp4( _
                            Cells(processing_row_num, article_col_num), _
                            vlookup_table_rng, _
                            vlookup_arg3, _
                            vlookup_arg4, _
                            counter2 _
                            )
                
                If Not IsError(hscode) Then ' <> "#Н/Д" Then
                    description_31 = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                vlookup_arg3, _
                                5, _
                                counter2 _
                                )
                    description_art = VLookUp4( _
                                Cells(processing_row_num, article_col_num), _
                                vlookup_table_rng, _
                                vlookup_arg3, _
                                6, _
                                counter2 _
                                )
                Else
                    description_31 = CVErr(xlErrNA)  ' "#Н/Д"
                    description_art = CVErr(xlErrNA) ' "#Н/Д"
                End If
        
                If Not IsError(hscode) Then ' <> "#Н/Д" Then
                    
                    Select Case counter2
                        Case upper_ - 0
                            Call paint_cells( _
                                             Selection.Columns.Count, _
                                             13819376, _
                                             counter1 _
                                            )
                            
                        Case upper_ - 1
                            Call paint_cells( _
                                             Selection.Columns.Count, _
                                             11321572, _
                                             counter1 _
                                            )
                            
                        Case upper_ - 2
                            Call paint_cells( _
                                             Selection.Columns.Count, _
                                             8823768, _
                                             counter1 _
                                            )
                            
                        Case 1 To upper_ - 3
                            Call paint_cells( _
                                             Selection.Columns.Count, _
                                             4025277, _
                                             counter1 _
                                            )
                            
                    End Select
'                   альтернативные заливки (оттенки зеленого):
'                        14348258, 11854022, 9359529, 3506772, 2315831
                    
                    Exit For
                End If
        
            Next counter2
        
        End If
        
        Selection.Cells(counter1) = hscode
        If Selection.Columns.Count = 2 Then
            Selection.Cells(counter1 + 1) = description_31
        ElseIf Selection.Columns.Count = 3 Then
            Selection.Cells(counter1 + 1) = description_31
            Selection.Cells(counter1 + 2) = description_art
        End If
        
        End If
        processing_row_num = processing_row_num + 1
    Next counter1
    
    t = Timer - t
    MsgBox "Готово." & " Время выполенния: " & Round(t, 1) & " sec"
    

'ПОДХОДЫ К НАПИСАНИЮ ПРОЦЕДУРЫ
'---

'Первый вариант: Использование метода Application.Inputbox
'для ввода формулы:
'   ActiveCell = Application.InputBox(prompt:= _
                "Введите формулу: _
    VLookUp2(SearchValue;Table;SearchColNum;ResultColNum;match_num)", _
                 Title:="Ввод формулы", Default:="=VLookUp2()", Type:=0)
'после ввода формулы без аргументов, можно войти в ячейку и вызвать
'с помощью Shift+F3 - окно "Аргументы функции" для заполнения
'её аргументов;

'Второй вариант: использовния диалогового окна мастера функций:
'            Application.CommandBars.ExecuteMso "FunctionWizard"

'Третий вариант: запросить все параметры для функции VLookUp
'с помощью нескольких методов Application.InputBox


'ActiveCell.FormulaLocal = "=VLookUp3()"
'далее Shift+F3 - вызов диалогового окна "Аргументы функции"

'или

'ActiveCell.FormulaLocal = "=VLookUp3("
'далее Ctrl+A для вызова диалога "Аргументы функции"

'или

'ActiveCell = Application.InputBox( _
             prompt:="Введите формулу: VLookUp2(SearchValue;<...>)", _
             Title:="Ввод формулы", _
             Default:="=VLookUp2()", _
             Type:=0 _
             )
'далее Enter и Shift+F3 для вызова диалога "Аргументы функции"

'Set Rng1 = Application.InputBox("Введите колонку с признаком по " & _
                                "которому будем осуществлять поиск", _
                                "Ввод колонки", _
                                 Type:=8)
'Print Rng1.Column

'Set Rgn2 = Application.InputBox( _
                "Введите диапазон откуда требуется подгрузить данные", _
                "Ввод диапазона", Type:=8)
'Print Rgn2.Column

End Sub

Private Sub paint_cells( _
                        sel_col_cnt As Integer, _
                        color_index_val As Long, _
                        cell_pointer As Long _
                       )

    Selection.Cells(cell_pointer).Interior.Color = color_index_val
    
    If sel_col_cnt = 2 Then
        Selection.Cells(cell_pointer + 1).Interior.Color = color_index_val
    ElseIf sel_col_cnt >= 3 Then
        Selection.Cells(cell_pointer + 1).Interior.Color = color_index_val
        Selection.Cells(cell_pointer + 2).Interior.Color = color_index_val
    End If

End Sub

Function custom_toll( _
                     custom_sum As Variant, _
                     Optional currency_rate As Single = 1 _
                    ) As Double
        
'ЭТО СТАРАЯ ВЕРСИЯ ФУНКЦИИ. НОВАЯ, РАБОЧАЯ ВЕРСИЯ В МОДУЛЕ my_funcs.xlam

'Назначение формулы.
'Эта функция подсчитывает сумму таможенных сборов в зависимости от суммы
'таможенной стоимости передаваемой по ссылке (первый аргумент), второй
'аргумент - необязательный - это курс, на который умножается! первый
'аргумент функции, чтобы получить сумму таможенной стоимости (в рублях)
'для расчета сборов. Если подразумевается что первый аргумент
'(обязательный) это таможенная стоимость в рублях, тогда второй аргумент
'(курс) можно не указывать или поставить 1 (единицу) - такое значение
'будет по умолчанию, если второй аргумент пропущен.

'Что делает функция.
'Если формула ссылается на сумму в валюте (долларах, евро или юанях -
'неважно) - нужно вторым аргументом указать курс, на который будет
'умножаться! первый аргумент, чтобы получить сумму в рублях. Далее
'полученная таможенная стоимость в рублях "прогоняется" по
'конструкции "IF ElseIF Else EndIF" - которая выдает сумму таможенных
'сборов в рублях в зависимости от таможенной стоимости. Если второй
'аргумент указан и отличен от 1 (единицы) - итог из конструкции "IF
'ElseIF Else EndIF" делится на этот курс и результат указывается в той
'же валюте, в которой подразумевается номинирована таможенная стоимость
'(первый) аргумент.

'Особенности работы.
'Функция обрабатывает сутуации, когда ячейка, на которую ссылается
'формула, содержит числовые данные форматированные в формате Даты,
'или Логического значения (True, False) или содержит значения в
'отрицательном диапазоне, а также если ячейка содержит текстовые данные.
'В таких случаях, когда данные, на которые ссылается функция,
'форматированы как ДАТА, ИСТИНА или ЛОЖЬ или содержат отрицательные
'значения, текстовые значения - выдает значение ошибки "#ЗНАЧ!".
'Также я сделал (но потом закомментировал) чтобы в таких случаях
'выдавалось сообщение-предупреждение "о неверном типе данных переданных
'в формулу" - сообщение, которое размещено после ключевого слова
'heandler.
'Я отключил это сообщение, поскольку, если таких формул (которые выдают
'ошибку) будет много на листе - при каждом открытии файла и при каждом
'пересчете листа будет выскакивать много таких сообщений, которые нужно
'будет "отщелкать" Enter-ом. Сообщения MsgBox, которые размещеные в
'управляющей конструкции "IF ElseIF ELSE EndIF" не рабоатают (даже если
'их раскомментировать) - не суть.

'Дисклеймер (оговорка).
'Сумма таможенных сборов считается в этой функции по стандартному
'алгоритму - т.е. в зависимости от суммы таможенной стоимости - без
'учёта кодов, по которым сразу начисляется 30000 руб. Другими словами в
'функции вообще не обрабатываются такие ситуации, когда сумма таможенных
'сборов зависит от кода ТНВЭД.
    
    
    Dim custom_sum_ru As Double
    Dim bool As Boolean
    Dim msg_string As String

    On Error GoTo handler:
    
    If TypeName(custom_sum) = "Range" Then
        bool = TypeName(custom_sum.Value) = "Boolean"
    Else
        bool = TypeName(custom_sum) = "Boolean"
    End If

    custom_sum_ru = custom_sum * currency_rate
    
    If IsDate(custom_sum) Then 'TypeName(custom_sum.Value) = "Date" Then
        custom_toll = CVErr(xlErrValue) ' "#DATE!"
        msg_string = "Функции передан неверный тип данных; " _
                      & vbCrLf & _
                     "Аргумент ссылается на дату!"
    ElseIf bool Then custom_toll = CVErr(xlErrValue) ' "#ЗНАЧ!"
        msg_string = "Функции передан неверный тип данных; " _
                      & vbCrLf & _
                     "Аргумент ссылается на логическое значение!"
    ElseIf custom_sum_ru >= 0 And custom_sum_ru <= 200000 Then
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
    ElseIf custom_sum_ru < 0 Then
        custom_toll = CVErr(xlErrValue) ' "#ЗНАЧ!"
        msg_string = "Таможенная стоимость не может быть " & _
                     "отрицательным числом. " & vbCrLf & _
                     "Проверьте переданные функции аргументы."
    Else
        custom_toll = CVErr(xlErrValue) ' "#ЗНАЧ!"
        msg_string = "Неверный тип данных. " & vbCrLf & _
                     "Формуле передан некорректный аргумент. " _
                      & vbCrLf & _
                     "Проверьте данные, на которые ссылается формула!"
    End If
    
    custom_toll = Round(custom_toll, 2)
    
    Exit Function

handler:
'        MsgBox "Неверный тип данных. " & _
                "Формуле передан некорректный аргумент " _
                & vbCrLf & _
                "(дата, логическое или отрицательное значение)." _
                & vbCrLf & vbCrLf & _
                "Проверьте данные, на которые ссылается формула."
'        MsgBox msg_string
        custom_toll = CVErr(xlErrValue) '= "#ЗНАЧ!"

End Function

Public Sub нумерация_ячеек()

'Процедура позволяет нумеровать строки диапазона, в котором некоторые
'строки скрыты (отфильтрованы). Нумерация производится только
'отображаемых (нескрытых) строк. Номера отобраемых строк вставляются
'как значения, т.е. после удаления фильтра или отображения всех
'строк - сделанная нумерация не меняется, не пересчитывается.
'Это намеренное свойство (feature) процедуры - чтобы нумерация
'не сбивалась при изменении фильтра, скрытия или отображения строк.
'Если нужен адаптивный фильтр - лучше для таких случаев использовать
'формулы (АГРЕГАТ, ПРОМЕЖУТОЧНЫЕ.ИТОГИ(СРЁТЗ;...) и т.д.)

'Нумерация начинается со значения, которе содержится в левой верхней
'ячейке выделенного диапазона; Если в этой ячейке содержатся текстовые
'значения или значения даты или пустое значение или значение меньше
'единицы, отрицательные значения -  тогда нумерация начинается
'с 1 (единицы).

'Использование макроса.
'Нужно выделить диапазон и вызвать макрос. Все отображаемые (нескрытые)
'строки диапазона будут пронумерованы по порядку - с единицы или
'начиная с числового значения в левой верхней ячейке диапазона.
'Пронумерована будет только левая колонка выделенного диапазона,
    
'одна ячейка - будет пронумерована только она по правилам описанным
'выше (учитывая формат и значение содержащееся в ячейке).


       
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
