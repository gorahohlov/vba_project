Attribute VB_Name = "Module1"

Function VLookUp2( _
                  search_value As Variant, _
                  table_rng As Range, _
                  search_col_num As Integer, _
                  result_col_num As Integer, _
                  match_num As Integer _
                 )

'search_value: искомое значение;
'table_rng: таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которой ищутся совпадения;
'result_col_num: номер колонки,  в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'match_num: номер совпадения значения (если совпадения множественные);

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

'search_value: искомое значение;
'table_rng: Таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которой ищутся совпадения;
'result_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'match_num: номер совпадения значения (если совпадения множественные);

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

'search_value: искомое значение;
'table_rng: таблица, диапазон ячеек, в которых ищутся совпадения и
'                   результаты;
'search_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   в которой ищутся совпадения;
'result_col_num: номер колонки, в диапазоне ячеек [table_rng],
'                   из которой извлекаются искомые данные;
'symbols_num: количество первых левых символов исходного значения
'                   (артикула), по которым будут искаться совпадения.

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

Function custom_toll( _
                     custom_sum As Variant, _
                     Optional currency_rate As Single = 1, _
                     Optional msg_flag As Boolean = False _
                    ) As Variant

'Назначение формулы.
'Эта функция подсчитывает сумму  таможенных сборов в зависимости от
'суммы таможенной стоимости передаваемой по ссылке (первый аргумент),
'второй аргумент - необязательный - это курс, на который умножается!
'первый аргумент функции, чтобы получить сумму таможенной стоимости
'(в рублях) для расчета сборов. Если подразумевается что первый аргумент
'(обязательный) это таможенная стоимость в рублях, тогда второй аргумент
'(курс) можно не указывать или поставить 1 (единицу) - такое значение
'будет по умолчанию, если второй аргумент пропущен.

'Что делает функция.
'Если формула ссылается на сумму в валюте (долларах, евро или
'юанях - неважно) - нужно вторым аргументом указать курс, на который
'будет умножаться! первый аргумент, чтобы получить сумму в рублях.
'Далее полученная таможенная стоимость в рублях "прогоняется" по
'конструкции "IF ElseIF Else EndIF" - которая выдает сумму таможенных
'сборов в рублях в зависимости от таможенной стоимости. Если второй
'аргумент указан и отличен от 1 (единицы) - итог из конструкции "IF
' ElseIF Else EndIF" делится на этот курс и результат указывается
'в той же валюте, в которой подразумевается номинирована таможенная
'стоимость (первый) аргумент.

'Параметр msg_flag (третий аргумент) отвечает нужно ли чтобы выводилось
'сообщение с информацией об ошибке, если она случится (если работа
'функции завершиться штатно никаких сообщений не появится).
'Бывает удобно знать в чем проблема когда она возникает; Но нужно
'иметь в виду, что если на листе много ячеек с этой формулой и по
'какой-то причине возникает ошибка в работе (удалены ячейки с исходными
'данными и т.д.) то придется много отщелкивать всплывающие окна
'сообщений при каждом! перерасчете листа или книги.
'По умолчанию формула отрабатывает без сообщений об ошибках даже, если
'они возникают (т.е. этот параметр "msg_string" по умолчанию отключен).
'Формула по умолчанию просто возвращает результат значением
'соответствующей ошибки в случае некорретных данных.

'Особенности работы.
'Функция обрабатывает сутуации, когда ячейка, на которую ссылается
'формула, содержит числовые данные форматированные в формате Даты,
'или Логического значения (True, False) или содержит значения в
'отрицательном диапазоне, а также если ячейка содержит текстовые данные.
'В таких случаях, когда данные, на которые ссылается функция,
'форматированы как ДАТА, ИСТИНА или ЛОЖЬ или содержат отрицательные
'значения, текстовые значения - выдает значение ошибок: "#ЗНАЧ!",
'"#ЧИСЛО!", "#ИМЯ?", "#ДЕЛ/0!".
'Также я сделал чтобы в таких случаях выдавалось
'сообщение-предупреждение "о неверных данных переданных в формулу".

'Дисклеймер (оговорка).
'Сумма таможенных сборов считается в этой функции по стандартному
'алгоритму - т.е. в зависимости от суммы таможенной стоимости - без
'учёта кодов, по которым сразу начисляется 30000 руб. Другими словами
'в функции вообще не обрабатываются такие ситуации, когда сумма
'таможенных сборов зависит от кода ТНВЭД.


    Dim custom_sum_ru As Variant
    Dim bool_1 As Boolean, _
        bool_2 As Boolean
    Dim msg_string As String

    On Error Resume Next

    If TypeName(custom_sum) = "Range" Then
        bool_1 = TypeName(custom_sum.Value) = "Boolean"
        bool_2 = TypeName(custom_sum.Value) = "Error"
    Else
        bool_1 = TypeName(custom_sum) = "Boolean"
        bool_2 = TypeName(custom_sum) = "Error"
    End If

    custom_sum_ru = custom_sum * currency_rate
    
    If IsDate(custom_sum) Then
        msg_string = "Функции передан неверный тип данных; " _
                      & vbCrLf & _
                     "Аргумент ссылается на дату!"
        custom_toll = CVErr(xlErrValue)
    ElseIf bool_1 Then
        msg_string = "Функции передан неверный тип данных; " _
                      & vbCrLf & _
                     "Аргумент ссылается на логическое значение!"
        custom_toll = CVErr(xlErrValue)
    ElseIf Application.WorksheetFunction.IsText(custom_sum) Then
        msg_string = "Функции передан неверный тип данных; " _
                      & vbCrLf & _
                     "Аргумент ссылается на строковое значение (текст)!"
        custom_toll = CVErr(xlErrValue)
    ElseIf bool_2 Then
        msg_string = "Функции передано неизвестное или удаленное имя:" _
                      & vbCrLf & _
                     "неверный именованный диапазон или ссылка на ячейку;" _
                      & vbCrLf & _
                      "Ошибка синтаксиса. Проверьте введенные данные."
        custom_toll = CVErr(xlErrName)
    ElseIf custom_sum < 0 Or currency_rate < 0 Then
        msg_string = "Таможенная стоимость или курс валюты не может " & _
                     "быть отрицательным числом. " & vbCrLf & _
                     "Проверьте переданные функции аргументы."
        custom_toll = CVErr(xlErrNum)
    ElseIf currency_rate = 0 Then
        msg_string = "В формуле предпринята попытка деления на ноль." _
                      & vbCrLf & _
                      "Проверьте аргументы и ссылки переданные в формулу!"
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
        msg_string = "Неверный тип данных. " _
                      & vbCrLf & _
                     "Формуле передан некорректный аргумент. " _
                      & vbCrLf & _
                     "Проверьте данные, на которые ссылается формула!"
        custom_toll = CVErr(xlErrValue)
    End If

    If TypeName(custom_toll) <> "Error" Then _
        custom_toll = Round(custom_toll, 2)
    
    If msg_string <> "" And msg_flag Then MsgBox msg_string

    On Error GoTo 0

End Function
