Private Sub CommandButton1_Click()
 'проведение частотного анализа текста в верхнем регистре и указание в процентах сколько раз встречается каждая буква'
 Dim mARR(), mStr$, dict As Object, i&, counter&, el 'Объявление переменных и массива-словаря'
        ActiveSheet.Cells.Delete() 'Очистка страницы после предыдущего анализа'
        'Счет букв в верхнем регистре без пробелов'
         mStr = Application.Trim(UCase("Это началось. " & _
            "Когда были, выкованы Великие кольца: " & _
                "три были даны эльфам, " & _
                    "семь повелителям гномов, " & _
                        "и девять колец были подарены расе людей -" & _
                            "больше всего жаждущим власти. " & _
                                "В этих кольцах содержалось могущество," & _
                                    "но они все были обмануты. " & _
                                        "Ибо было создано еще одно кольцо, " & _
                                            "чтобы править ими всеми.")) 
        [a1] = mStr
        dict = CreateObject("scripting.dictionary") 'Создание массива-словаря'
        dict.comparemode = vbBinaryCompare 'Сравнение массива-словаря по бинарному признаку, 
для определения совпадающих символов'
        For i = 1 To Len(mStr)
            If IsLetter(Mid(mStr, i, 1)) Then 'При возврате строчного массива со всеми символами'
                If dict.Exists(Mid(mStr, i, 1)) Then 'Если возврат строчного массива-словаря по ключу'
                    'Доступ к элементам строчного массива-словаря по ключу и перевод их в лонг тип'
                    dict.Item(Mid(mStr, i, 1)) = CLng(dict.Item(Mid(mStr, i, 1))) + 1
                Else
                    dict.Add(Mid(mStr, i, 1), 1) 'Добавление в массив-словарь ключа на новый символ'
                End If
            End If
        Next
        counter = 0
    ReDim mARR(1 To dict.Count, 1 To 3) 'поправка в массиве в связи с разделением заглавных, 
строчных и кириллицы букв по Ascii коду'
        For Each el In dict.keys 'Счетчик всех символов в ключе'
            counter = counter + 1
            mARR(counter, 1) = el 'добавление в категорию ключей el'
            mARR(counter, 2) = dict.Item(el) 'обращение к букве по ключу el'
            'округление до типа данных double, умножение ключа буквы на 100% 
и деление на сумму ключей этой буквы начиная со второй'
            mARR(counter, 3) = Round(CDbl((dict.Item(el) * 100) / Application.Sum(dict.items)), 2)
        Next
        Cells(4, 1).Value = "Letters": Cells(4, 2).Value = "Quantity"
        Cells(4, 3) = Space(6) & "%"
        'Установка размера ячейки для вывода массива'
        Cells(5, 1).Resize(UBound(mARR, 1), UBound(mARR, 2)).Value = mARR
        Erase mARR: dict = Nothing 'Очистка массива'
        MsgBox (Space(12) & "Дело сделано")
        
        
        
        Private Function IsLetter(strValue As String) As Boolean
        'Проверка перевода значений в строковый тип'
        Dim intPos As Integer
        For intPos = 1 To Len(strValue)
'Выборка со сравнением по коду символов из всех значений строки'
            Select Case Asc(Mid(strValue, intPos, 1)) 
'Разделение заглавных, строчных и кириллицы букв по ASCII коду'
                Case 65 To 90, 97 To 122, 192 To 255 
                    IsLetter = True
                Case Else
                    IsLetter = False
                    Exit For
            End Select
        Next
    End Function

End Sub

Private Sub CommandButton2_Click()
'Сброс выборки и закрытие формы'
Selection.Collapse
UserForm1.Hide
Set UserForm1 = Nothing
End Sub