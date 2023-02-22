Attribute VB_Name = "getlogin"
' Марос создания логина из Фамилии Имени Отчества
' Разработал Виктор Борисенко
' victor.borisenko81@gmail.com


' Функция ТРАНСЛИТ() Принимает значение и производит транслитерацию русских букв на английские, согласно стандарта ISO-R9-1968 и возвращает готовое значение в нижнем регистре
' Пример использования
'  =ТРАНСЛИТ("СТЕПАН")
'  =ТРАНСЛИТ(A1) где A1 это адрес ячейки


Public Function ТРАНСЛИТ(ТЕКСТ As String) As String
    Dim Rus As Variant, Eng As Variant
    Dim I As Long, j As Integer
    Dim simb As String
    Dim FindRus As Boolean
    Dim simbtrans As String
    Dim MergeText As String
 
    Rus = Array("а", "б", "в", "г", "д", "е", "ё", "ж", "з", "и", "й", "к", "л", "м", "н", "о", "п", "р", "с", "т", "у", "ф", "х", "ц", "ч", "ш", "щ", "ъ", "ы", "ь", "э", "ю", "я", "А", "Б", "В", "Г", "Д", "Е", "Ё", "Ж", "З", "И", "Й", "К", "Л", "М", "Н", "О", "П", "Р", "С", "Т", "У", "Ф", "Х", "Ц", "Ч", "Ш", "Щ", "Ъ", "Ы", "Ь", "Э", "Ю", "Я")
    Eng = Array("a", "b", "v", "g", "d", "e", "e", "zh", "z", "i", "j", "k", "l", "m", "n", "o", "p", "r", "s", "t", "u", "f", "kh", "ts", "ch", "sh", "shch", "", "y", "", "eh", "yu", "ya", "A", "B", "V", "G", "D", "E", "E", "ZH", "Z", "I", "J", "K", "L", "M", "N", "O", "P", "R", "S", "T", "U", "F", "KH", "TS", "CH", "SH", "SHCH", "", "Y", "", "EH", "YU", "YA")
 
    For I = 1 To Len(ТЕКСТ)
        simb = Mid(ТЕКСТ, I, 1)
        FindRus = False
        For j = 0 To 65
            If Rus(j) = simb Then
                simbtrans = Eng(j)
                FindRus = True
                Exit For
            End If
        Next
        If FindRus Then MergeText = MergeText & simbtrans Else MergeText = MergeText & simb
    Next
    ' Переводим все буквы в нижний регистр
    ТРАНСЛИТ = LCase(MergeText)
End Function

' Функция ТРАНСЛИТПЕРВЫЕБУКВЫ() Принимает значение и производит транслитерацию русских букв на английские, согласно стандарта ISO-R9-1968,
' оставляет только первую букву из каждого слова в переданном тесте и возвращает готовое значение в нижнем регистре
' Пример использования
'  =ТРАНСЛИТ(A1) где A1 это адрес ячейки
'
' Внимание!!! функция принимает только ячейки либо массив

Function ТРАНСЛИТПЕРВЫЕБУКВЫ(rng As Range) As String
    Dim arr
    Dim I As Long
    Dim tmp As String
    arr = VBA.Split(rng, " ")
    tmp = ТРАНСЛИТ(Join(arr, ", "))
    arr = VBA.Split(tmp, " ")

    If IsArray(arr) Then
        For I = LBound(arr) To UBound(arr)
            ТРАНСЛИТПЕРВЫЕБУКВЫ = ТРАНСЛИТПЕРВЫЕБУКВЫ & Left(arr(I), 1)
        Next I
    Else
        ТРАНСЛИТПЕРВЫЕБУКВЫ = Left(arr, 1)
    End If
End Function
