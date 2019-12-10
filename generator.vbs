Const CHARACTERS = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!#$%&()*+,-./:;<=>?@[\]^_`{|}~" 'Символы, которые будут содержаться в пароле
Dim strInput
Dim iPasslen
Dim iCharlen
Dim strPass
Dim tmp

Do
strInput = InputBox("Укажите требуемую длину пароля (от 1 до 256 символов)")
	If IsEmpty(strInput) OR strInput = "" Then 'При нажатии "Cancel", "x", "Esc", "Alt+F4" или вводе пустой строки
		WScript.Quit 'Прекратить выполнение скрипта
	Else If isNumeric(strInput) AND Len(strInput) < 4 Then 'Если строка содержит числовое значение длиной менее 4 символов
			If CStr(CInt(strInput)) = CStr(strInput) AND CInt(strInput) > 0 AND CInt(strInput) < 257 Then 'Если значение целочисленное, больше 0 и меньше 257
				iPasslen = CInt(strInput) 'То присвоить переменной длины пароля значение из окна ввода 
				Exit Do 'И продолжить выполнение
			End if
		End if
	End if
tmp = MsgBox("Длина пароля должна быть целым числом от 1 до 256", 48, "Ошибка ввода")
Loop 'Иначе вывести окно ввода заново

iCharlen = Len(CHARACTERS) 
Randomize
	For i = 1 to iPasslen
		strPass = strPass & Mid(CHARACTERS, Int(iCharlen*Rnd+1), 1) 'Добавляет к итоговой строке подстроку длиной в 1 символ из строки CHARACTERS начиная со случайного места в ней
	Next
tmp = InputBox("Ваш пароль:",,strPass) 'Вывод в окне для ввода для удобства выделения
