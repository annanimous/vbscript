Const PASSLENGTH = 16 'Задаём длину пароля
Const CHARACTERS = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!#$%&()*+,-./:;<=>?@[\]^_`{|}~" 'Символы, которые будут содержаться в пароле
Dim nCharlen
Dim strPass
Dim tmp

nCharlen = Len(CHARACTERS) 
Randomize
	For i = 1 to PASSLENGTH 
		strPass = strPass & Mid(CHARACTERS, Int(nCharlen*Rnd+1), 1) 'Добавляет к итоговой строке подстроку длиной в 1 символ из строки CHARACTERS начиная со случайного места в ней
	Next
tmp = InputBox("Ваш пароль:",,strPass) 'Вывод в окне для ввода для удобства выделения
