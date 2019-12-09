Dim nPasslen
nPasslen = 16 'Задаём длину пароля
Const CHARACTERS = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!#$%&()*+,-./:;<=>?@[\]^_`{|}~" 'Символы, которые будут содержаться в пароле
Dim nCharlen
nCharlen = Len(CHARACTERS) 
Dim strPass
Randomize
For i = 1 to nPasslen 
strPass = strPass & Mid(CHARACTERS, Int(nCharlen*Rnd+1), 1) 'Добавляет к итоговой строке подстроку длиной в 1 символ из строки CHARACTERS начиная со случайного места в ней
Next
Dim tmp
tmp = InputBox("Ваш пароль:",,strPass)
