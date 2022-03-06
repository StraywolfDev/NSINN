x = inputbox("Calculation (spaces between number and operator):", "NSINN Calculator")
y = Split(x)
if not isnumeric (y(0)) then 
wscript.quit 
end if 
if not isnumeric (y(2)) then 
wscript.quit 
end if 
if y(1) = "+" then
z = int(y(0)) + int(y(2))
msgbox(z)
end if
if y(1) = "-" then
z = int(y(0)) - int(y(2))
msgbox(z)
end if
if y(1) = "*" then
z = int(y(0)) * int(y(2))
msgbox(z)
end if
if y(1) = "/" then
z = int(y(0)) / int(y(2))
msgbox(z)
end if
if y(1) = "%" then
z = int(y(0)) MOD int(y(2))
msgbox(z)
end if