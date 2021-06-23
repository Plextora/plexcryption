set x = createObject("WScript.shell")
txt = inputbox("Type what you want to encrypt below")

txt = StrReverse(txt)
inputbox "Copy the text below", "Copy the text", (encode(""&(txt)))
function encode(s)

For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt) + 3)
coded = coded + (newtxt)
Next
encode = coded
End function