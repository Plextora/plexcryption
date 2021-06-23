set x = createObject("WScript.shell")
txt = inputbox("Paste in what you want to decrypt here:")
txt = StrReverse(txt)
x.Run"notepad"
wscript.sleep "1000"
x.sendkeys("Here's your decrypted message: " + (encode(""&(txt))))
function encode(s)

For i = 1 To Len(s)
newtxt = Mid(s, i, 1)
newtxt = Chr(Asc(newtxt) - 3)
coded = coded + (newtxt)
Next
encode = coded
End function