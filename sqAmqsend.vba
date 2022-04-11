Sub dataproc()
Dim d As String

d = "-815.0,233716,-89.6277690,-17.833333,20.000000,0,302"

httpProc (d)

End Sub

Sub httpProc(x As String)
 Dim httpObject As Object
 Set httpObject = CreateObject("MSXML2.XMLHTTP")
 sURL = "http://127.0.0.1:8080/sendMsg?msg="
 sRequest = sURL & x
 httpObject.Open "GET", sRequest, False
 httpObject.send
End Sub
