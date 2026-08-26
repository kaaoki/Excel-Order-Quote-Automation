Attribute VB_Name = "RunOrderProcessor"
Sub RunOrderProcessor()
    Dim exePath As String
    exePath = ThisWorkbook.Path & "\OrderProcessor.exe"

    If Dir(exePath) = "" Then
        MsgBox "OrderProcessor.exe が見つかりません。" & vbCrLf & exePath, vbExclamation
        Exit Sub
    End If

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    Application.Cursor = xlWait          ' カーソルを待機中(砂時計)に変更
    Application.StatusBar = "処理中です。しばらくお待ちください..."

    ' 第3引数 True = exeの終了を待つ
    wsh.Run """" & exePath & """", 1, True

    Application.Cursor = xlDefault       ' カーソルを元に戻す
    Application.StatusBar = False
End Sub
