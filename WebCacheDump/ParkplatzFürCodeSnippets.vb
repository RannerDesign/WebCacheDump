Imports Microsoft.VisualBasic

Public Class ParkplatzFürCodeSnippets
    '   Absprung zum Testprogramm (muss am Anfang von Main eingabaut werden+
    Sub MainErsatz()
        If False Then
            TestProgram()
            MsgBox("Ende Testprogramm")
            Exit Sub
        End If
    End Sub


    Private Sub TestProgram()
        Dim logLineStart As String = "----" + vbTab + "Testprogramm: "
        Console.WriteLine(logLineStart + "Start Testprogramm")

        Dim A As Byte() = {1, 2, 3, 4, 31, 30, 29, 28}
        Console.WriteLine(Byte2xlString(A))

        Dim B As Byte() = {40, 50, 60, 70, 80, 90, 100, 110, 120, 130, 200}
        Console.WriteLine(Byte2xlString(B))

    End Sub

End Class
