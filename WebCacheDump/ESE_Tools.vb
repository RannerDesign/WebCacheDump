'   ESE_Tools.vb
'   Auxiliary functions for the handling of ESE databases
'

Imports Microsoft.VisualBasic

Public Class ESE_Tools
    Public Shared Function RunCmd(Command As String, Arguments As String, Location As String) As Integer
        '   Execute Command as process in folder Location and wait for completion, return execution status of command
        Dim ErrOutput As String
        Dim p As Process = New Process()
        p.StartInfo.FileName = Command
        If Len(Arguments) > 0 Then p.StartInfo.Arguments = Arguments
        p.StartInfo.UseShellExecute = False
        If Len(Location) > 0 Then p.StartInfo.WorkingDirectory = Location
        p.StartInfo.RedirectStandardError = True
        p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
        Try
            p.Start()
            ErrOutput = p.StandardError.ReadToEnd()
            p.WaitForExit()
            If Len(ErrOutput) > 0 Then
                Debug.WriteLine("+++++ Fehler beim Ausführen von " + Command + " " + Arguments)
                Debug.WriteLine(ErrOutput)
            End If
            Return p.ExitCode
        Catch ex As Exception
            MsgBox(ex.Message)
            Return -1
        End Try
    End Function
End Class
