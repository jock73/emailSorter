Imports System.IO
Imports System.Configuration.ConfigurationManager
Imports System.Environment



Module logslogs


    Dim filen As String = AppSettings("logfile").ToString() + CStr(Now.Year.ToString()) + "-" + CStr(Now.Month.ToString()) + "-" + Now.Day.ToString() & ".xml"
    Dim writer As StreamWriter
    Dim datenow As String = CStr(Date.Now.ToString())

    Sub logfile()

        If File.Exists(filen) Then

            Dim aa As String = File.ReadAllText(filen)
            ''   System.IO.File.WriteAllText(filen, aa.Replace("</Root>", ""))

            '' write xml headers for pre exsisting xml
            writer = New StreamWriter(filen, True)
            Write("<Session>")
            Write("<Time>" + DateTime.Now.ToString("ddd, MMM yyyy hh:mm:ss") + "</Time>")

        Else

            writer = New StreamWriter(filen, False)


            Write("<Session>")
            Write("<Time>" + DateTime.Now.ToString("ddd, MMM yyyy hh:mm:ss") + "</Time>")


        End If
    End Sub



    Sub Write(ByVal Content As String)

        Console.Write(Content & NewLine)
        writer.WriteLine(Content)
    End Sub


    Sub Close()

        Write("</Session>")

        writer.Close()
        writer.Dispose()
    End Sub


End Module
