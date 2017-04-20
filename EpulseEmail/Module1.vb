Module Module1
    Dim server, email, password As String

    Sub Main()
        setup()
        Console.WriteLine("EpulseEmail version 1.00.00 20/04/2017")
        Console.WriteLine("Connect To {0} {1}", server, email)



        MsgBox(server)

    End Sub

    Sub setup()

        server = "imap-mail.outlook.com"
        email = "paulanhoskins@outlook.com"
        password = "00pudsey"

    End Sub
End Module
