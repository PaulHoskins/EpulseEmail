Imports System.Data
Imports System.Xml
Imports System.IO
Imports EAGetMail
Module Module1
    Dim server, email, password, ssl, port, servtype, fromemail, subject As String

    Sub Main()
        setup()
        Console.WriteLine("EpulseEmail version 1.00.00 20/04/2017")
        Console.WriteLine("Connect To {0} {1} {2} {3} {4} {5} {6} {7} ", server, email, ssl, port, servtype, password, fromemail, subject)

        ' GetEmail()

        MsgBox("Doone")



    End Sub

    Sub setup()



        Dim xmlFile As XmlReader
        xmlFile = XmlReader.Create("c:\epulse\config.xml", New XmlReaderSettings())
        Dim ds As New DataSet
        ds.ReadXml(xmlFile)
        Dim i As Integer
        For i = 0 To ds.Tables(0).Rows.Count - 1

            server = ds.Tables(0).Rows(i).Item(0)
            ssl = ds.Tables(0).Rows(i).Item(1)
            port = ds.Tables(0).Rows(i).Item(2)
            servtype = ds.Tables(0).Rows(i).Item(3)
            email = ds.Tables(0).Rows(i).Item(4)
            password = ds.Tables(0).Rows(i).Item(5)
            fromemail = ds.Tables(0).Rows(i).Item(6).tolower
            subject = ds.Tables(0).Rows(i).Item(7).tolower

        Next

    End Sub
    Sub GetEmail()
        Dim curpath As String = Directory.GetCurrentDirectory()
        Dim mailbox As String = [String].Format("{0}\inbox", curpath)

        ' If the folder is not existed, create it.
        If Not Directory.Exists(mailbox) Then
            Directory.CreateDirectory(mailbox)
        End If

        Dim oClient As New MailClient("TryIt")

        Dim oServer As New MailServer(server, email, password, servtype)

        oServer.SSLConnection = True

        oServer.Port = 993

        Try
            oClient.Connect(oServer)
            Dim infos As MailInfo() = oClient.GetMailInfos()
            For i As Integer = 0 To infos.Length - 1
                Dim info As MailInfo = infos(i)
                Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL)

                ' Receive email from IMAP4 server
                Dim oMail As Mail = oClient.GetMail(info)

                Console.WriteLine("From: {0}", oMail.From.ToString())
                Console.WriteLine("Subject: {0}" & vbCr & vbLf, oMail.Subject)

                ' Generate an email file name based on date time.
                Dim d As System.DateTime = System.DateTime.Now
                Dim cur As New System.Globalization.CultureInfo("en-US")
                Dim sdate As String = d.ToString("yyyyMMddHHmmss", cur)
                Dim fileName As String = [String].Format("{0}\{1}{2}{3}.eml",
                    mailbox, sdate, d.Millisecond.ToString("d3"), i)



                ' Save email to local disk
                'oMail.SaveAs(fileName, True)

                ' Mark email as deleted in Hotmail Account
                'oClient.Delete(info)
            Next
            ' Quit and purge emails marked as deleted from Hotmail IMAP4 server.
            oClient.Quit()
        Catch ep As Exception
            Console.WriteLine(ep.Message)
        End Try






    End Sub
End Module
