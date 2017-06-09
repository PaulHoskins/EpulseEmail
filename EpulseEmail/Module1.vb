Imports System.Data
Imports System.Xml
Imports System.IO
Imports EAGetMail
Imports System.Net
Imports System.Text
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates

Module Module1
    Dim server, email, password, ssl, port, servtype, fromemail, subject, posturl As String

    Sub Main()
        Setup()
        Console.WriteLine("EpulseEmail Version 1.00.00 24/04/2017")
        Console.WriteLine("Connect To {0} {1} {2} {3} {4} {5} {6} {7} {8}", server, email, ssl, port, servtype, password, fromemail, subject, posturl)

        GetEmail()

    End Sub

    Sub Setup()

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
            posturl = ds.Tables(0).Rows(i).Item(8)

        Next

    End Sub
    Sub GetEmail()

        Dim oClient As New MailClient("EG-B1374632949-00731-3DC3332UF1B46FBT-361C5EFV4E87E36U")
        Dim oServer As New MailServer(server, email, password, servtype)

        If ssl = "Y" Then
            oServer.SSLConnection = True
        End If

        If port <> "" Then
            oServer.Port = port
        End If

        Try
            oClient.Connect(oServer)

            Dim infos As MailInfo() = oClient.GetMailInfos()
            For i As Integer = 0 To infos.Length - 1
                Dim info As MailInfo = infos(i)
                Dim oMail As Mail = oClient.GetMail(info)
                Console.WriteLine("Index: {0}; Size: {1}; UIDL: {2}",
                        info.Index, info.Size, info.UIDL)
                Console.WriteLine("Name: {0}", oMail.From.Name())
                Console.WriteLine("Address: {0}", oMail.From.Address())
                Console.WriteLine("From: {0}", oMail.From.ToString())
                Console.WriteLine("Subject: {0}" & vbCr & vbLf, oMail.Subject)
                Console.WriteLine("Sent: {0}", oMail.SentDate)

                Dim fromcase, sublow As String

                fromcase = oMail.From.Address().ToLower
                sublow = oMail.Subject.ToLower

                If fromcase.Contains(fromemail) And sublow.Contains(subject) And sublow.Contains("problem(s)") Then

                    Console.WriteLine("match: {0} {1}", fromemail, fromcase)


                    Dim s As HttpWebRequest
                    Dim enc As UTF8Encoding
                    Dim postdata As String
                    Dim postdatabytes As Byte()
                    Console.WriteLine(posturl)

                    ServicePointManager.ServerCertificateValidationCallback = AddressOf AcceptAllCertifications

                    s = HttpWebRequest.Create(posturl)
                    enc = New System.Text.UTF8Encoding()
                    postdata = "uidl=" + info.UIDL + "&from=" + oMail.From.Address() + "&subject=" + oMail.Subject + "&body=" + oMail.TextBody + "&date=" + oMail.SentDate
                    postdatabytes = enc.GetBytes(postdata)
                    s.Method = "POST"
                    s.ContentType = "application/x-www-form-urlencoded"
                    s.ContentLength = postdatabytes.Length


                    Using stream = s.GetRequestStream()
                        stream.Write(postdatabytes, 0, postdatabytes.Length)
                    End Using

                    Dim result = s.GetResponse()
                    Dim resText As String
                    resText = CType(result, HttpWebResponse).StatusDescription
                    Console.WriteLine("Response = " & resText)
                    If resText = "OK" Then

                    End If



                    ' Mark email as deleted in Account
                    'oClient.Delete(info)

                End If

            Next
            oClient.Quit()
        Catch ep As Exception
            Console.WriteLine(ep.Message)

        End Try






    End Sub

    Private Function AcceptAllCertifications(sender As Object, certificate As X509Certificate, chain As X509Chain, sslPolicyErrors As SslPolicyErrors) As Boolean
        Return True
    End Function
End Module
