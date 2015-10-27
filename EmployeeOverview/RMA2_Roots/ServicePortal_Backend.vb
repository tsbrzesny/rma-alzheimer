Option Infer On

Imports System.Collections.Generic
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml
Imports RMA2_Roots.RMA2D
Imports RMA_Roots



Public Module RMA2B

    ' Service Portal Backend Web-Services

    Public backendRoot As String = AppConfig.GetItem("BackendRoot")
    Public backendAPIKey As String = AppConfig.GetItem("BackendApiKey")



    '------------------------------------------------------------------
    ' DocFlo functions


    Public Function L2Doc_GetNodesOfContact(ByVal contactid As Integer) As List(Of L2_Node)

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:get_l2nodes

        ' Note: node.mandant_id is not delivered by tis call.. get it from l2_nodes, if needed


        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "contact_id", .Value = contactid})
        p.Add(New DictionaryEntry With {.Key = "api_key", .Value = backendAPIKey})

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & "l2doc/nodes"
        Dim httpStatus = HTTPGet_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Throw New Exception(String.Format("RMA-Backend.L2Doc_GetNodesOfContact: HTTP-Code {0}", httpStatus))
        End If

        ' parse xml answer
        Dim nodeList = New List(Of L2_Node)
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            For Each _node In xDoc...<l2_node>
                Dim node As New L2_Node With {
                        .id = _node.@id,
                        .display_name = _node.@display_name,
                        .node_type = _node.@type_id
                    }
                nodeList.Add(node)
            Next

        Catch ex As Exception
        End Try

        Return nodeList
    End Function



    Public Function L2Doc_Checkout(ByVal doc_id As Integer, ByVal node_id As Integer, ByVal byWhom As String,
                                   Optional ByVal app_id As String = Nothing,
                                   Optional ByRef token As String = Nothing, Optional ByRef edgeList As List(Of L2_Edge) = Nothing) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:get_l2checkout
        ' return Nothing on success, an error message string else

        token = Nothing
        edgeList = Nothing

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "node_id", .Value = node_id})
        p.Add(New DictionaryEntry With {.Key = "by_whom", .Value = byWhom})
        If app_id IsNot Nothing Then
            p.Add(New DictionaryEntry With {.Key = "app_id", .Value = app_id})
        End If
        p.Add(New DictionaryEntry With {.Key = "api_key", .Value = backendAPIKey})

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & String.Format("l2doc/{0}/checkout", doc_id)
        Dim httpStatus = HTTPGet_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Return String.Format("RMA-Backend.L2Doc_Checkout: HTTP-Code {0}", httpStatus)
        End If

        ' parse xml answer
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            If xDoc...<l2_status>.@status <> "SUCCESS" Then
                ' call technically successful, but with DocFlo error: return error message
                Return xDoc...<l2_status>.@error_message
            End If

            ' extract information
            token = xDoc...<l2_status>.@token

            edgeList = New List(Of L2_Edge)
            For Each _edge In xDoc...<edge>
                Dim edge As New L2_Edge With {
                        .id = _edge.@id,
                        .name = _edge.@name,
                        .edge_type = _edge.@edge_type_id,
                        .src_node = _edge.@src_node_id,
                        .target_node = _edge.@trg_node_id
                    }
                edgeList.Add(edge)
            Next

        Catch ex As Exception
            Return "RMA-Backend.L2Doc_Checkout: Fehlerhaftes XML."
        End Try

        Return Nothing
    End Function


    Public Function L2Doc_Prepare(ByVal doc_id As Integer, ByVal node_id As Integer, ByVal byWhom As String, ByVal checkout_token As String,
                                  Optional ByVal app_id As String = Nothing) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:get_l2prepare

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "node_id", .Value = node_id})
        p.Add(New DictionaryEntry With {.Key = "token", .Value = checkout_token})
        p.Add(New DictionaryEntry With {.Key = "by_whom", .Value = byWhom})
        If app_id IsNot Nothing Then
            p.Add(New DictionaryEntry With {.Key = "app_id", .Value = app_id})
        End If
        p.Add(New DictionaryEntry With {.Key = "api_key", .Value = backendAPIKey})

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & String.Format("l2doc/{0}/prepare", doc_id)
        Dim httpStatus = HTTPGet_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Throw New Exception(String.Format("RMA-Backend.L2Doc_Prepare: HTTP-Code {0}", httpStatus))
        End If

        ' parse xml answer
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            If xDoc...<l2_status>.@status <> "SUCCESS" Then
                ' call technically successful, but with DocFlo error: return error message
                Return xDoc...<l2_status>.@error_message
            End If

        Catch ex As Exception
            Return "RMA-Backend.L2Doc_Prepare: Fehlerhaftes XML."
        End Try

        Return Nothing
    End Function


    Public Function L2Doc_Finalize(ByVal doc_id As Integer, ByVal node_id As Integer, ByVal byWhom As String, ByVal checkout_token As String,
                                   ByVal action As String, Optional ByVal target_node_id As Integer = Integer.MinValue,
                                   Optional ByVal app_id As String = Nothing) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:get_l2finalize

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "node_id", .Value = node_id})
        p.Add(New DictionaryEntry With {.Key = "token", .Value = checkout_token})
        p.Add(New DictionaryEntry With {.Key = "by_whom", .Value = byWhom})
        p.Add(New DictionaryEntry With {.Key = "action", .Value = action})
        If target_node_id <> Integer.MinValue Then
            p.Add(New DictionaryEntry With {.Key = "trg_node_id", .Value = target_node_id})
        End If
        If app_id IsNot Nothing Then
            p.Add(New DictionaryEntry With {.Key = "app_id", .Value = app_id})
        End If
        p.Add(New DictionaryEntry With {.Key = "api_key", .Value = backendAPIKey})

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & String.Format("l2doc/{0}/finalize", doc_id)
        Dim httpStatus = HTTPGet_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Throw New Exception(String.Format("RMA-Backend.L2Doc_Finalize: HTTP-Code {0}", httpStatus))
        End If

        ' parse xml answer
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            If xDoc...<l2_status>.@status <> "SUCCESS" Then
                ' call technically successful, but with DocFlo error: return error message
                Return xDoc...<l2_status>.@error_message
            End If

        Catch ex As Exception
            Return "RMA-Backend.L2Doc_Finalize: Fehlerhaftes XML."
        End Try

        Return Nothing
    End Function




    '------------------------------------------------------------------
    ' various functions


    Public Function CreateLocalVendor(ByVal app_id As String, ByVal mandant As String, ByRef vRec As RMA2D.SLVCRecord) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:post_vendors

        ' check if there is a local vendor with identical name (ignore case)
        RMA2D.SetNewSQLLedgerMandant(mandant)
        Dim vendorName = StrLength(vRec.name, 64)
        Dim identicalLV = RMA2D.slVendors.Find(Function(_lv) _lv.name.ToLower = vendorName.ToLower)
        If identicalLV IsNot Nothing Then
            vRec = identicalLV
            Return Nothing      ' ok
        End If

        Dim chartT = RMA2D.GetLedgerChart()

        ' standard payment account
        Dim payment_accno As String = Nothing
        Dim appaidAccounts = (From _ch In chartT Where Regex.IsMatch(_ch.link, "AP_paid\b") And _ch.accno >= "1020" Order By _ch.accno).ToList()
        If appaidAccounts.Count() = 0 Then
            Return "Kein Zahlungskonto 'AP_paid' gefunden! Kann Lieferant nicht anlegen ohne."
        End If
        If (appaidAccounts.Exists(Function(_ch) _ch.accno = "1020")) Then
            payment_accno = "1020"      ' CH main account
        ElseIf (appaidAccounts.Exists(Function(_ch) _ch.accno = "1800")) Then
            payment_accno = "1800"      ' DE main account
        Else
            payment_accno = appaidAccounts(0).accno
        End If

        ' find standard vendor account:
        ' - find all 'AP' accounts in the mandant's chart
        ' - select account 2000 or the lower account number
        ' - error if no AP account is found
        Dim apAccounts = From _ch In chartT Where Regex.IsMatch(_ch.link, "AP\b")
        If apAccounts.Count = 0 Then
            Return "Kein 'AP'-Konto gefunden!"
        End If

        Dim arap_accno = "2000"
        If Not apAccounts.Any(Function(_ch) _ch.accno = "2000") Then
            apAccounts = apAccounts.OrderBy(Function(_ch1) Val(_ch1.accno))
            arap_accno = apAccounts(0).accno
        End If

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "app_id", .Value = RMA2S.StringNothing2Empty(app_id)})

        p.Add(New DictionaryEntry With {.Key = "name", .Value = vendorName})
        p.Add(New DictionaryEntry With {.Key = "arap_accno", .Value = arap_accno})
        p.Add(New DictionaryEntry With {.Key = "payment_accno", .Value = payment_accno})
        ' use the defaults for the rest of the parameters

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & String.Format("clients/{0}/vendors", mandant)
        Dim httpStatus = HTTPPost_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Return String.Format("RMA-Backend.CreateLocalVendor: HTTP-Code {0}", httpStatus)
        End If

        ' parse xml answer to find the id of the new vendor record
        Dim newId As String = Nothing
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            newId = xDoc...<id>.Value
        Catch ex As Exception
        End Try

        If newId Is Nothing Then
            Return "RMA-Backend.CreateLocalVendor: New ID is NIL."
        End If

        ' reload vendor table and update vRec parameter
        RMA2D.ReloadLedgerVendorTable()
        vRec = RMA2D.slVendors.Find(Function(_lv) _lv.id = newId)
        If vRec Is Nothing Then
            Return "RMA-Backend.CreateLocalVendor: New record is NIL."
        End If

        Return Nothing      ' ok
    End Function


    Public Function CreateLocalCustomer(ByVal app_id As String, ByVal mandant As String, ByRef cRec As RMA2D.SLVCRecord) As String

        ' api doku:  https://app.runmyaccounts.com/doku/doku.php?id=api:api_post_customer

        ' check if there is a local customer with identical name (ignore case)
        RMA2D.SetNewSQLLedgerMandant(mandant)
        Dim customerName = StrLength(cRec.name, 64)
        Dim identicalLC = RMA2D.slCustomers.Find(Function(_lc) _lc.name.ToLower = customerName.ToLower)
        If identicalLC IsNot Nothing Then
            cRec = identicalLC
            Return Nothing      ' ok
        End If

        ' locate standard payment account 1020
        Dim payment_accno = 1020
        Dim chartT = RMA2D.GetLedgerChart()
        If Not chartT.Exists(Function(_ch) _ch.accno = payment_accno) Then
            Return String.Format("Fehler beim Anlegen eines lokalen Kunden: Zahlungskonto {0} nicht gefunden!", payment_accno)
        End If

        ' find standard vendor account:
        ' - find all 'AR' accounts in the mandant's chart
        ' - select account 1100 or the lower account number
        ' - error if no AR account is found
        Dim arAccounts = From _ch In chartT Where Regex.IsMatch(_ch.link, "AR\b")
        If arAccounts.Count = 0 Then
            Return "Fehler beim Anlegen eines lokalen Kunden: Kein 'AR'-Konto gefunden!"
        End If

        Dim arap_accno = 1100
        If Not arAccounts.Any(Function(_ch) _ch.accno = arap_accno) Then
            arAccounts = arAccounts.OrderBy(Function(_ch1) Val(_ch1.accno))
            arap_accno = arAccounts(0).accno
        End If

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "app_id", .Value = RMA2S.StringNothing2Empty(app_id)})

        p.Add(New DictionaryEntry With {.Key = "name", .Value = customerName})
        p.Add(New DictionaryEntry With {.Key = "arap_accno", .Value = arap_accno})
        p.Add(New DictionaryEntry With {.Key = "payment_accno", .Value = payment_accno})
        ' use the defaults for the rest of the parameters

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & String.Format("clients/{0}/customers?api_key={1}", mandant, backendAPIKey)
        Dim httpStatus = HTTPPost_XML(uri, xmlAnswer, Nothing, "customer", p)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Return String.Format("RMA-Backend.CreateLocalCustomer: HTTP-Code {0}", httpStatus)
        End If

        ' reload customer table and update cRec parameter
        RMA2D.ReloadLedgerCustomerTable()
        cRec = RMA2D.slCustomers.Find(Function(_lv) _lv.name = customerName)
        If cRec Is Nothing Then
            Return "RMA-Backend.CreateLocalCustomer: New record is NIL."
        End If

        Return Nothing      ' ok
    End Function


    Public Function GetReceiptStackInfo(ByVal filename As String, ByRef sender As String, ByRef subject As String, ByRef receivedDate As Date) As String

        Dim xmlAnswer As Stream = Nothing

        ' backend needs a '.', and the vb framework doesn't allow a . as the last char..
        filename &= ".wtfse"

        Dim uri = backendRoot & "receipt_stack/" & filename
        Dim httpStatus = HTTPGet_URLencoded(uri, Nothing, xmlAnswer)

        sender = Nothing
        subject = Nothing
        receivedDate = Nothing
        Try
            Dim xDoc = XDocument.Load(xmlAnswer)
            With xDoc...<receipt>
                sender = .@fromStr
                subject = .@subject
                receivedDate = XmlConvert.ToDateTime(.@receivedDatetime, Xml.XmlDateTimeSerializationMode.Local)
            End With
        Catch ex As Exception
        End Try

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Return String.Format("RMA-Backend.GetReceiptStackInfo: HTTP-Code {0}", httpStatus)
        End If

        Return Nothing      ' ok
    End Function


    Public Function FireLedgerCmd(ByVal app_id As String, ByVal doc_id As String,
                                  ByVal cmd As String, ByVal parameters As List(Of DictionaryEntry), ByRef response As String) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:post_script

        If RMA2S.CheckDEBUGState Then
            MsgBox(cmd & "-Buchung", MsgBoxStyle.OkOnly, "FireLedgerCmd")
            response = "Ok!"
            Return Nothing
        End If

        app_id = RMA2S.StringNothing2Empty(app_id)
        doc_id = RMA2S.StringNothing2Empty(doc_id)
        Dim uri = backendRoot & String.Format("script/{0}?app_id={1}&doc_id={2}", cmd, app_id, doc_id)

        Dim xmlAnswer As Stream = Nothing
        Dim httpStatus = HTTPPost_URLencoded(uri, parameters, xmlAnswer)

        ' the response is not xml.. :(
        Try
            Dim reader As New StreamReader(xmlAnswer)
            response = reader.ReadToEnd()

            ' unify the success response strings..
            Select Case response
                Case "Buchung getätigt!", "Zahlung gebucht!"
                    response = "Ok!"
            End Select
        Catch ex As Exception
        End Try

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            response = StringNothing2Empty(response)
            Return String.Format("RMA-Backend.FireLedgerCmd: HTTP-Code {0}, '{1}'", httpStatus, response)
        End If

        Return Nothing      ' ok
    End Function


    Public Function PostPayment(ByVal app_id As String, ByVal doc_id As String,
                                ByVal parameters As List(Of DictionaryEntry)) As String

        ' api doku: http://192.168.10.5/dokuwiki/doku.php?id=it:restful:post_zahlungen

        app_id = RMA2S.StringNothing2Empty(app_id)
        doc_id = RMA2S.StringNothing2Empty(doc_id)
        Dim uri = backendRoot & String.Format("zahlungen?api_key={0}&app_id={1}&doc_id={2}", backendAPIKey, app_id, doc_id)

        Dim xmlAnswer As Stream = Nothing
        Dim httpStatus = HTTPPost_URLencoded(uri, parameters, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            ' the response is not xml.. :(
            Dim plainResponse As String = ""
            Try
                Dim reader As New StreamReader(xmlAnswer)
                plainResponse = reader.ReadToEnd()
            Catch ex As Exception
            End Try

            Return String.Format("RMA-Backend.PostPayment: HTTP-Code {0}, '{1}'", httpStatus, plainResponse)
        End If

        Return Nothing      ' ok
    End Function


    Public Function PostBilling(ByVal parameters As List(Of DictionaryEntry)) As String

        ' api doku: none. ask Nils.

        Dim uri = backendRoot & "billing?api_key=" & backendAPIKey

        Dim xmlAnswer As Stream = Nothing
        Dim httpStatus = HTTPPost_URLencoded(uri, parameters, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            ' the response is not xml.. :(
            Dim plainResponse As String = ""
            Try
                Dim reader As New StreamReader(xmlAnswer)
                plainResponse = reader.ReadToEnd()
            Catch ex As Exception
            End Try

            Return String.Format("RMA-Backend.PostBilling: HTTP-Code {0}, '{1}'", httpStatus, plainResponse)
        End If

        Return Nothing      ' ok
    End Function



    Public Function PostInvoice(ByVal mandant As StammRecord, ByVal invoiceRequest As XDocument) As String

        ' api doku: http://www.runmyaccounts.ch/buchhaltungs-hilfe/doku.php/api/ua_f1_b1

        Dim uri = String.Format("{0}clients/{1}/invoices?api_key={2}", backendRoot, mandant.mandant, backendAPIKey)

        Dim xmlAnswer As Stream = Nothing
        Dim httpStatus = HTTPPost_XML(uri, xmlAnswer, xDoc:=invoiceRequest)

        If Not httpStatus.ToString.StartsWith("2") Then
            Dim errorStr = "Fehler beim Anlegen der Rechnung"
            Try
                Dim reader As New StreamReader(xmlAnswer)
                Dim r = XDocument.Parse(reader.ReadToEnd)
                errorStr = r...<error>.Value
            Catch ex As Exception
            End Try

            Return String.Format("RMA-Backend.PostInvoice: HTTP-Code {0}, '{1}'", httpStatus, errorStr)
        End If

        Return Nothing      ' ok
    End Function



    Public Function PostMessage2ServicePortal(ByVal recipient As Contact, ByVal from As String, ByVal mandant As String,
                                              ByVal subject As String, ByVal body As String) As String
        ' Nils online help:
        ' Http POST mit Content-Type application/x-www-form-urlencoded. Folgende Parameter müssen gesetzt werden:
        ' employee: Name des Mitarbeiters gemäss Tabelle login Stammdaten (Bsp. Patrick Vögeli)
        ' from:   Absender, varchar(1024)
        ' subject: Betreff der Nachricht, varchar(1024)
        ' body: Nachricht, beliebiger Text
        ' client_name: Mandantenname (optional)
        ' status: Fest(auf) 'message' gesetzt
        ' checksum: Checksum der Nachricht. Entscheidet, ob die Nachricht als Duplikat (aka Dublikat) erkannt wird oder nicht.
        '
        ' Beispiel mit einem seriösen Shell Tool:
        ' curl -d "employee=Nils+Samuelsson&from=ESR+Tool&subject=Doppelte+Zahlung&checksum=991231239017&body=&status=message" http://localhost:8080/rma-backend/latest/receipt_stack/add?api_key=h07vryj0Ta4MzwRJQn9wxKOKU8Sm4PxU

        Dim p As New List(Of DictionaryEntry)
        p.Add(New DictionaryEntry With {.Key = "employee", .Value = String.Format("{0} {1}", recipient.firstname, recipient.name)})
        p.Add(New DictionaryEntry With {.Key = "from", .Value = "ESR-Verarbeitung"})
        p.Add(New DictionaryEntry With {.Key = "subject", .Value = subject})
        p.Add(New DictionaryEntry With {.Key = "body", .Value = body})
        p.Add(New DictionaryEntry With {.Key = "client_name", .Value = mandant})
        p.Add(New DictionaryEntry With {.Key = "status", .Value = "message"})
        p.Add(New DictionaryEntry With {.Key = "checksum", .Value = Now.Ticks})
        ' use the defaults for the rest of the parameters

        Dim xmlAnswer As Stream = Nothing
        Dim uri = backendRoot & "receipt_stack/add?api_key=" & backendAPIKey
        Dim httpStatus = HTTPPost_URLencoded(uri, p, xmlAnswer)

        If Not httpStatus.ToString.StartsWith("2") Then
            ' http error..
            Return String.Format("RMA-Backend.PostMessage2ServicePortal: HTTP-Code {0}", httpStatus)
        End If

        Return Nothing      ' ok
    End Function





    '--------------------------------------
    '
    ' base interfaces

    Public Function HTTPPost_URLencoded(ByVal endPoint As String, ByVal parameters As List(Of DictionaryEntry),
                                        Optional ByRef answerStream As Stream = Nothing) As Integer
        ' returns the HTTP-code and optionally the embedded answer contents

        ' create Url-encoded, UTF8-encoded version of the parameter list
        Dim parameterList = New List(Of String)
        If parameters IsNot Nothing Then
            For Each parameter In parameters
                parameterList.Add(String.Format("{0}={1}", parameter.Key, RMA2S.UrlEncode(parameter.Value)))
            Next
        End If
        Dim paramterStr = RMA2S.EasyJoin("&", parameterList.ToArray)
        Dim pBlock = Encoding.UTF8.GetBytes(paramterStr)

        Dim response As HttpWebResponse = Nothing
        Try
            ' create a WebRequest object and set it up for Http-POST
            Dim request = WebRequest.Create(endPoint)
            request.Method = "POST"
            request.ContentLength = pBlock.Length
            request.ContentType = "application/x-www-form-urlencoded; charset=""utf-8"""
            request.GetRequestStream().Write(pBlock, 0, pBlock.Length)

            ' execute request & catch errors..
            response = request.GetResponse()
        Catch ex As Exception
        End Try
        If response Is Nothing Then
            ' the web-backend failed to return a valid answer - return a dummy HTTP error
            Return 503      ' service not available
        End If

        Dim httpStatus As Integer = response.StatusCode
        If httpStatus.ToString.StartsWith("2") Then
            Try
                answerStream = response.GetResponseStream
            Catch ex As Exception
                ' leave answerStream empty
            End Try
        End If

        Return httpStatus
    End Function


    Public Function HTTPPost_XML(ByVal endPoint As String, ByRef answerStream As Stream,
                                 Optional ByVal xDoc As XDocument = Nothing,
                                 Optional ByVal containerName As String = Nothing, Optional ByVal parameters As List(Of DictionaryEntry) = Nothing) As Integer
        ' returns the HTTP-code and optionally the embedded answer contents
        ' pass EITHER
        '   xDoc OR
        ' containerName, parameters


        ' create xml version of the parameter list
        Dim ps = New MemoryStream
        Dim xtw As XmlTextWriter = Nothing
        If xDoc Is Nothing Then
            ' containerName, parameters expected
            xtw = New XmlTextWriter(ps, System.Text.Encoding.UTF8)
            With xtw
                .WriteStartDocument(True)
                .WriteStartElement(containerName)
                For Each parameter In parameters
                    .WriteStartElement(parameter.Key)
                    .WriteString(parameter.Value)
                    .WriteEndElement()
                Next
                .WriteEndElement()
                .WriteEndDocument()
                .Flush()
            End With

        Else
            ' convert xDoc to a stream
            Dim xws = New XmlWriterSettings()
            xws.OmitXmlDeclaration = True
            xws.Indent = True
            xws.Encoding = System.Text.Encoding.UTF8

            Using xw = XmlWriter.Create(ps, xws)
                xDoc.WriteTo(xw)
            End Using
        End If

        Dim response As HttpWebResponse = Nothing
        Try
            ' create a WebRequest object and set it up for Http-POST
            Dim request = WebRequest.Create(endPoint)
            request.Method = "POST"
            request.ContentType = "application/xml"
            request.ContentLength = ps.Length

            Dim rs = request.GetRequestStream()
            Dim bufSize As Integer = 10000
            Dim buf(bufSize - 1) As Byte
            ps.Seek(0, SeekOrigin.Begin)
            Do
                bufSize = ps.Read(buf, 0, bufSize)
                rs.Write(buf, 0, bufSize)
            Loop Until bufSize < buf.Length

            ' execute request & catch errors..
            response = request.GetResponse()
        Catch ex As Exception
        End Try

        If xtw IsNot Nothing Then
            xtw.Close()
        End If
        ps.Close()

        If response Is Nothing Then
            ' the web-backend failed to return a valid answer - return a dummy HTTP error
            Return 503      ' service not available
        End If

        Dim httpStatus As Integer = response.StatusCode
        If httpStatus.ToString.StartsWith("2") Then
            Try
                answerStream = response.GetResponseStream
            Catch ex As Exception
                ' leave answerStream empty
            End Try
        End If

        Return httpStatus
    End Function


    Public Function HTTPGet_URLencoded(ByVal endPoint As String, ByVal parameters As List(Of DictionaryEntry),
                                       Optional ByRef answerStream As Stream = Nothing) As Integer
        ' returns the HTTP-code and optionally the embedded answer contents

        ' create Url-encoded, UTF8-encoded version of the parameter list
        If parameters IsNot Nothing Then
            Dim parameterList = New List(Of String)
            For Each parameter In parameters
                parameterList.Add(String.Format("{0}={1}", parameter.Key, RMA2S.UrlEncode(parameter.Value)))
            Next
            Dim paramterStr = RMA2S.EasyJoin("&", parameterList.ToArray)
            endPoint &= "?" & paramterStr
        End If

        Dim response As HttpWebResponse = Nothing
        Try
            ' create a WebRequest object and set it up for Http-POST
            Dim request = WebRequest.Create(endPoint)
            request.Method = "GET"

            ' execute request & catch errors..
            response = request.GetResponse()
        Catch ex As Exception
        End Try
        If response Is Nothing Then
            ' the web-backend failed to return a valid answer - return a dummy HTTP error
            Return 503      ' service not available
        End If

        Dim httpStatus As Integer = response.StatusCode
        If httpStatus.ToString.StartsWith("2") Then
            Try
                answerStream = response.GetResponseStream
            Catch ex As Exception
                ' leave answerStream empty
            End Try
        End If

        Return httpStatus
    End Function

End Module
