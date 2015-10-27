
Option Infer On

Imports System.Text.RegularExpressions
Imports System.Math
Imports System.Drawing
Imports System.Linq
Imports Npgsql
Imports RMA_Roots
Imports System.IO


'=============================================================================================================
'
'       Interfaces
'
'

Public Interface IExchangeRate
    ' Gets buy/sell rate --> CHF
    Function GetExchangeRate_toCHF(ByVal currency As String, ByVal xDate As Date, ByVal sellRate As Boolean, ByRef rate As Double) As Boolean

    ' Gets rate to convert fromCurr --> toCurr: toAmount = fromAmount * rate
    Function GetExchangeRate(ByVal fromCurr As String, ByVal toCurr As String, ByVal xDate As Date, ByRef rate As Double) As Boolean
End Interface




Public Class RMA2D



    '-----------------------------------------------------------
    '
    '       Value processors for PostgreSQL & XML
    '
    '


    ' PostgreSQL value writers

    Public Shared Function PSQL_String(ByVal str As String, Optional ByVal nothingAsNULL As Boolean = True, Optional ByVal emptyAsNULL As Boolean = True,
                                       Optional ByVal maxLength As Integer = -1) As String
        ' prepares an input string to be used as a PostgreSQL string parameter
        ' - Nothing --> empty string / NULL
        ' - replace ' with ''
        ' - enclose in single '
        If str Is Nothing Then
            If nothingAsNULL Then
                Return "NULL"
            End If
            str = ""

        Else
            str.Trim()
            If str.Length = 0 And emptyAsNULL Then
                Return "NULL"
            End If
            If maxLength > 0 Then
                str = RMA2S.StrLength(str, maxLength)
            End If
        End If

        str = str.Replace("'", "''")
        str = str.Replace("\", "\\")
        Return "E'" & str & "'"
    End Function


    Public Shared Function PSQL_Date(ByVal d As Date) As String
        If d = Nothing Then         ' = is correct.. don't IS this!
            Return "NULL"
        End If
        Return PSQL_String(d.ToString("yyyy-MM-dd"))
    End Function


    Public Shared Function PSQL_Timestamp(ByVal d As Date) As String
        If d = Nothing Then         ' = is correct.. din't IS this!
            Return "NULL"
        End If
        Return PSQL_String(d.ToString("yyyy-MM-dd HH:mm:ss"))
    End Function


    Public Shared Function PSQL_Double(ByVal d As Double, Optional ByVal digits As Integer = -1) As String
        If Double.IsNaN(d) Then
            Return "NULL"
        End If
        If digits > 0 Then
            d = Round(d, digits)
        End If
        Return d.ToString
    End Function


    Public Shared Function PSQL_Int(ByVal i As Integer) As String
        If i = Integer.MinValue Then
            Return "NULL"
        End If
        Return i.ToString
    End Function


    Public Shared Function PSQL_BooleanNIL(b As Boolean?) As String
        If Not b.HasValue Then
            Return "NULL"
        ElseIf b.Value Then
            Return "TRUE"
        End If
        Return "FALSE"
    End Function




    ' XML value writers

    Public Shared Function XML_String(ByVal str As String, Optional ByVal maxLength As Integer = -1) As String
        ' prepares an input string to be used as a XML string parameter
        ' - Nothing --> empty string
        str = StringNothing2Empty(str)
        If maxLength > 0 Then
            str = RMA2S.StrLength(str, maxLength)
        End If
        Return str
    End Function


    Public Shared Function XML_Date(ByVal d As Date) As String
        ' returns the UTC date/time format
        ' or '' for d = Nothing
        If d = Nothing Then         ' = is correct.. don't IS this!
            Return ""
        End If
        Return d.ToString("s")
    End Function


    Public Shared Function XML_Double(ByVal d As Double, Optional ByVal digits As Integer = -1) As String
        If Double.IsNaN(d) Then
            Return ""
        End If
        If digits > 0 Then
            d = Round(d, digits)
        End If
        Return d.ToString
    End Function


    Public Shared Function XML_Int(Optional ByVal i As Integer = Integer.MinValue) As String
        If i = Integer.MinValue Then
            Return ""
        End If
        Return i.ToString
    End Function



    ' XML value readers

    Public Shared Function XML_2Boolean(ByVal value As String, Optional ByVal def As Boolean = False) As Boolean
        If value IsNot Nothing Then
            value = value.ToUpper
            If value = "TRUE" Or value = "1" Then
                Return True
            ElseIf value = "FALSE" Or value = "0" Then
                Return False
            End If
            Return def
        End If
    End Function

    Public Shared Function XML_2Boolean(ByVal xml As XDocument, ByVal nodeName As String, Optional ByVal def As Boolean = False) As Boolean
        Dim node = xml.Descendants.Where(Function(_d) _d.Name = nodeName).FirstOrDefault
        If node Is Nothing Then
            Return def
        Else
            Return XML_2Boolean(node.Value, def)
        End If
    End Function


    Public Shared Function XML_2Integer(ByVal value As String, Optional ByVal def As Integer = Integer.MinValue) As Integer
        Dim i As Integer = def
        If value IsNot Nothing Then
            Integer.TryParse(value, i)
        End If
        Return i
    End Function

    Public Shared Function XML_2Integer(ByVal xml As XDocument, ByVal nodeName As String, Optional ByVal def As Integer = Integer.MinValue) As Integer
        Dim node = xml.Descendants.Where(Function(_d) _d.Name = nodeName).FirstOrDefault
        If node Is Nothing Then
            Return def
        Else
            Return XML_2Integer(node.Value, def)
        End If
    End Function


    Public Shared Function XML_2Double(ByVal value As String, Optional ByVal def As Double = Double.NaN) As Double
        Dim d As Double = def
        If value IsNot Nothing Then
            Double.TryParse(value, d)
        End If
        Return d
    End Function

    Public Shared Function XML_2Double(ByVal xml As XDocument, ByVal nodeName As String, Optional ByVal def As Double = Double.NaN) As Double
        Dim node = xml.Descendants.Where(Function(_d) _d.Name = nodeName).FirstOrDefault
        If node Is Nothing Then
            Return def
        Else
            Return XML_2Double(node.Value, def)
        End If
    End Function


    Public Shared Function XML_2Date(ByVal value As String) As Date
        Dim d As Date = Nothing
        If value IsNot Nothing Then
            AnalyzeDateStr(value, d)
        End If
        Return d
    End Function

    Public Shared Function XML_2Date(ByVal xml As XDocument, ByVal nodeName As String) As Date
        Dim node = xml.Descendants.Where(Function(_d) _d.Name = nodeName).FirstOrDefault
        If node Is Nothing Then
            Return Nothing
        Else
            Return XML_2Date(node.Value)
        End If
    End Function


    Public Shared Function XML_2Record(ByVal xml As XDocument, ByVal nodeName As String) As Object
        Dim node = xml.Descendants.Where(Function(_d) _d.Name = nodeName).FirstOrDefault
        If node Is Nothing Then
            Return Nothing
        Else
            Return GetRecordByXMLId(node.Value)
        End If
    End Function





    '=============================================================================================================
    '
    '       DB reload management
    '
    '

    Private Shared currentDBToken As Integer = 0
    Private Shared tokenList As New Dictionary(Of String, Integer)

    Public Shared Sub ReloadAll()
        ' generate new token; all subsequent calls to MustReload() will return True
        currentDBToken += 1
        tokenList.Clear()
    End Sub

    Public Shared Sub Reload(ByVal id As String)
        If tokenList.ContainsKey(id) Then
            tokenList.Remove(id)
        End If
    End Sub

    Public Shared Function MustReload(ByVal id As String) As Boolean
        ' returns True, if id does not exist in the tokenList or if the token value has changed
        If Not tokenList.ContainsKey(id) OrElse tokenList(id) <> currentDBToken Then
            tokenList(id) = currentDBToken
            Return True
        End If
        Return False
    End Function



    '=============================================================================================================
    '
    '       Generic Record Loader
    '
    ' Some records are able to return a 'xml id' (function GetXMLId()).
    ' This loader returns the corresponding record for the given id.

    Public Shared Function GetRecordByXMLId(ByVal xmlId As String) As Object
        ' Note: to get mandant related records, THE RIGHT MANDANT MUST BE SET! when calling this function.

        If xmlId Is Nothing Then
            Return Nothing
        End If

        Dim m = Regex.Match(xmlId, "(\w+)¬(\d+)")
        If Not m.Success Then
            Return Nothing
        End If

        Dim tableId = m.Groups(1).Value
        Dim recId = m.Groups(2).Value

        ' get generic data
        Select Case tableId
            Case "Mandant"
                ' global vendor/customer
                Return stammT.Find(Function(_m) _m.id = recId)

            Case "VC_Stamm"
                ' global vendor/customer
                Return vcCHStammT.Find(Function(_vc) _vc.id = recId)
        End Select


        ' get mandant specific data.. a mandant must be set already
        If Not SQLLedgerMandantIsSet() Then
            Return Nothing
        End If

        Select Case tableId
            Case "SLVendor"
                ' local vendor
                Return slVendors.Find(Function(_r) _r.id = recId)

            Case "SLCustomer"
                ' local customer
                Return slCustomers.Find(Function(_r) _r.id = recId)

            Case "SLProject"
                ' ledger project
                Return GetSLProjects.Find(Function(_r) _r.id = recId)

            Case "SLDept"
                ' ledger departement (Kostenstelle)
                Return GetSLDepartments.Find(Function(_r) _r.id = recId)
        End Select

        Throw New Exception("RMA2D.GetByXMLId: Unknown table name.")

    End Function


    Public Shared Function GetIdFromXMLId(ByVal xmlId As String) As Integer
        Dim m = Regex.Match(xmlId, "\w+¬(\d+)")
        If Not m.Success Then
            Return Integer.MinValue
        End If

        Return m.Groups(1).Value
    End Function




    '=============================================================================================================
    '
    '       Stammdaten
    '
    '


    '
    '   contact
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("firstname = {firstname}")> _
    <DebuggerDisplay("name = {name}")> _
    <DebuggerDisplay("acronym = {acronym}")> _
    <DebuggerDisplay("address1 = {address1}")> _
    <DebuggerDisplay("address2 = {address2}")> _
    <DebuggerDisplay("address3 = {address3}")> _
    <DebuggerDisplay("zip = {zip}")> _
    <DebuggerDisplay("city = {city}")> _
    <DebuggerDisplay("email = {email}")> _
    <DebuggerDisplay("tel1 = {tel1}")> _
    <DebuggerDisplay("tel2 = {tel3}")> _
    <DebuggerDisplay("fax = {fax}")> _
    <DebuggerDisplay("anrede = {anrede}")> _
    Public Class Contact
        Public id As Integer
        Public firstname As String
        Public name As String
        Public acronym As String
        Public address1 As String
        Public address2 As String
        Public address3 As String
        Public zip As String
        Public city As String
        Public email As String
        Public tel1 As String
        Public tel2 As String
        Public fax As String
        Public anrede As String
    End Class


    Public Shared Function LoadContact(Optional ByVal ids As IEnumerable(Of Integer) = Nothing) As Dictionary(Of Integer, Contact)
        ' loads & returns all matching contacts
        ' if idSelectSql is Nothing, ALL contacts match.
        ' ids may contain contact ids to load..

        Dim result = New Dictionary(Of Integer, Contact)
        If ids.Count() = 0 Then
            Return result
        End If

        Dim sql As String
        If ids Is Nothing Then
            sql = "select * from contact"
        Else
            Dim idListSB = ids.Aggregate("", Function(_s, _i) _s & " " & _i)
            idListSB = idListSB.Trim.Replace(" ", ",")
            sql = String.Format("select * from contact where id in ({0})", idListSB)
        End If
        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New Contact
            rec.id = rawrec.AsInteger("id")
            rec.firstname = rawrec.AsString("firstname")
            rec.name = rawrec.AsString("name")
            rec.acronym = rawrec.AsString("acronym")
            rec.address1 = rawrec.AsString("address1")
            rec.address2 = rawrec.AsString("address2")
            rec.address3 = rawrec.AsString("address3")
            rec.zip = rawrec.AsString("zip")
            rec.city = rawrec.AsString("city")
            rec.email = rawrec.AsString("email")
            rec.tel1 = rawrec.AsString("tel1")
            rec.tel2 = rawrec.AsString("tel2")
            rec.fax = rawrec.AsString("fax")
            rec.anrede = rawrec.AsString("anrede")
            result(rec.id) = rec
        Next

        Return result
    End Function


    Public Shared Function Get_Role_Contacts_Of_Mandant(ByVal mandant_ids As IEnumerable(Of Integer)) As Dictionary(Of Integer, Tuple(Of Integer, Integer, Integer, Integer))
        ' returns 4 contact ids (responsible, delegate, assistant, Mandatsleiter) for each given mandant

        Dim result As New Dictionary(Of Integer, Tuple(Of Integer, Integer, Integer, Integer))
        If (mandant_ids.Count() = 0) Then
            Return result
        End If

        Dim idList = mandant_ids.Aggregate("", Function(_s, _i) _s & " " & _i)
        idList = idList.Trim.Replace(" ", ",")

        Dim sql = String.Format("select x.mandant_ref, x.role_ref, x.contact_ref from xref_mandant_contact_role x where " & _
                                       "x.role_ref in (15, 18, 19, 36) and x.mandant_ref in ({0})", idList)

        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            Dim mandant_id = rawrec.AsInteger("mandant_ref")
            Dim role_id = rawrec.AsInteger("role_ref")
            Dim contact_id = rawrec.AsInteger("contact_ref")

            Dim tiii As Tuple(Of Integer, Integer, Integer, Integer)
            If Not result.ContainsKey(mandant_id) Then
                tiii = New Tuple(Of Integer, Integer, Integer, Integer)(-1, -1, -1, -1)
                result(mandant_id) = tiii
            Else
                tiii = result(mandant_id)
            End If

            Select Case role_id
                Case 15
                    result(mandant_id) = New Tuple(Of Integer, Integer, Integer, Integer)(contact_id, tiii.Item2, tiii.Item3, tiii.Item4)
                Case 18
                    result(mandant_id) = New Tuple(Of Integer, Integer, Integer, Integer)(tiii.Item1, contact_id, tiii.Item3, tiii.Item4)
                Case 19
                    result(mandant_id) = New Tuple(Of Integer, Integer, Integer, Integer)(tiii.Item1, tiii.Item2, contact_id, tiii.Item4)
                Case 36
                    result(mandant_id) = New Tuple(Of Integer, Integer, Integer, Integer)(tiii.Item1, tiii.Item2, tiii.Item3, contact_id)
            End Select
        Next

        Return result
    End Function



    Public Shared Function Get_Role_Mandants_Of_Contact(ByVal ofContactId As Integer, Optional ByVal as_OLB As Boolean = True,
                                                                                      Optional ByVal as_STV As Boolean = False,
                                                                                      Optional ByVal as_SBV As Boolean = False) As Dictionary(Of Integer, String)
        ' loads & returns all mandant ids where the given contact appears in the specified role(s):
        '   as_OLB      "mandant:responsible"
        '   as_STV      "mandant:delegate"
        '   as_SBV      "mandant:assistant"
        '   as_ML       "mandant:leader"        not yet implemented

        Dim result = New Dictionary(Of Integer, String)

        Dim xlatRole = New Dictionary(Of String, String) From
        {
            {"mandant:responsible", "OLB"},
            {"mandant:delegate", "STV"},
            {"mandant:assistant", "SBV"}
        }

        Dim roles As New List(Of String)
        If as_OLB Then
            roles.Add("mandant:responsible")
        End If
        If as_STV Then
            roles.Add("mandant:delegate")
        End If
        If as_SBV Then
            roles.Add("mandant:assistant")
        End If
        If roles.Count = 0 Then
            Return result
        End If

        Dim sql = String.Format("select distinct x.mandant_ref, r.name from xref_mandant_contact_role x, rma_role r where " & _
                                "x.contact_ref = {0} and x.role_ref = r.id and r.name in ('{1}')", ofContactId, Join(roles.ToArray, "', '"))

        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)
        Dim tempResult = New List(Of Tuple(Of Integer, String))

        For Each rawrec In rawData
            tempResult.Add(New Tuple(Of Integer, String)(rawrec.AsInteger("mandant_ref"), rawrec.AsString("name")))
        Next

        For Each _key In (From _rec In tempResult Select _rec.Item1 Distinct)
            result(_key) = Join((From _rec In tempResult Where _rec.Item1 = _key Select xlatRole(_rec.Item2)).ToArray, ", ")
        Next

        Return result
    End Function




    '
    '   Spesenmitarbeiter
    '
    '

    Public Class SpesenMitarbeiter
        Public mandant_id As Integer
        Public firstname As String
        Public name As String
        Public email As String
        Public account As String
    End Class

    Private Shared _spesenmaT As New List(Of SpesenMitarbeiter)
    Public Shared ReadOnly Property spesenmaT As List(Of SpesenMitarbeiter)
        Get
            SyncLock _spesenmaT
                If MustReload("spesenmaT") Then
                    _spesenmaT.Clear()
                    _spesenmaT.AddRange(LoadSpesenMitarbeiter())
                End If
            End SyncLock
            Return _spesenmaT
        End Get
    End Property

    Private Shared Function LoadSpesenMitarbeiter() As List(Of SpesenMitarbeiter)
        ' loads & returns all contacts that are 'Spesen'-enabled :)
        ' i.e. that are tied to a mandant with the role 'mandant:spesen'

        Dim sml = New List(Of SpesenMitarbeiter)

        Dim sql = "select x.mandant_ref, c.firstname, c.name, c.email, x.role_data " & _
                  "from xref_mandant_contact_role x, rma_role r, contact c where " & _
                  "x.contact_ref = c.id and x.role_ref = r.id and r.name = 'mandant:spesen'"

        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New SpesenMitarbeiter
            rec.mandant_id = rawrec.AsInteger("mandant_ref")
            rec.firstname = rawrec.AsString("firstname", True)
            rec.name = rawrec.AsString("name", True)
            rec.email = rawrec.AsString("email", True)
            rec.account = rawrec.AsString("role_data", True)
            sml.Add(rec)
        Next

        Return sml
    End Function




    '
    '   Mandanten
    '
    '
    Public Const MS_Reserve = "reserve"
    Public Const MS_Setup = "setup"
    Public Const MS_Aufarbeitung = "aufarbeitung"
    Public Const MS_Aktiv = "aktiv"
    Public Const MS_Gesperrt = "gesperrt"
    Public Const MS_Gekündigt = "gekündigt"
    Public Const MS_Beendet = "beendet"

    Public Class StammRecord
        ' table 'mandant':
        Public id As Integer
        Public mandant As String        ' field 'name_id'
        Public abschluss As Date
        Public defclose As Date
        Public lastBarcode As String
        Public barcodeReserve As Integer
        Public zahlungsModus As String
        Public zahlungsDelay As Integer
        Public canceledby As Date?
        Public defaultDepartment As Integer?

        ' table 'mandant_history_items':
        Public validFrom As Date
        Public validTo As Date?
        Public status As String

        Public doDMS As Boolean
        Public doZahlung As Boolean
        Public mwst As String
        Public doLohn As Boolean
        Public doLohnZahlungen As Boolean
        Public hasTags As Boolean
        Public lohnverbuchung As String     ' mandant_lohnverbuchung: 'unknown', 'Aufwand', 'Sollstellung'
        Public kkBezahlung As String        ' mandant_kreditkartenzahlung: 'keine', 'LSV', 'Rechnung'
        Public treuhand_mandant_ref As Integer?
        Public has_OBH As Boolean           ' this mandant has no OBH support
        Public barcodeless As Boolean       ' this mandant may send documents without barcodes

        ' from contact with role mandant:main
        Public firma As String = ""

        ' from login tied to contact with role login:rma_system
        Public user As String = ""
        Public pw As String = ""

        ' from login tied to contact with role mandant:responsible
        Public rmav As String = ""        ' firstname & " " & name
        Public rmav_acro As String = ""   ' acronym

        ' from login tied to contact with role mandant:esr_contact
        Public esrContacts As New List(Of Contact)


        Public Function FirmaAndID(Optional includeNumericalID As Boolean = False) As String
            Dim ids = New List(Of String)
            If (Not Regex.Replace(firma.ToLower(), "[\W]", "").Contains(mandant)) Then
                ' stringish id is only included, if it is not kinda part of the full company name
                ids.Add(mandant)
            End If
            If (includeNumericalID) Then
                ids.Add(id.ToString())
            End If

            If ids.Count() = 0 Then
                Return firma
            End If
            Return String.Format("{0} ({1})", firma, String.Join(", ", ids.ToArray()))
        End Function


        Private _MainCurrency As String = Nothing
        Public Function GetMainCurrency() As String
            If (_MainCurrency IsNot Nothing) Then
                Return _MainCurrency
            End If

            ' the main currency of a mandant is defined as the currency in which the Ledger book keeping is done.
            ' --> the 'no conversion accounting currency'
            Dim currentDBMandant As StammRecord = Nothing
            If (Not SQLLedgerMandantIsSet(currentDBMandant)) Then
                SetNewSQLLedgerMandant(Me)
            ElseIf (currentDBMandant <> Me) Then
                Throw New ApplicationException("Error getting MainCurrency: DB Mandant must be set correctly.")
            End If

            _MainCurrency = dbSQLLedger.SQL2O("select curr from curr where rn = 1")
            If (_MainCurrency Is Nothing) Then
                Throw New ApplicationException("Error getting MainCurrency: No currency with rn = 1 found.")
            End If

            Return _MainCurrency
        End Function

        Private Shared _HasSwitchedToNRLO As List(Of Integer) = Nothing
        Public ReadOnly Property HasSwitchedToNRLO As Boolean
            Get
                If MustReload("HasSwitchedToNRLO") OrElse _HasSwitchedToNRLO Is Nothing Then
                    ' load the ids of all Mandanten that have switched to NRLO (Neue Rechnungslegungsordnung, introduced end of 2014)
                    Dim sql = "select m.id from mandant m " & _
                              "join xref_mandant_feature x on (m.id = x.mandant_ref and x.valid_from <= now() and x.valid_to is null) " & _
                              "join mandant_feature f on (f.id = x.feature_ref and f.name = 'NRLO')"
                    Dim result = DBAccess.dbStamm.SQL2LO(sql)
                    _HasSwitchedToNRLO = result.Select(Function(_o) DirectCast(_o, Integer)).ToList()
                End If
                Return _HasSwitchedToNRLO.Contains(id)
            End Get
        End Property

        Private _AssignedContactIDs As Tuple(Of Integer, Integer, Integer, Integer) = Nothing
        Public ReadOnly Property AssignedContactIDs As Tuple(Of Integer, Integer, Integer, Integer)
            ' returns 4 contact ids (responsible, delegate, assistant, Mandatsleiter), each of them is negative if not set
            Get
                If (_AssignedContactIDs Is Nothing) Then
                    Dim acidList = Get_Role_Contacts_Of_Mandant(New List(Of Integer) From {id})
                    If (acidList.Count = 1) Then
                        _AssignedContactIDs = acidList(id)
                    End If
                End If

                Return _AssignedContactIDs
            End Get
        End Property


        Public ReadOnly Property IsSetup As Boolean
            Get
                Return status = MS_Setup
            End Get
        End Property

        Public ReadOnly Property IsAufarbeitung As Boolean
            Get
                Return status = MS_Aufarbeitung
            End Get
        End Property

        Public ReadOnly Property IsAktiv As Boolean
            Get
                Return status = MS_Aktiv
            End Get
        End Property

        Public ReadOnly Property IsGesperrt As Boolean
            Get
                Return status = MS_Gesperrt
            End Get
        End Property

        Public ReadOnly Property IsGekündigt As Boolean
            Get
                Return status = MS_Gekündigt
            End Get
        End Property

        Public ReadOnly Property IsBeendet As Boolean
            Get
                Return status = MS_Beendet
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return firma & If(IsSetup, " (Setup)", "")
        End Function

        Public Function GetXMLId() As String
            Return String.Format("Mandant¬{0} ({1})", id, firma)
        End Function

        Public Shared Operator =(ByVal a As StammRecord, ByVal b As StammRecord) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As StammRecord, ByVal b As StammRecord) As Boolean
            Return Not a = b
        End Operator

    End Class


    Public Class StammRecordTable
        Inherits List(Of StammRecord)


        Public Function FindMandant(ByVal str As String) As StammRecord
            If str Is Nothing Then
                Return Nothing
            End If

            ' direct hit
            str = str.ToLower.Trim
            Dim mandant = Find(Function(_m) _m.firma.ToLower = str Or _m.id.ToString = str Or _m.mandant = str)
            If mandant IsNot Nothing Then
                Return mandant
            End If

            ' single match on split string
            Dim parts = Regex.Split(str, "\s").ToList
            Dim mandanten = FindAll(Function(_m) parts.All(Function(_p) _m.firma.ToLower.Contains(_p)))
            If mandanten.Count = 1 Then
                ' exactly 1 mandant fits..
                Return mandanten(0)
            End If

            Return Nothing
        End Function
    End Class


    Private Shared _stammT As New StammRecordTable
    Public Shared ReadOnly Property stammT As StammRecordTable
        ' returns a list of all active Mandanten
        Get
            SyncLock _stammT
                If MustReload("stammT") Then
                    LoadStammdaten()
                End If
            End SyncLock
            Return _stammT
        End Get
    End Property

    Private Shared _stammT_All As New List(Of StammRecord)
    Public Shared ReadOnly Property stammT_All As List(Of StammRecord)
        ' returns a list of ALL mandant records, no matter what their status is
        Get
            Dim dummy = stammT
            Return _stammT_All
        End Get
    End Property

    Public Shared Sub LoadStammdaten()
        _stammT.Clear()
        _stammT_All.Clear()

        ' load mandant login/pw
        Dim sql = "select x.mandant_ref, l.login, l.password from mandant m, xref_mandant_contact_role x, contact c, login l, rma_role r where " & _
                  "x.contact_ref = c.id and x.role_ref = r.id and x.mandant_ref = m.id and r.name = 'login:rma_system' and c.id = l.contact_ref and x.id = l.xref_ref"
        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)
        Dim logins = rawData.ToDictionary(Function(_rd) _rd.AsInteger("mandant_ref"), Function(_rd) New With {.login = _rd.AsString("login"), .pw = _rd.AsString("password")})

        ' load company names
        sql = "select x.mandant_ref, c.name from mandant m, xref_mandant_contact_role x, contact c, rma_role r where " & _
              "x.contact_ref = c.id and x.role_ref = r.id and x.mandant_ref = m.id and r.name = 'mandant:main'"
        rawData = DBAccess.dbStamm.SQL2RD(sql)
        Dim company = rawData.ToDictionary(Function(_rd) _rd.AsInteger("mandant_ref"), Function(_rd) New With {.company = _rd.AsString("name")})

        ' load rma responsible names
        sql = "select x.mandant_ref, c.name, c.firstname, c.acronym from mandant m, xref_mandant_contact_role x, contact c, rma_role r where " & _
              "x.contact_ref = c.id and x.role_ref = r.id and x.mandant_ref = m.id and r.name = 'mandant:responsible'"
        rawData = DBAccess.dbStamm.SQL2RD(sql)
        Dim rmaResponsible = rawData.ToDictionary(Function(_rd) _rd.AsInteger("mandant_ref"), Function(_rd) New With {.rmav = _rd.AsString("name") & " " & _rd.AsString("firstname"), _
                                                                                                                      .rmav_acro = _rd.AsString("acronym")})

        ' load esr contacts
        sql = "select x.mandant_ref, x.contact_ref from xref_mandant_contact_role x, rma_role r where " & _
              "x.role_ref = r.id and r.name = 'mandant:esr_contact'"
        rawData = DBAccess.dbStamm.SQL2RD(sql)
        Dim esrContactXref = From _rd In rawData Select New With {.mandant_id = _rd.AsInteger("mandant_ref"), .contact_id = _rd.AsInteger("contact_ref")}
        Dim esr_c_ids = From _ecx In esrContactXref Select _ecx.contact_id
        Dim esrContact = LoadContact(esr_c_ids)

        sql = String.Format("select m.id as m_id, * from mandant m, mandant_history_items mhi where m.id = mhi.mandant_ref and " & _
                            "mhi.valid_from <= '{0}' and (mhi.valid_to > '{0}' or mhi.valid_to is NULL)", Now.ToShortDateString)
        rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New StammRecord
            ' from table mandant..
            rec.id = rawrec.AsInteger("m_id")
            rec.mandant = rawrec.AsString("name_id")
            rec.abschluss = rawrec.AsDate("abschluss")
            rec.defclose = rawrec.AsDate("defclose")
            rec.lastBarcode = rawrec.AsString("lastbarcode")
            rec.barcodeReserve = rawrec.AsInteger("barcodereserve")
            rec.zahlungsModus = rawrec.AsString("zahlungsmodus")
            rec.zahlungsDelay = rawrec.AsInteger("zahlungsdelay")
            rec.canceledby = rawrec.AsDateNIL("canceledby")
            rec.defaultDepartment = rawrec.AsIntegerNIL("default_department_id")

            ' from table mandant_history_items..
            rec.validFrom = rawrec.AsDate("valid_from")
            rec.validTo = rawrec.AsDateNIL("valid_to")
            rec.status = rawrec.AsString("status")
            rec.doDMS = rawrec.AsBoolean("do_dms")
            rec.doZahlung = rawrec.AsBoolean("do_zahlungen")
            rec.mwst = rawrec.AsString("mwst")
            rec.doLohn = rawrec.AsBoolean("do_lohn")
            rec.doLohnZahlungen = rawrec.AsBoolean("lohnzahlungen_ausfuehren")
            rec.hasTags = rawrec.AsBoolean("has_tags")
            rec.lohnverbuchung = rawrec.AsString("lohn_methode")
            rec.kkBezahlung = rawrec.AsString("kk_zahlung")
            rec.treuhand_mandant_ref = rawrec.AsIntegerNIL("treuhand_mandant_ref")
            rec.has_OBH = rawrec.AsBoolean("has_obh")
            rec.barcodeless = rawrec.AsBoolean("barcodeless")

            ' login
            If logins.ContainsKey(rec.id) Then
                rec.user = logins(rec.id).login
                rec.pw = logins(rec.id).pw
            End If

            ' company name
            If company.ContainsKey(rec.id) Then
                rec.firma = company(rec.id).company
            Else
                rec.firma = rec.mandant
            End If

            ' rma responsible
            If rmaResponsible.ContainsKey(rec.id) Then
                rec.rmav = rmaResponsible(rec.id).rmav
                rec.rmav_acro = rmaResponsible(rec.id).rmav_acro
            End If

            ' esr contacts
            For Each c_id In (From _ecx In esrContactXref Where _ecx.mandant_id = rec.id Select _ecx.contact_id)
                If esrContact.ContainsKey(c_id) Then
                    rec.esrContacts.Add(esrContact(c_id))
                End If
            Next


            If _stammT_All.Exists(Function(_m) _m.id = rec.id) Then
                Throw New Exception("Überschneidende mandant_history_items für Mandant " & rec.id)
            End If
            _stammT_All.Add(rec)
            If Not (rec.status = MS_Reserve Or rec.status = MS_Beendet) Then
                _stammT.Add(rec)
            End If
        Next

    End Sub


    '
    ' Mandanten Features
    '
    '

    Public Const MandantFeatureID_MessageWidget = 1

    Public Shared Function GetMandantFeatures(ByVal m As StammRecord) As List(Of Integer)
        Dim features = New List(Of Integer)

        If m IsNot Nothing Then
            Dim sql = "select f.id from mandant_feature f join xref_mandant_feature x on (f.id = x.feature_ref) " & _
                      String.Format("where (x.mandant_ref = {0} And x.valid_from < Now() And (x.valid_to Is null or x.valid_to > Now()))", m.id)
            Dim rawData = DBAccess.dbStamm.SQL2RD(sql)
            features = (From _rd In rawData Select _rd.AsInteger("id")).ToList
        End If

        Return features
    End Function


    '
    ' RMA roles
    '
    '

    Public Const Role_login_belegverarbeiter = 7
    Public Const Role_login_belegsupervisor = 8
    Public Const Role_login_developer = 9
    Public Const Role_login_zahlungsverantwortlicher = 10
    Public Const Role_login_mandanten_administrator = 11
    Public Const Role_login_rma_mitarbeiter = 12
    Public Const Role_login_rma_system = 13
    Public Const Role_mandant_contact = 14
    Public Const Role_mandant_responsible = 15
    Public Const Role_mandant_esr_contact = 16
    Public Const Role_login_api_user = 17
    Public Const Role_mandant_delegate = 18
    Public Const Role_mandant_assistant = 19
    Public Const Role_mandant_main = 20
    Public Const Role_mandant_spesen = 21
    Public Const Role_sqlledger_massgeschneidert = 22
    Public Const Role_sqlledger_finanzverantwortlicher = 23
    Public Const Role_sqlledger_treuhänder = 24
    Public Const Role_sqlledger_run_my_accounts_user = 25
    Public Const Role_sqlledger_sachbearbeiter = 26
    Public Const Role_sqlledger_verkäufer = 27
    Public Const Role_mandant_blind_contact = 28
    Public Const Role_mandant_treuhand_contact = 29
    Public Const Role_login_gl = 30
    Public Const Role_mandant_lohn = 31
    Public Const Role_mandant_nachsendeadresse = 32
    Public Const Role_login_rma_tool = 33
    Public Const Role_mandant_abschluss = 34


    '
    ' RMA Employees
    '
    '

    <DebuggerDisplay("id = {contact_id}")> _
    <DebuggerDisplay("vorname = {vorname}")> _
    <DebuggerDisplay("nachname = {nachname}")> _
    <DebuggerDisplay("zeichen = {zeichen}")> _
    <DebuggerDisplay("email = {email}")> _
    Public Class RMAMitarbeiter
        ' record data (from table Stammdaten.contact)
        Public contact_id As Integer
        Public vorname As String
        Public nachname As String
        Public zeichen As String
        Public email As String

        Public ReadOnly Property fullname As String
            Get
                Return RMA2S.EasyJoin(" ", vorname, nachname)
            End Get
        End Property

        Private Shared ContactImageRoot As String = Nothing
        '
        Public Function GetContactImagePath() As FileInfo
            If (ContactImageRoot Is Nothing) Then
                ContactImageRoot = AppConfig.GetItem("ContactImages", "")
            End If

            Dim cimfi = New FileInfo(CombinePathFile(ContactImageRoot, String.Format("{0}.png", contact_id)))
            If cimfi.Exists() Then
                Return cimfi
            End If

            cimfi = New FileInfo(CombinePathFile(ContactImageRoot, "unknown.png"))
            If cimfi.Exists() Then
                Return cimfi
            End If

            Return Nothing
        End Function

        Public Function GetPicture() As Bitmap
            ' direct code (see below) - the pics a directly loaded from disk
            Dim cimfi = GetContactImagePath()
            If (cimfi Is Nothing) Then
                Return Nothing
            End If

            Dim img2 As Bitmap = Nothing
            Try
                img2 = Bitmap.FromFile(cimfi.FullName)
            Catch ex As Exception
            End Try
            Return img2

            ' below the code to retrieve the contact images using the backend.. 
            ' Dim url = backendRoot & String.Format("contacts/images/{0}.png?api_key={1}", contact_id, backendAPIKey)
            ' Dim request = Net.HttpWebRequest.Create(url)
            ' Dim response = request.GetResponse()
            ' Dim img = Bitmap.FromStream(response.GetResponseStream())
            ' response.Close()
            ' Return img
        End Function

        Private _AllRoles As HashSet(Of Integer) = Nothing
        '
        Public ReadOnly Property AllRoles() As HashSet(Of Integer)
            Get
                If _AllRoles Is Nothing Then
                    Dim sql = "select distinct role_ref from xref_mandant_contact_role where contact_ref = " & contact_id.ToString()
                    Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

                    _AllRoles = New HashSet(Of Integer)
                    For Each rawrec In rawData
                        _AllRoles.Add(rawrec.AsInteger("role_ref"))
                    Next
                End If
                Return _AllRoles
            End Get
        End Property

        Public Function HasRole(roleId As Integer) As Boolean
            Return AllRoles.Contains(roleId)
        End Function

        Public Function GetLogID() As String
            Return String.Format("{0}, cid={1}", fullname, contact_id.ToString())
        End Function

        Public Overrides Function ToString() As String
            Return fullname
        End Function

        Public Shared Operator =(ByVal a As RMAMitarbeiter, ByVal b As RMAMitarbeiter) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.contact_id = b.contact_id
        End Operator

        Public Shared Operator <>(ByVal a As RMAMitarbeiter, ByVal b As RMAMitarbeiter) As Boolean
            Return Not a = b
        End Operator

    End Class


    Private Shared _mitarbeiterT As New List(Of RMAMitarbeiter)
    Public Shared ReadOnly Property mitarbeiterT As List(Of RMAMitarbeiter)
        Get
            SyncLock _mitarbeiterT
                If MustReload("mitarbeiterT") Then
                    _mitarbeiterT.Clear()
                    _mitarbeiterT.AddRange(LoadRMAMitarbeiter())
                End If
            End SyncLock
            Return _mitarbeiterT
        End Get
    End Property

    Private Shared Function LoadRMAMitarbeiter() As List(Of RMAMitarbeiter)
        Dim mT As New List(Of RMAMitarbeiter)

        Dim sql = "select c.* from contact c " & _
                    "inner join xref_mandant_contact_role xmcr on xmcr.contact_ref = c.id " & _
                    "inner join rma_role r on xmcr.role_ref = r.id " & _
                    "where r.name = 'login:rma-mitarbeiter' and not c.acronym is Null"

        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New RMAMitarbeiter
            rec.contact_id = rawrec.AsInteger("id")
            rec.vorname = rawrec.AsString("firstname")
            rec.nachname = rawrec.AsString("name")
            rec.zeichen = rawrec.AsString("acronym")
            rec.email = rawrec.AsString("email")

            mT.Add(rec)
        Next
        Return mT
    End Function

    Public Shared Function GetRMAMitarbeiterByFullName(ByVal fullName As String) As RMAMitarbeiter

        Dim rmaM = mitarbeiterT.Find(Function(_m) fullName Like String.Format("*{0}*", _m.vorname) And _
                                                  fullName Like String.Format("*{0}*", _m.nachname))

        Return rmaM
    End Function



    ' Identification by MAC address

    Public Shared Sub RegisterContact4MAC(ByVal mac As String, contact As RMAMitarbeiter)
        mac = Regex.Replace(StringNothing2Empty(mac).ToLower, "[^\da-f]", "")
        If (mac.Length <> 12) Or (contact Is Nothing) Then
            Return
        End If
        mac = mac.Insert(10, ":").Insert(8, ":").Insert(6, ":").Insert(4, ":").Insert(2, ":")

        Dim sql = String.Format("delete from mac_address where mac = {0};" & _
                                "insert into mac_address (id, name, mac, contact_id) values (default, {1}, {0}, {2})",
                                PSQL_String(mac), PSQL_String(contact.fullname), PSQL_Int(contact.contact_id))
        DBAccess.dbServicePortal.SQLExec(sql)
    End Sub

    Private Shared Function FindContact4MAC(ByRef mac As String) As RMAMitarbeiter
        mac = Regex.Replace(StringNothing2Empty(mac).ToLower, "[^\da-f]", "")
        If mac.Length <> 12 Then
            mac = Nothing
            Return Nothing
        End If
        mac = mac.Insert(10, ":").Insert(8, ":").Insert(6, ":").Insert(4, ":").Insert(2, ":")

        Dim sql = String.Format("select * from mac_address where mac = {0}", PSQL_String(mac))
        Dim rawdata = DBAccess.dbServicePortal.SQL2RD(sql)

        For Each rawrec In rawdata
            Dim cid = rawrec.AsInteger("contact_id")
            Dim rmaM = mitarbeiterT.Find(Function(_m) _m.contact_id = cid)
            If rmaM IsNot Nothing Then
                Return rmaM
            End If
        Next

        Return Nothing
    End Function

    Public Shared Function IdentifyRMAMitarbeiterByMAC(Optional ByRef mac As String = Nothing) As RMAMitarbeiter

        mac = Nothing

        ' see, if the local machine is known - look for MAC address in
        ' ipconfig -all
        Dim process As New Process()
        With process.StartInfo
            .FileName = "ipconfig.exe"
            .Arguments = "-all"
            .CreateNoWindow = True
            .ErrorDialog = False
            .RedirectStandardOutput = True
            .UseShellExecute = False
        End With
        process.Start()
        process.WaitForExit(5000)
        Dim output As String = process.StandardOutput.ReadToEnd()

        Dim m = Regex.Match(output, "([\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2})", RegexOptions.IgnoreCase)
        If m.Success Then
            mac = m.Groups(1).Value
            Dim ma = FindContact4MAC(mac)
            If ma IsNot Nothing Then
                Return ma
            End If
        End If


        ' is this a Remote Desktop connection?

        ' netstat -aonp tcp
        process = New Process()
        With process.StartInfo
            .FileName = "netstat.exe"
            .Arguments = "-aonp tcp"
            .CreateNoWindow = True
            .ErrorDialog = False
            .RedirectStandardOutput = True
            .UseShellExecute = False
        End With
        process.Start()
        process.WaitForExit(5000)
        output = process.StandardOutput.ReadToEnd()

        m = Regex.Match(output, ":3389\s+(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}):\d+\s+HERGESTELLT", RegexOptions.IgnoreCase)

        If m.Success Then
            Dim remoteIP = m.Groups(1).Value

            ' find remote MAC with
            ' arp -a ipAddr
            process = New Process()
            With process.StartInfo
                .FileName = "arp.exe"
                .Arguments = "-a " & remoteIP
                .CreateNoWindow = True
                .ErrorDialog = False
                .RedirectStandardOutput = True
                .UseShellExecute = False
            End With
            process.Start()
            process.WaitForExit(5000)
            output = process.StandardOutput.ReadToEnd()

            m = Regex.Match(output, "([\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2}-[\da-f]{2})", RegexOptions.IgnoreCase)
            If m.Success Then
                mac = m.Groups(1).Value
                Dim ma = FindContact4MAC(mac)
                If ma IsNot Nothing Then
                    Return ma
                End If
            End If
        End If

        Return Nothing
    End Function




    '
    ' Table Stammdaten.zahlungskonto
    '
    '

    Public Class ZStammRecord
        Public Property id As Integer
        Public Property mandant_ref As Integer
        Public Property bezeichnung As String
        Public Property bh_konto As String
        Public Property konto_nr As String
        Public Property konto_nr_mt940 As String
        Public Property konto_iban As String
        Public Property account_type As String
        Public Property bank_kurzname As String
        Public Property bank_blz As String
        Public Property bank_land As String
        Public Property currency As String
        Public Property associated_esr As String
        Public Property esr_tn As String
        Public Property esr_kid As String
        Public Property esr_enabled As Boolean
        Public Property dta_id As String
        Public Property dta_abs As String
        Public Property request_auszug_months As Integer
        Public Property mt940_from As Date
        Public Property mt940_to As Date
        Public Property valid_for_payments As Boolean
        Public Property is_standard As Boolean

        Public ReadOnly Property doMT940 As Boolean
            Get
                Return RMA2S.CheckDatePeriod(mt940_from, mt940_to)
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return bezeichnung
        End Function

        Public Shared Operator =(ByVal a As ZStammRecord, ByVal b As ZStammRecord) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As ZStammRecord, ByVal b As ZStammRecord) As Boolean
            Return Not a = b
        End Operator

    End Class



    Private Shared _zstammT As New List(Of ZStammRecord)
    Public Shared ReadOnly Property zstammT As List(Of ZStammRecord)
        Get
            SyncLock _zstammT
                If MustReload("zstammT") Then
                    _zstammT = DBAccess.dbStamm.LoadDBItems(Of ZStammRecord)("select * from zahlungskonto order by id")
                End If
            End SyncLock
            Return _zstammT
        End Get
    End Property


    '
    ' Table Stammdaten.bankfeiertage
    '
    '

    Private Shared _bankfeiertagT As New List(Of Date)
    Public Shared ReadOnly Property bankfeiertagT As List(Of Date)
        Get
            SyncLock _bankfeiertagT
                If MustReload("bankfeiertagT") Then
                    _bankfeiertagT.Clear()
                    _bankfeiertagT.AddRange(LoadBankfeiertage())
                End If
            End SyncLock
            Return _bankfeiertagT
        End Get
    End Property

    Private Shared Function LoadBankfeiertage() As List(Of Date)
        Dim bftT As New List(Of Date)

        Dim sql = "select * from bankfeiertage"
        Dim rawData = DBAccess.dbStamm.SQL2RD(sql)

        For Each rawrec In rawData
            bftT.Add(rawrec.AsDate("feiertag"))
        Next
        Return bftT
    End Function



    '=============================================================================================================
    '
    '       SQL-Ledger
    '
    ' SQL-Ledger is Mandant oriented.. it creates an own Data Base for each Mandant, requiring
    ' an individual set of connections for each Mandant.
    '
    ' While the connections are pooled for every Mandant (see RMA2_DBRoot), this section of RMA2_BaseData
    ' is constructed such that it throws away all loaded data tables whenever the Mandant changes. The idea
    ' behind this that the usage of SQL-Ledger data should be Mandant centered, meaning that if you work
    ' on a Mandant, then do all required tasks in one go.
    '

    Private Shared sqlLedger_Mandant As StammRecord = Nothing
    Private Shared sqlLedger_Key As Integer = 0

    Public Shared Function SetNewSQLLedgerMandant(ByVal newMandant As String) As Boolean
        ' sets a new mandant for all SQLLedger data getters
        ' returns True if the mandant has changed, False if the same mandant was already set before

        If newMandant Is Nothing Then
            Return SetNewSQLLedgerMandant(CType(Nothing, StammRecord))
        End If

        ' find the specified mandant..
        Dim mRec = stammT.Find(Function(_s) _s.mandant = newMandant.Trim.ToLower)
        If mRec Is Nothing Then
            Throw New Exception("SetNewSQLLedgerMandant: invalid mandant name.")
        End If
        Return SetNewSQLLedgerMandant(mRec)
    End Function

    Public Shared dbSQLLedger As WPF_Roots.ConnectionPool(Of NpgsqlConnection) = Nothing
    '
    Public Shared Function SetNewSQLLedgerMandant(ByVal newMandant As StammRecord) As Boolean
        ' sets a new mandant for all SQLLedger data getters
        ' returns True if the mandant has changed, False if the same mandant was already set before

        If newMandant Is Nothing Then
            ' setting a NULL mandant invalidates SQLLedger access
            Dim hasChanged = (sqlLedger_Mandant IsNot Nothing)
            InvalidateSQLLedgerMandant()
            Return hasChanged
        End If

        If newMandant.Equals(sqlLedger_Mandant) Then
            ' no change
            Return False
        End If

        InvalidateSQLLedgerMandant()
        sqlLedger_Key += 1
        sqlLedger_Mandant = newMandant

        ' initialize db connection pool to mandant tables (SQLLedger)
        dbSQLLedger = DBAccess.GetLedgerPool(newMandant.mandant)

        ' load MainCurrency
        Dim dummy = newMandant.GetMainCurrency()

        Return True
    End Function

    Public Shared Sub InvalidateSQLLedgerMandant()
        ' any subsequent attempt to get SQLLedger data will fail..
        If sqlLedger_Mandant IsNot Nothing Then
            dbSQLLedger = Nothing
        End If
        sqlLedger_Mandant = Nothing
    End Sub

    Public Shared Function SQLLedgerMandantIsSet(Optional ByRef mandant As StammRecord = Nothing) As Boolean
        mandant = sqlLedger_Mandant
        Return (sqlLedger_Mandant IsNot Nothing)
    End Function

    Private Shared Sub ReloadSqlLedger(ByVal key As String)
        Reload(key & "¬" & sqlLedger_Key)
    End Sub

    Private Shared Function MustReloadSqlLedger(ByVal key As String) As Boolean
        If sqlLedger_Mandant Is Nothing Then
            Throw New ApplicationException("RMA2_BaseData: Call RMA2BD_SetNewSQLLedgerMandant() before trying to use SQLLedger data.")
        End If

        Return MustReload(key & "¬" & sqlLedger_Key)
    End Function


    '
    ' Table SQLLedger.(mandant).chart
    '
    '

    ' symbolic links
    Public Enum Chart_SymbolicLink
        ChSL_null
        ChSL_kassa
        ChSL_verrechnungssteuer
        ChSL_lohnzahlung
        ChSL_kontokorrent
        ChSL_debitorskonto
        ChSL_debitordifferenz
        ChSL_miete
        ChSL_schuldzins
        ChSL_drittspesen
        ChSL_kreditorsifferenz
        ChSL_guthabenzins
        ChSL_transfer
        ChSL_abklärung

        ChSL_forderungenschweiz
        ChSL_forderungenauslandchf
        ChSL_forderungenauslandandere

        ChSL_durchlaufkonto0
        ChSL_durchlaufkonto1
        ChSL_durchlaufkonto2
        ChSL_durchlaufkonto3
        ChSL_durchlaufkonto4
        ChSL_durchlaufkonto5
        ChSL_durchlaufkonto6
        ChSL_durchlaufkonto7
        ChSL_durchlaufkonto8
        ChSL_durchlaufkonto9
    End Enum


    Public Class SymbolicChart

        Private Shared Symbols As New List(Of Tuple(Of Chart_SymbolicLink, String)) From
            {
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_kassa, "kassa"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_verrechnungssteuer, "verrechnungssteuer"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_lohnzahlung, "lohnzahlung"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_kontokorrent, "kontokorrent"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_debitorskonto, "debitorskonto"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_debitordifferenz, "debitordifferenz"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_miete, "miete"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_schuldzins, "schuldzins"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_drittspesen, "drittspesen"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_kreditorsifferenz, "kreditordifferenz"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_guthabenzins, "guthabenzins"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_transfer, "transfer"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_abklärung, "abklärung"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_forderungenschweiz, "forderungenschweiz"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_forderungenauslandchf, "forderungenauslandchf"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_forderungenauslandandere, "forderungenauslandandere"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto0, "durchlaufkonto0"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto1, "durchlaufkonto1"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto2, "durchlaufkonto2"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto3, "durchlaufkonto3"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto4, "durchlaufkonto4"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto5, "durchlaufkonto5"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto6, "durchlaufkonto6"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto7, "durchlaufkonto7"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto8, "durchlaufkonto8"),
                New Tuple(Of Chart_SymbolicLink, String)(Chart_SymbolicLink.ChSL_durchlaufkonto9, "durchlaufkonto9")
            }

        Private Shared SymbolicAccounts As List(Of Tuple(Of Chart_SymbolicLink, String, SLChartRecord))
        '
        Public Shared Function GetSymbolicAccountTable() As List(Of Tuple(Of Chart_SymbolicLink, String, SLChartRecord))
            If MustReloadSqlLedger("SymbolicChart") Then
                SymbolicAccounts = New List(Of Tuple(Of Chart_SymbolicLink, String, SLChartRecord))

                Dim chart = GetLedgerChart()
                For Each sym In Symbols
                    Dim accounts = chart.FindAll(Function(_a) _a.symbolic_link.Contains(sym.Item2))
                    For Each _acc In accounts
                        SymbolicAccounts.Add(New Tuple(Of Chart_SymbolicLink, String, SLChartRecord)(sym.Item1, sym.Item2, _acc))
                    Next
                Next
            End If

            Return New List(Of Tuple(Of Chart_SymbolicLink, String, SLChartRecord))(SymbolicAccounts)
        End Function

        Public Shared Function GetAccountSymbol(chsl As Chart_SymbolicLink) As String
            Dim s = Symbols.Find(Function(_s) _s.Item1 = chsl)
            If s IsNot Nothing Then
                Return s.Item2
            End If
            Return Nothing
        End Function

        Public Shared Function GetAccount(chsl As Chart_SymbolicLink) As SLChartRecord
            GetSymbolicAccountTable()

            Dim slt = SymbolicAccounts.Find(Function(_slt) _slt.Item1 = chsl)
            If slt IsNot Nothing Then
                Return slt.Item3
            End If
            Return Nothing
        End Function

        Public Shared Function Exists(chsl As Chart_SymbolicLink) As Boolean
            Return (GetAccount(chsl) IsNot Nothing)
        End Function

        Public Shared Function GetAccountBySymbol(symbol As String) As SLChartRecord
            ' reload chart & symbol table, if necessary
            GetAccount(Chart_SymbolicLink.ChSL_null)

            Dim slt = SymbolicAccounts.Find(Function(_slt) _slt.Item2 = symbol)
            If slt IsNot Nothing Then
                Return slt.Item3
            End If
            Return Nothing
        End Function


        Private Shared _NRLO_Chart_Translation As Dictionary(Of String, String) = Nothing
        Private Shared ReadOnly Property NRLO_Chart_Translation As Dictionary(Of String, String)
            Get
                ' Wenn ein Mandant den Übergang zur NRLO macht, wird eine Konten-Übersetzungstabelle angefertigt.
                ' In misc:chart_nrlo_transition sind unter der entsprechenden Mandanten-Id die (altes Konto, neues Konto)-Paare gespeichert.
                ' Das zurückgegebene Dictionary kann dazu verwendet werden, alte Kontenbezeichnungen automatisch auf den neuen Kontenplan
                ' zu übersetzen.
                If MustReloadSqlLedger("SymbolicChart_NRLO") Then
                    _NRLO_Chart_Translation = New Dictionary(Of String, String)
                    Dim sql = String.Format("select * from chart_nrlo_transition where mandant_ref = {0}", sqlLedger_Mandant.id)
                    Dim rawData = DBAccess.dbMisc.SQL2RD(sql)

                    For Each rawrec In rawData
                        Dim accno = rawrec.AsString("accno")
                        Dim new_accno = rawrec.AsString("new_accno")
                        _NRLO_Chart_Translation(accno) = new_accno
                    Next
                End If
                Return _NRLO_Chart_Translation
            End Get
        End Property


        Public Shared Function ResolveAccount(mandant As StammRecord, account As String, Optional ByRef symbolicAccount As String = Nothing) As String

            account = NoWhitespace(StringNothing2Empty(account).ToLower())
            symbolicAccount = Nothing

            If Regex.IsMatch(account, "^\d+$") Then
                ' direct account number, does not need a translation
                Return account

            ElseIf Regex.IsMatch(account, "old-\d+") Then
                ' reference to an pre-NRLO account number - resolve using the NRLO-transition table in db misc
                symbolicAccount = account
                account = (account.Substring(4))      ' cut off 'old-'

                If mandant.HasSwitchedToNRLO Then
                    Dim translated_account As String = Nothing
                    If NRLO_Chart_Translation.TryGetValue(account, translated_account) Then
                        Return translated_account
                    End If
                End If
                Return account

            ElseIf Regex.IsMatch(account, "[a-z]") Then
                ' account number is not purely numeric - try to resolve account symbol
                symbolicAccount = account
                Dim symCh = SymbolicChart.GetAccountBySymbol(account)
                If symCh IsNot Nothing Then
                    Return symCh.accno
                End If
            End If

            Return Nothing
        End Function

    End Class


    Public Const AccID_Bezugsteuer = -1
    Public Const AccID_keineMWSt = -2
    Public Const AccID_Zoll_MatDL = -3
    Public Const AccID_Zoll_InvBA = -4
    '
    Public Class SLChartRecord
        ' record data
        Public id As Integer
        Public accno As String
        Public desc As String        ' ~description
        Public type As String        ' ~charttype
        Public cat As String         ' ~category
        Public link As String
        Public allow_gl As Boolean
        Public symbolic_link As String
        Public taxRate As Double     ' filled only for tax accounts, filled from table 'tax'

        Public Overrides Function ToString() As String
            If link IsNot Nothing AndAlso link.Contains("_tax") Then
                ' tax accounts: make the tax rate lead the display string
                ' - this makes combo box selections easier
                Dim taxDesc = Regex.Replace(desc, "^\D+", "")
                Return String.Format("{0} ({1})", taxDesc, accno).Trim
            End If
            Return String.Format("{0} {1}", accno, desc).Trim
        End Function

        Public Function GetXMLId() As String
            Return accno
        End Function

        ' Public ReadOnly Property IsBezugSteuerAccount As Boolean
        '    Get
        '        Return accno IsNot Nothing AndAlso
        '               ((accno = "1176") Or (accno = "6730") Or (accno = "2204") Or (accno = "Bezugsteuer"))
        '    End Get
        ' End Property

        Public Shared Operator =(ByVal a As SLChartRecord, ByVal b As SLChartRecord) As Boolean
            If Object.ReferenceEquals(a, b) Then
                Return True
            ElseIf DirectCast(a, Object) Is Nothing Then
                Return False
            End If
            Return a.Equals(b)
        End Operator

        Public Shared Operator <>(ByVal a As SLChartRecord, ByVal b As SLChartRecord) As Boolean
            Return Not a = b
        End Operator

        Public Overrides Function Equals(obj As Object) As Boolean
            If (TypeOf (obj) Is SLChartRecord) Then
                Dim otherChart = CType(obj, SLChartRecord)
                Return id = otherChart.id
            End If
            Return False
        End Function
    End Class


    Private Shared _slChart As New List(Of SLChartRecord)
    Public Shared Function GetLedgerChart() As List(Of SLChartRecord)
        SyncLock _slChart
            If MustReloadSqlLedger("slChart") Then
                ' reload data
                _slChart.Clear()
                Dim sql = "Select chart.*, tax.rate " & _
                          "from chart LEFT OUTER JOIN tax ON (chart.id = tax.chart_id) where tax.validto is NULL "
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                For Each rawrec In rawData
                    Dim rec As New SLChartRecord
                    rec.id = rawrec.AsInteger("id")
                    rec.accno = rawrec.AsString("accno")
                    If rec.accno Is Nothing Then
                        Continue For
                    End If

                    rec.desc = rawrec.AsString("description")
                    rec.type = rawrec.AsString("charttype")
                    rec.cat = rawrec.AsString("category")
                    rec.link = rawrec.AsString("link")
                    rec.allow_gl = rawrec.AsBoolean("allow_gl")
                    rec.symbolic_link = rawrec.AsString("symbol_link", nullAsEmpty:=True)
                    rec.taxRate = rawrec.AsDouble("rate")

                    ' Juli 2012: inaktive Konten werden neu mit
                    '   accno = 'xxxx (inaktiv ab 01.07.2012)'
                    ' bezeichnet. Solche Konten werden gar nicht aufgenommen
                    If rec.accno.Contains("(") Or rec.accno.Contains(")") Then
                        ' ignore
                    Else
                        _slChart.Add(rec)
                    End If
                Next
            End If
        End SyncLock

        Return _slChart
    End Function




    '
    ' Table SQLLedger.(mandant).curr
    '
    '

    Public Class SLCurrRecord
        ' record data
        Public rn As Integer
        Public curr As String
        Public precision As Integer

        Public Overrides Function ToString() As String
            Return curr
        End Function

        Public Shared Operator =(ByVal a As SLCurrRecord, ByVal b As SLCurrRecord) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.curr = b.curr
        End Operator

        Public Shared Operator <>(ByVal a As SLCurrRecord, ByVal b As SLCurrRecord) As Boolean
            Return Not a = b
        End Operator
    End Class


    Private Shared _slCurr As New List(Of SLCurrRecord)
    Public Shared Function GetLedgerCurr() As List(Of SLCurrRecord)
        SyncLock _slCurr
            If MustReloadSqlLedger("slCurr") Then
                ' reload data
                _slCurr.Clear()
                Dim sql = "Select rn, curr, precision from curr"
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                For Each rawrec In rawData
                    Dim rec As New SLCurrRecord
                    rec.rn = rawrec.AsInteger("rn")
                    rec.curr = rawrec.AsString("curr")
                    rec.precision = rawrec.AsInteger("precision")
                    _slCurr.Add(rec)
                Next
            End If
        End SyncLock

        Return _slCurr
    End Function




    '
    ' Table SQLLedger.(mandant).department
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("description = {description}")> _
    Public Class SLDepartment
        ' record data
        Public id As Integer
        Public description As String

        Public Overrides Function ToString() As String
            Return description
        End Function

        Public Function GetXMLId() As String
            Return String.Format("SLDept¬{0} ({1})", id, description)
        End Function

        Public Function GetLedgerAPIString() As String
            Return String.Format("{0}--{1}", description, id)
        End Function

        ' object --> String identity
        Public Shared Operator =(ByVal a As SLDepartment, ByVal b As String)
            Return a IsNot Nothing AndAlso a.description.ToLower.StartsWith(b.ToLower)
        End Operator
        Public Shared Operator <>(ByVal a As SLDepartment, ByVal b As String)
            Return Not (a = b)
        End Operator

        ' direct object identity
        Public Shared Operator =(ByVal a As SLDepartment, ByVal b As SLDepartment) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As SLDepartment, ByVal b As SLDepartment) As Boolean
            Return Not a = b
        End Operator

    End Class


    Private Shared _slDept As New List(Of SLDepartment)
    Public Shared Function GetSLDepartments() As List(Of SLDepartment)
        SyncLock _slDept
            If MustReloadSqlLedger("slDept") Then
                Dim sql = "Select * from department"
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                _slDept.Clear()
                For Each rawrec In rawData
                    Dim rec As New SLDepartment
                    rec.id = rawrec.AsInteger("id")
                    rec.description = rawrec.AsString("description")
                    _slDept.Add(rec)
                Next
            End If
        End SyncLock

        Return _slDept
    End Function


    Public Shared Function GetSLDepartmentByKST(ByVal kst As String) As SLDepartment
        ' returns the current mandant's department that starts with the given string (kst = Kostenstelle)
        kst = RMA2S.NoWhitespace(kst).ToLower.Trim
        Dim sld = GetSLDepartments().Find(Function(_d) _d.description.ToLower.Trim.StartsWith(kst))
        Return sld
    End Function



    '
    ' Table SQLLedger.(mandant).project
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("projectnumber = {projectnumber}")> _
    <DebuggerDisplay("description = {description}")> _
    <DebuggerDisplay("startDate = {startDate}")> _
    <DebuggerDisplay("endDate = {endDate}")> _
    <DebuggerDisplay("customerID = {customerID}")> _
    Public Class SLProject
        ' record data
        Public id As Integer
        Public projectnumber As String
        Public description As String
        Public startDate As Date
        Public endDate As Date
        Public customerID As Integer

        Public Overrides Function ToString() As String
            Return description
        End Function

        Public Function GetXMLId() As String
            Return String.Format("SLProject¬{0} ({1})", id, description)
        End Function

        Public Function GetLedgerAPIString() As String
            Return String.Format("{0}--{1}", description, id)
        End Function

        ' object --> String identity
        Public Shared Operator =(ByVal a As SLProject, ByVal b As String)
            Return a IsNot Nothing AndAlso a.description.ToLower.StartsWith(b.ToLower)
        End Operator
        Public Shared Operator <>(ByVal a As SLProject, ByVal b As String)
            Return Not (a = b)
        End Operator

        ' direct object identity
        Public Shared Operator =(ByVal a As SLProject, ByVal b As SLProject) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As SLProject, ByVal b As SLProject) As Boolean
            Return Not a = b
        End Operator
    End Class


    Private Shared _slProj As New List(Of SLProject)
    Public Shared Function GetSLProjects() As List(Of SLProject)
        SyncLock _slProj
            If MustReloadSqlLedger("slProj") Then
                Dim sql = "Select * from project"
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                _slProj.Clear()
                For Each rawrec In rawData
                    Dim rec As New SLProject
                    rec.id = rawrec.AsInteger("id")
                    rec.projectnumber = rawrec.AsString("projectnumber")
                    rec.description = rawrec.AsString("description")
                    If StringEmpty2Nothing(rec.description, True) Is Nothing Then
                        rec.description = rec.projectnumber
                    End If
                    rec.startDate = rawrec.AsDate("startdate")
                    rec.endDate = rawrec.AsDate("enddate")
                    rec.customerID = rawrec.AsInteger("customer_id")
                    _slProj.Add(rec)
                Next
            End If
        End SyncLock

        Return _slProj
    End Function




    '
    ' Table SQLLedger.(mandant).vendor
    ' Table SQLLedger.(mandant).customer
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("name = {name}")> _
    <DebuggerDisplay("isActive = {isActive}")> _
    Public Class SLVCRecord
        ' mandant reference
        Public mandant As StammRecord

        ' record data
        Public id As Integer
        Public name As String
        Public startdate As Date
        Public enddate As Date
        Public isActive As Boolean                        ' set if enddate = NULL
        Public isVendor As Boolean
        Public arap_accno_id As Integer
        Public payment_accno_id As Integer
        Public nrlo_Is_Swiss As Boolean? = Nothing        ' only defined for debitor
        ' Public x As New Dictionary(Of String, Object)   ' space for additional data items


        Public ReadOnly Property GetJoinedGlobalRecord() As VC_CH_Stamm
            ' returns the joined global VC,
            ' or Nothing, if this local VC is not joined.

            Get
                ' load xref tables & break if no global twin exists
                Dim l2gXRef As Dictionary(Of Integer, Integer) = Nothing
                Dim g2lXRef As Dictionary(Of Integer, Integer) = Nothing
                Dim isSwiss As Dictionary(Of Integer, Boolean) = Nothing
                If isVendor Then
                    RMA2D.LoadVendorXREFDicts(mandant.mandant, l2gXRef, g2lXRef)
                Else
                    RMA2D.LoadCustomerXREFDicts(mandant.mandant, l2gXRef, g2lXRef, isSwiss)
                End If
                If Not l2gXRef.ContainsKey(id) Then
                    Return Nothing
                End If

                ' load & return global record
                Dim gid = l2gXRef(id)
                Return vcCHStammT.Find(Function(_ch) _ch.id = gid)
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return String.Format("{0} ({1})", name, id)
        End Function

        Public Function GetXMLId() As String
            If isVendor Then
                Return String.Format("SLVendor¬{0} ({1})", id, name)
            Else
                Return String.Format("SLCustomer¬{0} ({1})", id, name)
            End If
        End Function

        Public Shared Operator =(ByVal a As SLVCRecord, ByVal b As SLVCRecord) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As SLVCRecord, ByVal b As SLVCRecord) As Boolean
            Return Not a = b
        End Operator

    End Class


    Private Shared _slVendors As New List(Of SLVCRecord)
    Public Shared ReadOnly Property slVendors As List(Of SLVCRecord)
        Get
            SyncLock _slVendors
                If MustReloadSqlLedger("slVendor") Then
                    _slVendors.Clear()
                    _slVendors.AddRange(LoadLedgerVendorTable())
                End If
            End SyncLock
            Return _slVendors
        End Get
    End Property

    Public Shared Sub ReloadLedgerVendorTable()
        ReloadSqlLedger("slVendor")
        lvxref_mandant = Nothing
    End Sub

    Private Shared Function LoadLedgerVendorTable() As List(Of SLVCRecord)
        ' returns a fresh copy of the current mandant's vendor table
        Dim slvt As New List(Of SLVCRecord)
        Dim sql = "Select id, name, startdate, enddate, enddate is null as isactive, arap_accno_id, payment_accno_id from vendor"
        Dim rawData = dbSQLLedger.SQL2RD(sql)

        Dim mandant As StammRecord = Nothing
        SQLLedgerMandantIsSet(mandant)

        For Each rawrec In rawData
            Dim rec As New SLVCRecord With {.mandant = mandant, .isVendor = True}
            rec.id = rawrec.AsInteger("id")
            rec.name = rawrec.AsString("name")
            rec.startdate = rawrec.AsDate("startdate")
            rec.enddate = rawrec.AsDate("enddate")
            rec.isActive = rawrec.AsBoolean("isactive")
            rec.arap_accno_id = rawrec.AsInteger("arap_accno_id")
            rec.payment_accno_id = rawrec.AsInteger("payment_accno_id")
            slvt.Add(rec)
        Next
        Return slvt
    End Function

    Public Shared Function TerminateLedgerVendor(ByVal vendorId As Integer) As Boolean
        ' sets the enddate of the given vendor entry to Today

        ' update entry
        Dim sql = String.Format("update vendor set enddate = {0} where id = {1}", RMA2D.PSQL_Date(Now), vendorId)
        Dim affectedRows = dbSQLLedger.SQLExec(sql)

        RMA2D.ReloadLedgerVendorTable()
        Return (affectedRows = 1)
    End Function


    Private Shared _slCustomers As New List(Of SLVCRecord)
    Public Shared ReadOnly Property slCustomers As List(Of SLVCRecord)
        Get
            SyncLock _slCustomers
                If MustReloadSqlLedger("slCustomer") Then
                    _slCustomers.Clear()
                    _slCustomers.AddRange(LoadLedgerCustomerTable())
                End If
            End SyncLock
            Return _slCustomers
        End Get
    End Property

    Public Shared Sub ReloadLedgerCustomerTable()
        ReloadSqlLedger("slCustomer")
        lcxref_mandant = Nothing
    End Sub

    Private Shared Function LoadLedgerCustomerTable() As List(Of SLVCRecord)
        ' returns a fresh copy of the current mandant's customer table
        ' note: only vendors with enddate = NULL are loaded
        Dim slct As New List(Of SLVCRecord)
        Dim sql = "Select id, name, contact, startdate, enddate, enddate is null as isactive from Customer"
        Dim rawData = dbSQLLedger.SQL2RD(sql)

        Dim mandant As StammRecord = Nothing
        SQLLedgerMandantIsSet(mandant)

        Dim l2gXRef As Dictionary(Of Integer, Integer) = Nothing
        Dim g2lXRef As Dictionary(Of Integer, Integer) = Nothing
        Dim isSwiss As Dictionary(Of Integer, Boolean) = Nothing
        LoadCustomerXREFDicts(mandant.mandant, l2gXRef, g2lXRef, isSwiss, False)

        For Each rawrec In rawData
            Dim rec As New SLVCRecord With {.mandant = mandant, .isVendor = False}
            rec.id = rawrec.AsInteger("id")
            rec.name = rawrec.AsString("name")
            Dim contact = StringEmpty2Nothing(rawrec.AsString("contact"))
            If contact IsNot Nothing AndAlso rec.name IsNot Nothing AndAlso
                Not Regex.Split(contact, "\W+").All(Function(_w) rec.name.Contains(_w)) Then
                rec.name &= " - " & contact
            End If
            rec.startdate = rawrec.AsDate("startdate")
            rec.enddate = rawrec.AsDate("enddate")
            rec.isActive = rawrec.AsBoolean("isactive")
            If isSwiss.ContainsKey(rec.id) Then
                rec.nrlo_Is_Swiss = isSwiss(rec.id)
            End If
            slct.Add(rec)
        Next

        Return slct
    End Function

    Public Shared Function TerminateLedgerCustomer(ByVal customerId As Integer) As Boolean
        ' sets the enddate of the given customer entry to Today

        ' update entry
        Dim sql = String.Format("update customer set enddate = {0} where id = {1}", RMA2D.PSQL_Date(Now), customerId)
        Dim affectedRows = dbSQLLedger.SQLExec(sql)

        RMA2D.ReloadLedgerCustomerTable()
        Return (affectedRows = 1)
    End Function




    '
    ' Table SQLLedger.(mandant).ap
    ' Table SQLLedger.(mandant).ar
    '
    '

    Public Class SLARAPRecord
        ' NOTE: all amount signs are changed for AP records.
        ' This makes a concurrent processing of ar and ap records possible

        ' record data
        Public id As Integer
        Public isAR As Boolean          ' set if this entry is a Debitor (ar table)
        Public transdate As Date
        Public vcId As Integer
        Public amount As Double         ' always in CHF
        Public netamount As Double      ' always in CHF
        Public paid As Double           ' always in CHF
        Public duedate As Date
        Public curr As String           ' currency of the original payment
        Public exchangerate As Double   ' value of 1 'curr' in CHF (example: might be 1.5231 for curr = 'EUR')
        Public invnumber As String
        Public ordnumber As String
        Public ponumber As String
        Public description As String
        Public dcn As String

        ' additional stuff
        Public openCHF As Double        ' always in CHF, = amount - paid
        Public openCurr As Double       ' = openCHF / exchangeRate, open amount in original OP currency [curr]
        Public isOpen As Boolean        ' True if openCurr > 0

        ' additional stuff needed in BankIn
        Public exactAmount As Boolean = False
        Public deviationRate As Double = 0.0
        Public vc_ref As SLVCRecord = Nothing


        Public Sub AmountPaidToOP(ByVal payAmount As Double, ByVal curr As String)
            ' adjust memory representation by given amount. curr MUST be == arap.curr
            ' Note: the exchange rate can be neglected here, because Ledger automatically books rate differences
            '  example: a 125 EUR amount always closes an 125 EUR arap, no matter what rate is applied.
            If Me.curr <> curr Then
                Throw New ApplicationException("ARAP mit falscher Währung beglichen.")
            End If

            openCurr = Round(openCurr - payAmount, 2)
            paid = Round(paid + payAmount * exchangerate, 2)
            openCHF = Round(amount - paid, 2)
            isOpen = (openCurr <> 0.0)
        End Sub

    End Class


    ' not the whole ap table is loaded, but only open records.. (amount - paid != 0) is used to decide this.
    ' NOTE: the signs of all amounts are switched for ap records
    Private Shared _slAP As New List(Of SLARAPRecord)
    Public Shared Function GetLedgerAP(Optional ByVal loadFreshCopy As Boolean = False,
                                       Optional ByVal loadAll As Boolean = False) As List(Of SLARAPRecord)
        If loadFreshCopy OrElse loadAll OrElse MustReloadSqlLedger("slAP") Then
            ' reload data
            Dim sql = "Select id, invnumber, ordnumber, ponumber, transdate, vendor_id, amount, netamount, paid, duedate, exchangerate, fxamount, fxpaid, curr, dcn, description from ap"
            If Not loadAll Then
                sql &= " where (amount - paid <> 0)"
            End If
            sql &= " order by id"

            Dim rawData = dbSQLLedger.SQL2RD(sql)

            Dim localAP As New List(Of SLARAPRecord)
            For Each rawrec In rawData
                Dim rec As New SLARAPRecord With {.isAR = False}
                With rec
                    .id = rawrec.AsInteger("id")
                    .invnumber = rawrec.AsString("invnumber")
                    .ordnumber = rawrec.AsString("ordnumber")
                    .ponumber = rawrec.AsString("ponumber")
                    .transdate = rawrec.AsDate("transdate")
                    .vcId = rawrec.AsInteger("vendor_id")
                    .amount = -rawrec.AsDouble("amount")
                    .netamount = -rawrec.AsDouble("netamount")
                    .paid = -rawrec.AsDouble("paid")
                    .duedate = rawrec.AsDate("duedate")
                    .exchangerate = rawrec.AsDouble("exchangerate")
                    .curr = rawrec.AsString("curr")
                    .dcn = rawrec.AsString("dcn", nullAsEmpty:=True)
                    .description = rawrec.AsString("description", nullAsEmpty:=True)

                    .openCHF = Round(.amount - .paid, 2)
                    If .exchangerate <> 0 Then
                        ' this calculation seems to be more Ledger-like
                        .openCurr = Round(.netamount / .exchangerate + (.amount - .netamount) / .exchangerate - .paid / .exchangerate, 2)
                        ' than this:
                        ' .openCurr = Round(.openCHF / .exchangerate, 2)
                    End If

                    ' some records have a broken vc reference..
                    ' some records have ridiculously small, but non-zero values..
                    .vc_ref = slVendors.Find(Function(_v) _v.id = .vcId)
                    If .vc_ref IsNot Nothing AndAlso .openCHF <> 0 Then
                        .isOpen = True
                        localAP.Add(rec)
                    ElseIf loadAll Then
                        .isOpen = False
                        localAP.Add(rec)
                    End If
                End With
            Next

            If loadAll Then
                Return localAP
            End If

            SyncLock _slAP
                _slAP = localAP
            End SyncLock
        End If

        Return _slAP
    End Function


    ' not the whole ar table is loaded, but only open records.. (amount - paid != 0) is used to decide this.
    Private Shared _slAR As New List(Of SLARAPRecord)
    Public Shared Function GetLedgerAR(Optional ByVal loadFreshCopy As Boolean = False,
                                       Optional ByVal loadAll As Boolean = False) As List(Of SLARAPRecord)
        If loadFreshCopy OrElse loadAll OrElse MustReloadSqlLedger("slAR") Then
            ' reload data
            Dim sql = "Select id, invnumber, ordnumber, ponumber, transdate, customer_id, amount, netamount, paid, duedate, exchangerate, fxamount, fxpaid, curr, dcn, description from ar"
            If Not loadAll Then
                sql &= " where (amount - paid <> 0)"
            End If
            sql &= " order by id"

            Dim rawData = dbSQLLedger.SQL2RD(sql)

            Dim localAR As New List(Of SLARAPRecord)
            For Each rawrec In rawData
                Dim rec As New SLARAPRecord With {.isAR = True}
                With rec
                    .id = rawrec.AsInteger("id")
                    .invnumber = rawrec.AsString("invnumber")
                    .ordnumber = rawrec.AsString("ordnumber")
                    .ponumber = rawrec.AsString("ponumber")
                    .transdate = rawrec.AsDate("transdate")
                    .vcId = rawrec.AsInteger("customer_id")
                    .amount = rawrec.AsDouble("amount")
                    .netamount = rawrec.AsDouble("netamount")
                    .paid = rawrec.AsDouble("paid")
                    .duedate = rawrec.AsDate("duedate")
                    .exchangerate = rawrec.AsDouble("exchangerate")
                    .curr = rawrec.AsString("curr")
                    .dcn = rawrec.AsString("dcn", nullAsEmpty:=True)
                    .description = rawrec.AsString("description", nullAsEmpty:=True)

                    .openCHF = Round(.amount - .paid, 2)
                    If .exchangerate <> 0 Then
                        If sqlLedger_Mandant.mandant = "blackroll" Then
                            ' hmm.. even better? NO! gives wrong amounts in some cases!!!
                            ' only for Mandant Blackroll, to prevent +/- 0.01 CHF errors.
                            .openCurr = Round(rawrec.AsDouble("fxamount"), 2) - Round(.paid / .exchangerate, 2)
                            If (.openCurr = 0) Then
                                ' old Debitoren seem to have fxamount = 0
                                .openCurr = Round(.netamount / .exchangerate, 2) + Round((.amount - .netamount) / .exchangerate, 2) - Round(.paid / .exchangerate, 2)
                            End If

                        Else
                            ' this calculation seems to be more Ledger-like
                            .openCurr = Round(.netamount / .exchangerate, 2) + Round((.amount - .netamount) / .exchangerate, 2) - Round(.paid / .exchangerate, 2)
                        End If

                        ' than this:
                        ' .openCurr = Round(.openCHF / .exchangerate, 2)
                    End If

                    ' some records have a broken vc reference..
                    ' some records have ridiculously small, but non-zero values..
                    .vc_ref = slCustomers.Find(Function(_c) _c.id = .vcId)
                    If .vc_ref IsNot Nothing AndAlso .openCHF <> 0 Then
                        .isOpen = True
                        localAR.Add(rec)
                    ElseIf loadAll Then
                        .isOpen = False
                        localAR.Add(rec)
                    End If
                End With
            Next

            If loadAll Then
                Return localAR
            End If

            SyncLock _slAR
                _slAR = localAR
            End SyncLock
        End If

        Return _slAR
    End Function


    '
    ' Table SQLLedger.(mandant).parts
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("partnumber = {partnumber}")> _
    <DebuggerDisplay("description = {description}")> _
    <DebuggerDisplay("unit = {unit}")> _
    Public Class SLPart
        ' record data
        Public id As Integer
        Public partnumber As String
        Public description As String
        Public unit As String

        ' direct object identity
        Public Shared Operator =(ByVal a As SLPart, ByVal b As SLPart) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As SLPart, ByVal b As SLPart) As Boolean
            Return Not a = b
        End Operator
    End Class


    Private Shared _slPart As New List(Of SLPart)
    Public Shared Function GetSLParts() As List(Of SLPart)
        SyncLock _slPart
            If MustReloadSqlLedger("slPart") Then
                Dim sql = "Select id, partnumber, description, unit from parts"
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                _slPart.Clear()
                For Each rawrec In rawData
                    Dim rec As New SLPart
                    rec.id = rawrec.AsInteger("id")
                    rec.partnumber = rawrec.AsString("partnumber")
                    rec.description = rawrec.AsString("description")
                    rec.unit = rawrec.AsString("unit")
                    _slPart.Add(rec)
                Next
            End If
        End SyncLock

        Return _slPart
    End Function


    '
    ' Table SQLLedger.(mandant).defaults
    '
    '

    <DebuggerDisplay("fldname = {fldname}")> _
    <DebuggerDisplay("fldvalue = {fldvalue}")> _
    Public Class SLDefRecord
        ' record data
        Public fldname As String
        Public fldvalue As String
    End Class

    Private Shared _slDefaults As New List(Of SLDefRecord)
    Public Shared Function GetLedgerDef() As List(Of SLDefRecord)
        SyncLock _slDefaults
            If MustReloadSqlLedger("slDefaults") Then
                ' reload data
                _slDefaults.Clear()
                Dim sql = "Select * from defaults"
                Dim rawData = dbSQLLedger.SQL2RD(sql)

                For Each rawrec In rawData
                    Dim rec As New SLDefRecord
                    rec.fldname = rawrec.AsString("fldname")
                    rec.fldvalue = rawrec.AsString("fldvalue")
                    _slDefaults.Add(rec)
                Next
            End If
        End SyncLock

        Return _slDefaults
    End Function




    '=============================================================================================================
    '
    '       SQL-Ledger command file interfaces
    '
    '

    <DebuggerDisplay("cmd = {cmd}")> _
    <DebuggerDisplay("pList = {pList}")> _
    Public Class LedgerAPICmd

        Public cmd As String                        ' ap.pl, cp.pl etc
        Public pList As New List(Of DictionaryEntry)
        Public _involvedAccounts As Integer = 0     ' effectively used accounts
        Public _billedAccounts As Integer = 0       ' when an exchange rate <> 1.0 is involved, the number of billed accounts is higher
        '                                             (see how SQLLedger creates bookings in foreign currency)
        Public _isCHFCurrency As Boolean = True


        Public Sub New(ByVal cmd As String, defaultCurrency As String, ByVal login As String, ByVal pw As String)
            ' cmd:      command like ap.pl, cp.pl ect. No leading ./ !
            Me.cmd = cmd

            InitForAll(login, pw, defaultCurrency)
        End Sub


        Public Sub SetParameter(ByVal name As String, ByVal value As Object, Optional ByVal allowMultipleEntries As Boolean = False)
            ' adds the given name/value pair to the parameter list
            ' if the same parameter name already exists, the value is overwritten

            Dim existingIndex = pList.FindIndex(Function(p) p.Key = name)
            If (existingIndex >= 0) And Not allowMultipleEntries Then
                ' overwrite existing entry
                pList(existingIndex) = New DictionaryEntry With {.Key = name, .Value = value}
            Else
                pList.Add(New DictionaryEntry With {.Key = name, .Value = value})
            End If
        End Sub


        Public Function GetFullLedgerCmd() As String
            Dim pStrList As New List(Of String)
            For Each cp In pList
                If cp.Value Is Nothing Or (TypeOf cp.Value Is String AndAlso cp.Value = "") Then
                    Continue For
                End If
                Dim pValue As String = RMA2S.UrlEncode(RMA2S.CleanUpWhitespace(RMA2S.StringNothing2Empty(cp.Value)))
                pStrList.Add(String.Format("{0}={1}", cp.Key, pValue))
            Next
            Dim pArray As String() = pStrList.ToArray
            Dim pConcat = RMA2S.EasyJoin("&", pArray)

            Dim ledgerCommand As String = String.Format("./{0} ""{1}""", cmd, pConcat)
            Return ledgerCommand
        End Function


        Public Function Execute(mandant As StammRecord, ByVal app_id As String, ByVal doc_id As String, ByRef ledgerResponse As String,
                                Optional ByRef involvedAccounts As Integer = 0, Optional ByRef billedAccounts As Integer = 0, Optional ByRef billingLines As Integer = 0,
                                Optional ByVal tx As XElement = Nothing) As String
            ' calls the Ledger backend service for the prepared command

            involvedAccounts += _involvedAccounts
            billedAccounts += _billedAccounts

            ' new paramter billingLines (July 2014 bk), needed for new Billingmodel BM1407
            ' Debitoren, Kreditoren: = position count (no tax/currency dependency)
            ' Hauptbuchung: = max(involed Sollaccounts, involved Habenaccouts)
            If cmd.StartsWith("ap") Then
                billingLines += ap_position_counter
            ElseIf cmd.StartsWith("ar") Then
                billingLines += ar_position_counter
            ElseIf cmd.StartsWith("gl") Then
                billingLines += Math.Max(soll_accounts.Count, haben_accounts.Count)
            End If

            If RMA2S.CheckDEBUGState Then

                My.Computer.Clipboard.SetText(GetFullLedgerCmd())
                MsgBox(cmd & "-Buchung (-->siehe Clipboard)", MsgBoxStyle.OkOnly, "FireLedgerCmd")
                ledgerResponse = "Ok!"

                Return Nothing
            End If

            Dim ledgerError As String = Nothing
            If tx Is Nothing Then
                ' use direct backend service
                ledgerError = RMA2B.FireLedgerCmd(app_id, doc_id, cmd, PListForBackend(pList), ledgerResponse)

            Else
                ' new XML based document processing: add a <booking> section to the passed in XML container
                Dim lacXML =
                    <booking>
                        <script><%= cmd %></script>
                        <app_id><%= app_id %></app_id>
                        <doc_id><%= doc_id %></doc_id>
                        <%= If(mandant.HasSwitchedToNRLO, <is_nrlo/>, Nothing) %>
                        <involved_accounts><%= _involvedAccounts %></involved_accounts>
                        <billed_accounts><%= _billedAccounts %></billed_accounts>
                        <billing_lines><%= billingLines %></billing_lines>
                        <%= From p In PListForXML(pList)
                            Select <<%= p.Key.ToString %>><%= p.Value %></>
                        %>
                    </booking>
                tx.Add(lacXML)

                ' simulate successful SQLLedger interaction
                ledgerResponse = "Ok!"
            End If

            Return ledgerError
        End Function

        Private Function PListForBackend(ByVal pList As List(Of DictionaryEntry)) As List(Of DictionaryEntry)
            Dim backendList As New List(Of DictionaryEntry)
            For Each de In pList
                If de.Key.ToString.StartsWith("rma_") Then
                    ' ignore addtional entries
                ElseIf TypeOf de.Value Is Date Then
                    Dim d As Date = de.Value
                    backendList.Add(New DictionaryEntry With {.Key = de.Key, .Value = d.ToString("yyyyMMdd")})
                Else
                    backendList.Add(de)
                End If
            Next

            Return backendList
        End Function

        Private Function PListForXML(ByVal pList As List(Of DictionaryEntry)) As List(Of DictionaryEntry)
            Dim xmlList As New List(Of DictionaryEntry)
            For Each de In pList
                If TypeOf de.Value Is Date Then
                    Dim d As Date = de.Value
                    xmlList.Add(New DictionaryEntry With {.Key = de.Key, .Value = d.ToString("s")})

                ElseIf de.Key.ToString = "login" Or de.Key.ToString = "password" Then
                    ' xml version does not contain this confidential information

                Else
                    xmlList.Add(de)
                End If
            Next

            Return xmlList
        End Function



        ' Parameter group members:

        Public Sub InitForAll(ByVal login As String, ByVal pw As String, defaultCurrency As String)
            ' sets common items
            SetParameter("path", "bin/mozilla")
            If login IsNot Nothing Then
                SetParameter("login", login)
                SetParameter("password", pw)
            End If
            SetParameter("action", "post")
            SetParameter("precision", "2")
            SetParameter("defaultcurrency", defaultCurrency)
        End Sub

        Public Sub AddVendor(ByVal vendor As String, ByVal vendorId As Integer)
            ' adds the necessary entries
            '   - vendor
            '   - vendor_id
            '   - oldvendor
            Dim vendorStr As String
            If RMA2S.StringNothing2Empty(vendor) = "" Then
                vendorStr = vendorId
            Else
                vendorStr = String.Format("{0}--{1}", vendor, vendorId)
            End If

            SetParameter("vendor", vendorStr)
            SetParameter("vendor_id", vendorId)
            SetParameter("oldvendor", vendorStr)

            _involvedAccounts += 1
            _billedAccounts += 1
        End Sub

        Public Sub AddCustomer(ByVal customer As String, ByVal customerId As Integer)
            ' adds the necessary entries
            '   - customer
            '   - customer_id
            '   - oldcustomer
            Dim customerStr As String
            If RMA2S.StringNothing2Empty(customer) = "" Then
                customerStr = customerId
            Else
                customerStr = String.Format("{0}--{1}", customer, customerId)
            End If

            SetParameter("customer", customerStr)
            SetParameter("customer_id", customerId)
            SetParameter("oldcustomer", customerStr)

            _involvedAccounts += 1
            _billedAccounts += 1
        End Sub

        Public Sub AddCurrExchange(ByVal currency As String, ByVal exchangeRate As Double)
            ' adds the parameters
            '   - currency
            '   - exchangerate
            SetParameter("currency", currency)
            SetParameter("exchangerate", exchangeRate)
            _isCHFCurrency = (exchangeRate = 1.0)
        End Sub


        '---------------------------------
        ' AP specific parameters:

        Private ap_payment_counter = 0
        Public Sub Add_AP_Payment(ByVal amount_paid As Double, ByVal account As String, ByVal date_paid As Date, Optional ByVal exchRate As Double = 1.0)
            ' adds the parameters
            '   - paid_x
            '   - AP_paid_x
            '   - datepaid_x
            '   - exchangerate_x
            '   - paidaccounts

            ap_payment_counter += 1
            _involvedAccounts += 2       ' creates 2 journal entries
            _billedAccounts += (1 + If(_isCHFCurrency, 1, 2))

            SetParameter(String.Format("paid_{0}", ap_payment_counter), Round(amount_paid, 2))
            SetParameter(String.Format("AP_paid_{0}", ap_payment_counter), account)
            SetParameter(String.Format("datepaid_{0}", ap_payment_counter), date_paid)
            SetParameter(String.Format("exchangerate_{0}", ap_payment_counter), exchRate)
            SetParameter("paidaccounts", ap_payment_counter)
        End Sub

        Private ap_position_counter = 0
        Public Sub Add_AP_Position(ByVal amount As Double, ByVal account As String, ByVal description As String, ByVal taxAccount As String,
                                   Optional ByVal project As SLProject = Nothing)
            ' adds the parameters
            '   - amount_x
            '   - AP_amount_x
            '   - description_x
            ' [ - projectnumber_x ]
            '   - rma_taxaccount_x
            '   - rowcount

            ap_position_counter += 1
            _involvedAccounts += 1
            _billedAccounts += If(_isCHFCurrency, 1, 2)

            SetParameter(String.Format("amount_{0}", ap_position_counter), Round(amount, 2))
            SetParameter(String.Format("AP_amount_{0}", ap_position_counter), account)
            SetParameter(String.Format("description_{0}", ap_position_counter), description)
            If project IsNot Nothing Then
                SetParameter(String.Format("projectnumber_{0}", ap_position_counter), project.GetLedgerAPIString)
            End If
            SetParameter(String.Format("rma_taxaccount_{0}", ap_position_counter), taxAccount)
            SetParameter("rowcount", ap_position_counter)
        End Sub

        Private arap_tax_accounts As New HashSet(Of String)
        Public Sub Add_ARAP_Tax(ByVal taxAmount As Double, ByVal taxAccount As String, ByVal isUserValue As Boolean)
            ' adds the parameters
            '   - tax_{account}
            '   - taxaccounts
            '   - rma_customtax
            Dim tA = taxAccount.Trim
            Dim tA_Key = "tax_" & tA
            SetParameter(tA_Key, Round(taxAmount, 2))
            If taxAmount <> 0.0 Then
                _involvedAccounts += 1
                _billedAccounts += If(_isCHFCurrency, 1, 2)
            End If

            If isUserValue Then
                SetParameter("rma_customtax", taxAccount, True)
            End If

            arap_tax_accounts.Add(tA)
            Dim tA_Listing = RMA2S.EasyJoin(" ", arap_tax_accounts.ToArray)
            SetParameter("taxaccounts", tA_Listing)
        End Sub


        '---------------------------------
        ' AR specific parameters:

        Private ar_payment_counter = 0
        Public Sub Add_AR_Payment(ByVal amount_paid As Double, ByVal account As String, ByVal date_paid As Date, Optional ByVal exchRate As Double = 1.0)
            ' adds the parameters
            '   - paid_x
            '   - AR_paid_x
            '   - datepaid_x
            '   - exchangerate_x
            '   - paidaccounts

            ar_payment_counter += 1
            _involvedAccounts += 2       ' creates 2 journal entries
            _billedAccounts += (1 + If(_isCHFCurrency, 1, 2))

            SetParameter(String.Format("paid_{0}", ar_payment_counter), Round(amount_paid, 2))
            SetParameter(String.Format("AR_paid_{0}", ar_payment_counter), account)
            SetParameter(String.Format("datepaid_{0}", ar_payment_counter), date_paid)
            SetParameter(String.Format("exchangerate_{0}", ar_payment_counter), exchRate)
            SetParameter("paidaccounts", ar_payment_counter)
        End Sub

        Private ar_position_counter = 0
        Public Sub Add_AR_Position(ByVal amount As Double, ByVal account As String, ByVal description As String, ByVal taxAccount As String,
                                   Optional ByVal project As SLProject = Nothing)
            ' adds the parameters
            '   - amount_x
            '   - AR_amount_x
            '   - description_x
            ' [ - projectnumber_x ]
            '   - rma_taxaccount_x
            '   - rowcount

            ar_position_counter += 1
            _involvedAccounts += 1
            _billedAccounts += If(_isCHFCurrency, 1, 2)

            SetParameter(String.Format("amount_{0}", ar_position_counter), Round(amount, 2))
            SetParameter(String.Format("AR_amount_{0}", ar_position_counter), account)
            SetParameter(String.Format("description_{0}", ar_position_counter), description)
            If project IsNot Nothing Then
                SetParameter(String.Format("projectnumber_{0}", ar_position_counter), project.GetLedgerAPIString)
            End If
            SetParameter(String.Format("rma_taxaccount_{0}", ar_position_counter), taxAccount)
            SetParameter("rowcount", ar_position_counter)
        End Sub

        Private ar_tax_accounts As New List(Of String)
        Public Sub Add_AR_Tax(ByVal taxAmount As Double, ByVal taxAccount As String)
            ' adds the parameters
            '   - tax_{account}
            '   - taxaccounts
            Dim tA = taxAccount.Trim
            Dim tA_Key = "tax_" & tA
            If taxAmount = 0.0 Then
                SetParameter(tA_Key, "")
            Else
                SetParameter(tA_Key, Round(taxAmount, 2))
                _involvedAccounts += 1
                _billedAccounts += If(_isCHFCurrency, 1, 2)
            End If

            ar_tax_accounts.Add(tA)
            Dim tA_Listing = RMA2S.EasyJoin(" ", ar_tax_accounts.ToArray)
            SetParameter("taxaccounts", tA_Listing)
        End Sub


        '---------------------------------
        ' GL specific parameters:

        Private gl_position_counter = 0
        Private soll_accounts = New HashSet(Of String)
        Private haben_accounts = New HashSet(Of String)
        '
        Public Sub Add_GL_Position(ByVal account As String, ByVal sollAmount As Double, ByVal habenAmount As Double,
                                   Optional ByVal description As String = Nothing, Optional ByVal project As SLProject = Nothing)
            ' adds the parameters
            '   - accno_x
            '   - debit_x           (if sollAmount <> 0)
            '   - credit_x          (if habenAmount <> 0)
            '   - memo_x
            ' [ -projectnumber_x ]
            '   - rowcount

            sollAmount = Round(sollAmount, 2)
            habenAmount = Round(habenAmount, 2)
            If sollAmount = 0.0 And habenAmount = 0.0 Then
                ' can't book that..
                Return
            End If

            gl_position_counter += 1
            _involvedAccounts += 1
            _billedAccounts += If(_isCHFCurrency, 1, 2)

            SetParameter(String.Format("accno_{0}", gl_position_counter), account)

            If sollAmount <> 0 Then
                SetParameter(String.Format("debit_{0}", gl_position_counter), sollAmount)
                soll_accounts.Add(account)

            Else
                SetParameter(String.Format("credit_{0}", gl_position_counter), habenAmount)
                haben_accounts.Add(account)
            End If

            description = RMA2S.StringNothing2Empty(description).Trim
            If description.Length > 0 Then
                SetParameter(String.Format("memo_{0}", gl_position_counter), description)
            End If

            If project IsNot Nothing Then
                SetParameter(String.Format("projectnumber_{0}", ap_position_counter), project.GetLedgerAPIString)
            End If

            SetParameter("rowcount", gl_position_counter)
        End Sub


        '---------------------------------
        ' CP specific parameters:

        Public Sub InitForCP(ByVal op As SLARAPRecord, ByVal bhKonto As String)
            SetParameter("payment", "payment")

            Dim vcString As String = String.Format("{0}--{1}", op.vc_ref.name, op.vc_ref.id)
            If op.isAR Then
                ' AR specific items
                SetParameter("ARAP", "AR")
                SetParameter("arap", "ar")
                SetParameter("vc", "customer")

                SetParameter("customer_id", op.vc_ref.id)
                SetParameter("oldcustomer", vcString)
                SetParameter("selectcustomer", vcString)
                SetParameter("customer", vcString)
                SetParameter("AR_paid", bhKonto)

            Else
                ' AP specific items
                SetParameter("ARAP", "AP")
                SetParameter("arap", "ap")
                SetParameter("vc", "vendor")
                SetParameter("AP", "2000")

                SetParameter("vendor_id", op.vc_ref.id)
                SetParameter("oldvendor", vcString)
                SetParameter("selectvendor", vcString)
                SetParameter("vendor", vcString)
                SetParameter("AP_paid", bhKonto)
            End If
            _involvedAccounts += 1
            _billedAccounts += 1
        End Sub

        Private cp_position_counter = 0
        Public Sub Add_CP_Position(ByVal amount As Double, ByVal opId As Integer)
            ' adds the parameters
            '   - paid_x
            '   - id_x
            '   - rowcount

            cp_position_counter += 1
            _involvedAccounts += 1
            _billedAccounts += If(_isCHFCurrency, 1, 2)

            SetParameter(String.Format("paid_{0}", cp_position_counter), Round(amount, 2))
            SetParameter(String.Format("id_{0}", cp_position_counter), opId)
            SetParameter("rowcount", cp_position_counter)
        End Sub

    End Class





    '=============================================================================================================
    '
    '       Billing
    '
    '



    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("class = {classStr}")> _
    <DebuggerDisplay("subclass = {subclassStr}")> _
    Public Class BillingClass
        Public id As Integer
        Public classStr As String
        Public subclassStr As String

        Public ReadOnly Property classKey As String
            Get
                Return classStr & ":" & subclassStr
            End Get
        End Property
    End Class

    Private Shared _billingClassT As List(Of BillingClass) = Nothing
    Public Shared ReadOnly Property billingClassT As List(Of BillingClass)
        Get
            If _billingClassT Is Nothing Then
                _billingClassT = LoadBillingClassT()
            End If
            Return _billingClassT
        End Get
    End Property
    Private Shared Function LoadBillingClassT() As List(Of BillingClass)
        Dim bcT As New List(Of BillingClass)

        Dim sql = "select * from billing_class"
        Dim rawData = DBAccess.dbBilling.SQL2RD(sql)
        For Each rawrec In rawData
            Dim rec As New BillingClass
            rec.id = rawrec.AsInteger("id")
            rec.classStr = rawrec.AsString("class").Trim
            rec.subclassStr = rawrec.AsString("subclass").Trim
            bcT.Add(rec)
        Next
        Return bcT
    End Function



    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("billingclass_ref = {billingclass_ref}")> _
    <DebuggerDisplay("mandant_ref = {mandant_ref}")> _
    <DebuggerDisplay("cration_date = {cration_date}")> _
    <DebuggerDisplay("created_by = {created_by}")> _
    <DebuggerDisplay("billing_amount = {billing_amount}")> _
    <DebuggerDisplay("doc_id = {doc_id}")> _
    <DebuggerDisplay("doc_date = {doc_date}")> _
    <DebuggerDisplay("doc_amount = {doc_amount}")> _
    <DebuggerDisplay("doc_currency = {doc_currency}")> _
    <DebuggerDisplay("text1 = {text1}")> _
    <DebuggerDisplay("text2 = {text2}")> _
    <DebuggerDisplay("text3 = {text3}")> _
    <DebuggerDisplay("int1 = {int1}")> _
    <DebuggerDisplay("int2 = {int2}")> _
    <DebuggerDisplay("int3 = {int3}")> _
    Public Class BillingRecord
        Public id As Integer
        Public billingclass_ref As Integer
        Public mandant_ref As Integer
        Public creation_date As Date
        Public created_by As String
        Public billing_amount As Double
        Public doc_id As String
        Public doc_date As Date
        Public doc_amount As Double
        Public doc_currency As String
        Public text1 As String
        Public text2 As String
        Public text3 As String
        Public int1 As Integer
        Public int2 As Integer
        Public int3 As Integer
    End Class

    Public Shared Function LoadBillingRecords(ByVal sql As String) As List(Of BillingRecord)
        Dim brT As New List(Of BillingRecord)

        Dim rawData = DBAccess.dbBilling.SQL2RD(sql)
        For Each rawrec In rawData
            Dim rec As New BillingRecord
            rec.id = rawrec.AsInteger("id")
            rec.billingclass_ref = rawrec.AsInteger("billingclass_ref")
            rec.mandant_ref = rawrec.AsInteger("mandant_ref")
            rec.creation_date = rawrec.AsDate("creation_date")
            rec.created_by = rawrec.AsString("created_by")
            rec.billing_amount = rawrec.AsDouble("billing_amount")
            rec.doc_id = rawrec.AsString("doc_id")
            rec.doc_date = rawrec.AsDate("doc_date")
            rec.doc_amount = rawrec.AsDouble("doc_amount")
            rec.doc_currency = rawrec.AsString("doc_currency")
            rec.text1 = rawrec.AsString("text1")
            rec.text2 = rawrec.AsString("text2")
            rec.text3 = rawrec.AsString("text3")
            rec.int1 = rawrec.AsInteger("int1")
            rec.int2 = rawrec.AsInteger("int2")
            rec.int3 = rawrec.AsInteger("int3")
            brT.Add(rec)
        Next
        Return brT
    End Function



    Private Shared Sub WriteGenericBillingEntry(ByVal bc_ref As Integer, ByVal m_ref As Integer, ByVal createdBy As String, ByVal bAmount As Double,
                                                ByVal doc_id As String, ByVal doc_date As Date, ByVal doc_amount As Double, ByVal doc_curr As String,
                                                ByVal text1 As String, ByVal text2 As String, ByVal text3 As String,
                                                ByVal int1 As Integer, ByVal int2 As Integer, ByVal int3 As Integer,
                                                Optional ByVal creationDate As Date = Nothing)


        ' check references
        Dim bc_rec = billingClassT.Find(Function(_bc) _bc.id = bc_ref)
        Dim m_rec = stammT.Find(Function(_m) _m.id = m_ref)
        If bc_rec Is Nothing Then
            Throw New Exception("WriteGenericBillingEntry: Invalid bc_ref " & bc_ref)
        ElseIf m_rec Is Nothing Then
            Throw New Exception("WriteGenericBillingEntry: Invalid m_ref " & m_ref)
        End If

        ' prepare some parameters..
        Dim creaTimestampStr = "default"       ' --> default value is Now(), see definition of billing table
        If creationDate <> Nothing Then
            creaTimestampStr = PSQL_Timestamp(creationDate)
        End If
        Dim createdByStr = PSQL_String(createdBy, maxLength:=8)
        Dim bAmountStr = PSQL_Double(bAmount)
        Dim docIdStr = PSQL_String(doc_id, maxLength:=50)
        Dim docDateStr = PSQL_Date(doc_date)
        Dim docAmountStr = PSQL_Double(doc_amount)
        Dim docCurrStr = PSQL_String(doc_curr, maxLength:=3)


        If RMA2S.CheckDEBUGState() Then
            ' show debug messagee & return
            RMA2S.DoDEBUGOutput("WriteGenericBillingEntry..", String.Format("billing_class: {0}/{1}", bc_rec.classStr, bc_rec.subclassStr), "Mandant: " & m_rec.mandant)
            Return
        End If

        ' for some doc classes, duplicates are prohibited
        If (bc_rec.classStr = "payment" Or bc_rec.classStr = "kofax_doc") And (doc_id IsNot Nothing AndAlso doc_id <> "") Then
            Dim dupSql = String.Format("delete from billing where billingclass_ref = {0} and doc_id = {1}", bc_ref, PSQL_String(doc_id))
            DBAccess.dbBilling.SQLExec(dupSql)
        End If


        ' Name	            Format	    Beispiel
        ' -------------------------------------------------------------------
        ' clientname	    String	    development1
        ' document_class	String	    kofax_doc
        ' document_subclass	String	    Bankbeleg
        ' creation_date	    Timestamp	2012-06-30T12:05
        ' created_by	    String	    bk
        ' billing_amount	Double	    2.0
        ' doc_id	        String	    10002900003
        ' doc_date	        Timestamp	2012-06-30T12:05
        ' doc_amount	    Double	    1410.50
        ' doc_currency	    String	    CHF
        ' text1	            String
        ' text2	            String
        ' text3	            String
        ' int1	            Integer
        ' int2	            Integer
        ' int3	            Integer
        ' 
        ' PostBilling(parameters)


        ' build the command..
        Dim valuesStr = RMA2S.EasyJoin(", ", "default", bc_ref, m_ref, creaTimestampStr, createdByStr, bAmountStr, docIdStr, docDateStr, docAmountStr, docCurrStr, PSQL_String(text1, maxLength:=200), PSQL_String(text2, maxLength:=100), PSQL_String(text3, maxLength:=100), PSQL_Int(int1), PSQL_Int(int2), PSQL_Int(int3))
        Dim sql = String.Format("insert into billing (id, billingclass_ref, mandant_ref, creation_date, created_by, billing_amount, doc_id, doc_date, doc_amount, doc_currency, text1, text2, text3, int1, int2, int3) VALUES ({0})", valuesStr)

        Dim affectedRows = DBAccess.dbBilling.SQLExec(sql)

        If affectedRows <> 1 Then
            Throw New Exception("WriteGenericBillingEntry: insert failed!")
        End If
    End Sub



    Public Shared Sub BillingWrapper(ByVal bc As String, ByVal m_id As String,
                                     Optional ByVal createdBy As String = Nothing, Optional ByVal bAmount As Double = Double.NaN,
                                     Optional ByVal doc_id As String = Nothing, Optional ByVal doc_date As Date = Nothing, Optional ByVal doc_amount As Double = Double.NaN, Optional ByVal doc_curr As String = Nothing,
                                     Optional ByVal text1 As String = Nothing, Optional ByVal text2 As String = Nothing, Optional ByVal text3 As String = Nothing,
                                     Optional ByVal int1 As Integer = Integer.MinValue, Optional ByVal int2 As Integer = Integer.MinValue, Optional ByVal int3 As Integer = Integer.MinValue,
                                     Optional ByVal creationDate As Date = Nothing, Optional ByVal tx As XElement = Nothing)
        '
        ' bc:   may be a billing_class id or for example "kofax_doc:DMS"
        ' m:    may be a mandant id or a shortname
        ' creationDate: don't specify this field unless necessary.. it defaults to Now()

        ' find billing_class reference
        bc = RMA2S.StringNothing2Empty(bc).Trim.ToLower
        Dim bc_rec = billingClassT.Find(Function(_bc) bc = _bc.id.ToString Or bc = String.Format("{0}:{1}", _bc.classStr, _bc.subclassStr).ToLower)
        If bc_rec Is Nothing Then
            Throw New Exception("BillingWrapper: Invalid bc identifier '" & bc & "'")
        End If

        ' find mandant reference
        m_id = RMA2S.StringNothing2Empty(m_id).Trim.ToLower
        Dim m_rec = stammT.Find(Function(_m) m_id = _m.id.ToString Or m_id = _m.mandant.ToLower)
        If m_rec Is Nothing Then
            Throw New Exception("BillingWrapper: Invalid m identifier '" & m_id & "'")
        End If

        ' .. ready to bill
        If tx Is Nothing Then
            ' use direct backend service
            WriteGenericBillingEntry(bc_rec.id, m_rec.id, createdBy, bAmount, doc_id, doc_date, doc_amount, doc_curr, text1, text2, text3, int1, int2, int3, creationDate)

        Else
            ' write to XML
            Dim bXML =
                <billing>
                    <class><%= bc %></class>
                    <mandant_ref><%= m_id %></mandant_ref>
                    <created_by><%= XML_String(createdBy) %></created_by>
                    <billing_amount><%= XML_Double(bAmount) %></billing_amount>
                    <doc_id><%= XML_String(doc_id) %></doc_id>
                    <doc_date><%= XML_Date(doc_date) %></doc_date>
                    <doc_amount><%= XML_Double(doc_amount) %></doc_amount>
                    <doc_curr><%= XML_String(doc_curr) %></doc_curr>
                    <text1><%= XML_String(text1) %></text1>
                    <text2><%= XML_String(text2) %></text2>
                    <text3><%= XML_String(text3) %></text3>
                    <int1><%= XML_Int(int1) %></int1>
                    <int2><%= XML_Int(int2) %></int2>
                    <int3><%= XML_Int(int3) %></int3>
                </billing>
            tx.Add(bXML)
        End If
    End Sub



    Public Shared Function CheckBillingBarcode(ByVal barcode As String) As Boolean
        ' check if the given barcode has already been used in the billing DB
        ' works for all types of barcodes

        Dim sql = String.Format("select count(*) from billing where doc_id = {0} and " & _
                                "billingclass_ref in (select id from billing_class where class = 'kofax_doc' or id = 29)",
                                RMA2D.PSQL_String(barcode))
        Dim count = DBAccess.dbBilling.SQL2O(sql)
        Return count <> 0
    End Function


    Public Shared Function GetNextFreeUserBarcode(ByVal mandant As StammRecord) As String
        ' evaluates & returns the next free user barcode (mmmmm9xxxxx) for the given Mandant

        If mandant Is Nothing Then
            Return Nothing
        End If

        Dim sql = String.Format("select distinct doc_id from billing where " & _
                                "billingclass_ref in (select id from billing_class where class in ('kofax_doc', 'internal')) and " & _
                                "doc_id like '{0}9%'", mandant.id)
        Dim rawData = DBAccess.dbBilling.SQL2LO(sql)

        Dim smallestBarcode As Integer = Integer.MaxValue
        Dim allBarcodes As New HashSet(Of Integer)
        For Each barcodeStr In rawData
            Dim barcodeInt As Integer
            If Integer.TryParse(Right(barcodeStr, 6), barcodeInt) Then
                If (barcodeInt < 910000) And (barcodeInt < smallestBarcode) Then
                    smallestBarcode = barcodeInt
                End If
                allBarcodes.Add(barcodeInt)
            End If
        Next

        If (allBarcodes.Count = 0) Or (smallestBarcode = Integer.MaxValue) Then
            ' no user barcode assigned yet, or all barcodes >= 910000 ..
            Return mandant.id & "900001"
        End If

        ' starting with the smallest found, look for the first free barcode
        Dim nextFree As Integer = smallestBarcode
        While allBarcodes.Contains(nextFree) AndAlso nextFree < 940000
            nextFree += 1
        End While
        Return mandant.id & nextFree.ToString
    End Function


    Public Shared Function GetNextFreeUserBarcode(ByVal existingBarcode As String) As String
        ' takes an existing barcode, tries to load the associated mandant and calls GetNextFreeUserBarcode(mandant)

        If existingBarcode IsNot Nothing Then
            existingBarcode = existingBarcode.Trim
            If Regex.IsMatch(existingBarcode, "^\d{11}$") Then
                Dim mandant = stammT.Find(Function(_m) _m.id = existingBarcode.Substring(0, 5))
                Return GetNextFreeUserBarcode(mandant)
            End If
        End If

        Return Nothing
    End Function


    Public Shared Function GetNextFreeUserBarcode(ByVal mandant_id As Integer) As String
        ' returns the next free user barcode for the given mandant

        Dim mandant = stammT.Find(Function(_m) _m.id = mandant_id)
        If mandant IsNot Nothing Then
            Return GetNextFreeUserBarcode(mandant)
        End If

        Return Nothing
    End Function


    Public Shared Function GetNextFree2015Barcode(ByVal mandant_id As String, day As Date, ByRef dateAndSalt As String,
                                                  Optional createReservationRecord As Boolean = True) As String
        ' returns the next free 2015-type barcode (mmmmm--yymmdd-xxxx)

        Dim mandant = stammT.Find(Function(_m) _m.id = mandant_id)
        If mandant Is Nothing OrElse Not mandant.barcodeless Then
            Throw New Exception(String.Format("Invalid Mandant or Mandant is not enabled for 2015-type barcodes (id='{0}').", mandant_id))
        End If

        Dim yymmdd = day.ToString("yyMMdd")
        Dim sql = String.Format("select distinct doc_id from billing where " & _
                                "creation_date > '20150101' and doc_id like '{0}-{1}-____'", mandant_id, yymmdd)
        Dim rawData = DBAccess.dbBilling.SQL2LO(sql)

        Dim laufnummern As New HashSet(Of Integer)
        For Each barcodeStr In rawData
            Dim barcodeInt As Integer
            If Integer.TryParse(Right(barcodeStr, 4), barcodeInt) Then
                laufnummern.Add(barcodeInt)
            End If
        Next

        ' starting with the smallest found, look for the first free barcode
        Dim nextFree As Integer = 1
        While laufnummern.Contains(nextFree) AndAlso nextFree < 10000
            nextFree += 1
        End While

        dateAndSalt = String.Format("{0}-{1}", yymmdd, nextFree.ToString("D4"))
        Dim newBarcode = String.Format("{0}-{1}", mandant_id, dateAndSalt)

        ' create a type 27 billing record - this early registration for manual barcodes prevents the barcode from being used/created again
        If createReservationRecord Then
            BillingWrapper(27, mandant_id, doc_id:=newBarcode)
        End If

        Return newBarcode
    End Function


    Public Shared Sub RevokeReservationFor2015Barcode(ByVal barcode2015 As String)
        ' erases the type 29 record in the billing for the given barcode,
        ' only IF there is ONLY the type 29 entry.

        If barcode2015 Is Nothing OrElse Not Regex.IsMatch(barcode2015, "^\d{5}-\d{6}-\d{4}$") Then
            Return
        End If

        Dim sql = String.Format("select distinct billingclass_ref from billing where doc_id = '{0}'", barcode2015)
        Dim rawData = DBAccess.dbBilling.SQL2LO(sql)

        If rawData.Count <> 1 OrElse rawData(0) <> 29 Then
            Return
        End If

        sql = String.Format("delete from billing where doc_id = '{0}' and billingclass_ref = 29", barcode2015)
        DBAccess.dbBilling.SQLExec(sql)
    End Sub


    Public Enum RMABarcodeType
        Unknown
        PlainOld11Digits
        SPS
        Barcode2015
    End Enum

    Public Shared Function IsValidRMABarcode(ByRef barcode As String, ByRef mandant As StammRecord, ByRef barcodeType As RMABarcodeType) As Boolean

        mandant = Nothing
        barcodeType = RMABarcodeType.Unknown
        Dim mandant_id As Integer = Integer.MinValue
        '
        If barcode Is Nothing Then
            Return False

        ElseIf Regex.IsMatch(barcode, "^\d{11}$") Then
            barcodeType = RMABarcodeType.PlainOld11Digits
            mandant_id = Integer.Parse(Left(barcode, 5))

        Else
            Dim m1 = Regex.Match(barcode, "(\d{5}-\d{6}-\d{4})$")
            If m1.Success Then
                barcodeType = RMABarcodeType.Barcode2015
                barcode = m1.Groups(1).Value
                mandant_id = Integer.Parse(Left(barcode, 5))

            Else
                Dim m2 = Regex.Match(barcode, "^\w*;\w*;(\d{11})$")
                If m2.Success Then
                    barcodeType = RMABarcodeType.SPS
                    barcode = m2.Groups(1).Value
                    mandant_id = Integer.Parse(Left(barcode, 5))

                Else
                    Return False
                End If
            End If
        End If

        ' evaluate mandant
        mandant = stammT.Find(Function(_m) _m.id = mandant_id)

        Return Not (mandant Is Nothing)
    End Function


    Public Shared Function IsValidRMABarcode(barcode As String, ByRef mandant As StammRecord) As Boolean
        Dim barcodeType As RMABarcodeType
        Return IsValidRMABarcode(barcode, mandant, barcodeType)
    End Function


    Public Shared Function IsValidRMABarcode(barcode As String) As Boolean
        Dim barcodeType As RMABarcodeType
        Dim mandant As StammRecord = Nothing
        Return IsValidRMABarcode(barcode, mandant, barcodeType)
    End Function



    '=============================================================================================================
    '
    '       BankIO - Tables
    '
    '


    '
    ' Tabelle BankIO.bio_bank
    '
    '

    Public Class BIOBankRecord
        ' record data
        Public id As Integer
        Public shortname As String
        Public longname As String
        Public usescriptofid As Integer?
    End Class


    Private Shared _bioBank As List(Of BIOBankRecord) = Nothing
    Public Shared ReadOnly Property bioBank As List(Of BIOBankRecord)
        Get
            If _bioBank Is Nothing Then
                _bioBank = LoadBankData()
            End If
            Return _bioBank
        End Get
    End Property


    Private Shared Function LoadBankData() As List(Of BIOBankRecord)
        ' returns  all records from bio_bank
        Dim bT As New List(Of BIOBankRecord)

        Dim sql = "select * from bio_bank"
        Dim rawData = DBAccess.dbBankIO.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New BIOBankRecord
            rec.id = rawrec.AsInteger("id")
            rec.shortname = rawrec.AsString("shortname")
            rec.longname = rawrec.AsString("longname")
            rec.usescriptofid = rawrec.AsIntegerNIL("usescriptofid")
            bT.Add(rec)
        Next
        Return bT
    End Function



    '
    ' Tabelle BankIO.bio_account
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("bankid = {bankid}")> _
    <DebuggerDisplay("account = {account}")> _
    <DebuggerDisplay("curr = {curr}")> _
    Public Class BIOAccountRecord
        ' record data
        Public id As Integer
        Public bankid As Integer
        Public account As String
        Public curr As String
        Public datemode As Integer?
    End Class


    Private Shared _bioAccount As List(Of BIOAccountRecord) = Nothing
    Public Shared ReadOnly Property bioAccount As List(Of BIOAccountRecord)
        Get
            If _bioAccount Is Nothing Then
                _bioAccount = LoadAccountData()
            End If
            Return _bioAccount
        End Get
    End Property


    Private Shared Function LoadAccountData() As List(Of BIOAccountRecord)
        ' fills the public bioTable with all records from bio_bank
        Dim aT As New List(Of BIOAccountRecord)

        Dim sql = "select * from bio_account"
        Dim rawData = DBAccess.dbBankIO.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New BIOAccountRecord
            rec.id = rawrec.AsInteger("id")
            rec.bankid = rawrec.AsInteger("bio_bank_id")
            rec.account = rawrec.AsString("account")
            rec.curr = rawrec.AsString("currency")
            rec.datemode = rawrec.AsIntegerNIL("datemode")
            aT.Add(rec)
        Next
        Return aT
    End Function



    '
    ' Table BankIO.estv_rates
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("month = {month}")> _
    <DebuggerDisplay("currency = {currency}")> _
    <DebuggerDisplay("factor = {factor}")> _
    <DebuggerDisplay("rate = {rate}")> _
    Public Class ESTVRate
        ' record data
        Public id As Integer
        Public month As String
        Public currency As String
        Public factor As Integer
        Public rate As Double
    End Class

    Private Shared _estvRatesT As List(Of ESTVRate) = Nothing
    Public Shared ReadOnly Property estvRatesT As List(Of ESTVRate)
        Get
            If _estvRatesT Is Nothing Then
                _estvRatesT = LoadESTVRates()
            End If
            Return _estvRatesT
        End Get
    End Property

    Private Shared Function LoadESTVRates() As List(Of ESTVRate)
        Dim erT As New List(Of ESTVRate)

        Dim sql = "select * from estv_exchrate"
        Dim rawData = DBAccess.dbBankIO.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New ESTVRate
            rec.id = rawrec.AsInteger("id")
            rec.month = rawrec.AsString("month")
            rec.currency = rawrec.AsString("currency")
            rec.factor = rawrec.AsInteger("factor")
            rec.rate = rawrec.AsDouble("exchange_rate")
            erT.Add(rec)
        Next
        Return erT
    End Function

    Public Shared Function GetESTVRate_toCHF(ByVal sourceCurrency As String, ByVal ym As Date, ByRef rate As Double) As Boolean
        ' returns the rate that is necessary to convert 1 (sourceCurrency) into CHF
        ' example: 100 EUR = 100 * GetESTVRate_toCHF('EUR')
        Dim ymStr = ym.ToString("MMyyyy")
        sourceCurrency = sourceCurrency.Trim.ToUpper

        If sourceCurrency = "CHF" Or sourceCurrency = "CHW" Then
            rate = 1
            Return True
        End If

        Dim theRate = estvRatesT.Find(Function(_er) _er.month = ymStr And _er.currency = sourceCurrency)
        If theRate Is Nothing Then
            Return False
        End If

        rate = theRate.rate / theRate.factor
        Return True
    End Function



    '
    ' IExchangeRate, using table estv_exchrate
    '
    '
    ' The exchange rate table is not loaded into memory, but queried each time a value is used.
    ' However, values already found are cached in memory.


    Public Class ESTVExchangeRate
        Private Shared exchRateCache As New Dictionary(Of String, Double)

        Public Shared Function GetExchangeRate_toCHF(ByVal currency As String, ByVal xDate As Date, ByVal sellRate As Boolean, ByRef rate As Double) As Boolean
            ' the buy/sell aspect is ignored.

            ' handle trivial cases first
            If currency Is Nothing Then
                Return False
            End If
            currency = currency.Trim.ToUpper
            If currency = "CHF" Or currency = "CHW" Then        ' CHW = WIR-Bank always 1:1
                rate = 1
                Return True
            End If

            ' rate cached?
            Dim key As String = String.Format("{0}{1}", currency, xDate.ToString("yyyyMM"))
            If exchRateCache.ContainsKey(key) Then
                rate = exchRateCache(key)
                Return True
            End If

            ' value is not cached..
            If Not GetESTVRate_toCHF(currency, xDate, rate) Then
                Return False
            End If

            ' cache evaluated rate
            exchRateCache(key) = rate
            Return True
        End Function

        Public Shared Function GetExchangeRate(ByVal fromCurr As String, ByVal toCurr As String, ByVal xDate As Date, ByRef rate As Double) As Boolean
            If IsSameCurrency(fromCurr, toCurr) Then
                rate = 1.0
                Return True

            ElseIf RMA2S.IsSameCurrency(toCurr, "CHF") Then
                Return GetExchangeRate_toCHF(fromCurr, xDate, True, rate)

            Else
                ' get the two buyRates --> CHF, then divide
                Dim fromRate As Double
                Dim toRate As Double
                If GetExchangeRate_toCHF(fromCurr, xDate, True, fromRate) AndAlso GetExchangeRate_toCHF(toCurr, xDate, True, toRate) Then
                    If toRate <> 0 Then
                        rate = fromRate / toRate
                        Return True
                    End If
                End If

            End If

            rate = Double.NaN
            Return False
        End Function
    End Class



    '=============================================================================================================
    '
    '       Payment
    '
    '


    Public Enum Payment_Type
        ESR_PC          ' ESR payment to a PC account
        ESR_Bank        ' ESR payment to a bank account
        CH_PC           ' payment to Swiss PC
        CH_Bank         ' payment to Swiss bank account (non-IBAN)
        CH_IBAN         ' payment to Swiss/Liechtensteiner bank IBAN account
        INT_Bank        ' international bank payment (non-IBAN)
        INT_IBAN        ' international IBAN account
    End Enum

    Public Class PaymentType
        Public type As Payment_Type
        Public displayName As String
        Public shortDisplayName As String
        Public xmlName As String

        Public Overrides Function ToString() As String
            Return displayName
        End Function

        Public Function GetXMLId() As String
            Return xmlName
        End Function

        Public Shared Operator =(ByVal a As PaymentType, ByVal b As PaymentType) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.type = b.type
        End Operator

        Public Shared Operator <>(ByVal a As PaymentType, ByVal b As PaymentType) As Boolean
            Return Not a = b
        End Operator
    End Class

    Public Shared PaymentTypes As New List(Of PaymentType) From {
        New PaymentType With {.type = Payment_Type.ESR_PC, .displayName = "ESR auf Postkonto", .shortDisplayName = "Post-ESR", .xmlName = "esr_post"},
        New PaymentType With {.type = Payment_Type.ESR_Bank, .displayName = "ESR auf Bankkonto", .shortDisplayName = "Bank-ESR", .xmlName = "esr_bank"},
        New PaymentType With {.type = Payment_Type.CH_PC, .displayName = "Postkonto", .shortDisplayName = "Post-Konto", .xmlName = "post"},
        New PaymentType With {.type = Payment_Type.CH_Bank, .displayName = "Bankkonto einer CH/LI Bank", .shortDisplayName = "CH/LI Bank-Konto", .xmlName = "ch_bank"},
        New PaymentType With {.type = Payment_Type.CH_IBAN, .displayName = "IBAN-Konto einer CH/LI Bank", .shortDisplayName = "CH/LI IBAN-Konto", .xmlName = "ch_iban"},
        New PaymentType With {.type = Payment_Type.INT_Bank, .displayName = "Bankkonto einer ausländischen Bank", .shortDisplayName = "Intl. Bank-Konto", .xmlName = "foreign_bank"},
        New PaymentType With {.type = Payment_Type.INT_IBAN, .displayName = "IBAN-Konto einer ausländischen Bank", .shortDisplayName = "Intl. IBAN-Konto", .xmlName = "foreign_iban"}
    }


    Public Enum Payment_Spesen
        Keine
        OUR
        SHA
        BEN
    End Enum


    Public Enum MWSt_Mode
        unknown
        keine
        netto
        brutto
    End Enum

    Public Class MWStMode
        Public type As MWSt_Mode
        Public displayName As String
        Public shortDisplayName As String
        Public xmlName As String

        Public Overrides Function ToString() As String
            Return displayName
        End Function

        Public Function GetXMLId() As String
            Return xmlName
        End Function

        Public Shared Operator =(ByVal a As MWStMode, ByVal b As MWStMode) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.type = b.type
        End Operator

        Public Shared Operator <>(ByVal a As MWStMode, ByVal b As MWStMode) As Boolean
            Return Not a = b
        End Operator
    End Class

    Public Shared MWStModes As New List(Of MWStMode) From {
        New MWStMode With {.type = MWSt_Mode.keine, .displayName = "keine MWSt", .shortDisplayName = "keine", .xmlName = "none"},
        New MWStMode With {.type = MWSt_Mode.netto, .displayName = "netto (Beträge ohne MWSt)", .shortDisplayName = "netto", .xmlName = "netto"},
        New MWStMode With {.type = MWSt_Mode.brutto, .displayName = "brutto (Beträge inkl. MWSt)", .shortDisplayName = "brutto", .xmlName = "brutto"}
    }


    '
    ' Table Payment.zahlungs_wege
    '
    '

    ' beneficiary type string constants
    Public Const BeneType_LocalVendor = "localV"
    Public Const BeneType_LocalCustomer = "localC"
    Public Const BeneType_GlobalVC = "globalVC"
    Public Const BeneType_Other = "other"


    <DebuggerDisplay("id = {_id}")> _
    <DebuggerDisplay("followed_by = {followed_by}")> _
    <DebuggerDisplay("creation_date = {creation_date}")> _
    <DebuggerDisplay("validated_by = {validated_by}")> _
    <DebuggerDisplay("mandant_id = {mandant_id}")> _
    <DebuggerDisplay("beneficiary_type = {beneficiary_type}")> _
    <DebuggerDisplay("beneficiary_id = {beneficiary_id}")> _
    <DebuggerDisplay("xml = {xml}")> _
    Public Class PaymentWay
        ' record data
        Friend _id As Integer = Integer.MinValue
        Public followed_by As Integer = Integer.MinValue
        Public creation_date As Date = Nothing
        Public validated_by As String = Nothing
        Public mandant_id As Integer = Integer.MinValue
        Public beneficiary_type As String
        Public beneficiary_id As Integer
        Public xml As String = Nothing

        Public Sub New()
            ' parameter-less version for db loading
        End Sub

        Public Sub New(ByVal mandant As StammRecord, ByVal vcRecord As SLVCRecord)

            If vcRecord Is Nothing Then
                beneficiary_type = BeneType_Other

            ElseIf vcRecord.GetJoinedGlobalRecord IsNot Nothing Then
                beneficiary_type = BeneType_GlobalVC
                beneficiary_id = vcRecord.GetJoinedGlobalRecord().id

            ElseIf vcRecord.isVendor Then
                beneficiary_type = BeneType_LocalVendor
                beneficiary_id = vcRecord.id

            Else
                beneficiary_type = BeneType_LocalCustomer
                beneficiary_id = vcRecord.id
            End If

            If beneficiary_type <> BeneType_GlobalVC Then
                If mandant IsNot Nothing Then
                    mandant_id = mandant.id
                Else
                    Throw New Exception(String.Format("PaymentWay.New: Type '{0}' must have mandant.id", beneficiary_type))
                End If
            End If
        End Sub

        Public ReadOnly Property id As Integer
            Get
                Return _id
            End Get
        End Property

        Public Function Kill() As Boolean
            ' payment ways are not really erased from their table, but followed_by is set to -1, making them invisible to the LoadPaymentWays function.
            Dim sql = String.Format("update zahlungs_wege set followed_by = -1 where id = {0}", PSQL_Int(_id))
            Try
                DBAccess.dbPayment.SQL2O(sql)
            Catch ex As Exception
                Return False
            End Try

            Return True
        End Function

        Public Sub Save()
            ' store old id.. 
            Dim oldId = _id

            ' create new payment way record..
            Dim sql = String.Format("insert into zahlungs_wege (id, followed_by, creation_date, validated_by, mandant, beneficiary_type, beneficiary_id, xml) values (default, null, default, null, {0}, {1}, {2}, {3}) returning id",
                                    PSQL_Int(mandant_id), PSQL_String(beneficiary_type), PSQL_Int(beneficiary_id), PSQL_String(xml))

            Dim newId As Object = Nothing
            newId = DBAccess.dbPayment.SQL2O(sql)
            If newId Is Nothing Then
                ' insert failed
                Throw New Exception("PaymentWay.Save: Creation of new record failed.")
            End If

            ' store new id & update old record
            _id = newId
            If oldId > 0 Then
                sql = String.Format("update zahlungs_wege set followed_by = {0} where id = {1}", PSQL_Int(_id), PSQL_Int(oldId))
                Try
                    ' assumed to work.. :)
                    DBAccess.dbPayment.SQL2O(sql)
                Catch ex As Exception
                End Try
            End If
        End Sub

        Public Overrides Function ToString() As String
            ' return a display version of this payment way

            Dim xDoc As XDocument = Nothing
            Try
                xDoc = XDocument.Parse(xml)
            Catch ex As Exception
            End Try

            If xDoc Is Nothing Then
                Return "(nicht konfiguriert)"
            End If

            Dim ptXMLName = xDoc...<paymentType>.Value
            Dim account = StringNothing2Empty(xDoc...<account>.Value)
            Dim zg1 = xDoc...<zg1>.Value
            Dim for1 = StringNothing2Empty(xDoc...<for1>.Value)

            Dim paymentType = PaymentTypes.Find(Function(_pt) _pt.xmlName = ptXMLName)

            Dim dStr As String = ""
            If zg1 Is Nothing Then
                dStr = for1 & ", "
            Else
                dStr = StringNothing2Empty(zg1) & ", "
            End If

            If paymentType Is Nothing Then
                dStr &= "(kein Typ), "
            Else
                dStr &= paymentType.shortDisplayName & ", "
            End If

            dStr &= account

            Return dStr
        End Function

        Public Shared Operator =(ByVal a As PaymentWay, ByVal b As PaymentWay)
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As PaymentWay, ByVal b As PaymentWay)
            Return Not a = b
        End Operator

    End Class


    Public Shared Function LoadPaymentWay(ByVal pwId As Integer) As PaymentWay

        ' create sql query
        Dim sql = String.Format("select * from zahlungs_wege where id = {0}", PSQL_Int(pwId))
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

        Dim rec As PaymentWay = Nothing
        If rawData.Count = 1 Then
            rec = New PaymentWay
            Dim rawrec = rawData(0)

            rec._id = rawrec.AsInteger("id")
            rec.followed_by = rawrec.AsInteger("followed_by", True)
            rec.creation_date = rawrec.AsDate("creation_date")
            rec.validated_by = rawrec.AsString("validated_by")
            rec.mandant_id = rawrec.AsInteger("mandant", True)
            rec.beneficiary_type = rawrec.AsString("beneficiary_type")
            rec.beneficiary_id = rawrec.AsInteger("beneficiary_id")
            rec.xml = rawrec.AsString("xml")
        End If

        Return rec
    End Function


    Public Shared Function LoadPaymentWays(ByVal _mandant As StammRecord, ByVal vcRecord As SLVCRecord) As List(Of PaymentWay)
        Dim pws As New List(Of PaymentWay)

        ' create sql query
        Dim sql As String
        Dim sql_root = "select * from zahlungs_wege where followed_by is null and "

        If vcRecord Is Nothing Then
            ' show 'others'
            ' this type has no target reference id - load ALL 'other' payment ways of this mandant
            sql = sql_root & String.Format(" (beneficiary_type = {0} and mandant = {1})", PSQL_String(BeneType_Other), PSQL_Int(_mandant.id))

        ElseIf vcRecord.GetJoinedGlobalRecord() IsNot Nothing Then
            ' show pws of joined global VC
            Dim gvc = vcRecord.GetJoinedGlobalRecord()
            sql = sql_root & String.Format(" (beneficiary_type = {0} and beneficiary_id = {1})", PSQL_String(BeneType_GlobalVC), gvc.id)

        ElseIf vcRecord.isVendor Then
            ' local vendor
            sql = sql_root & String.Format(" (beneficiary_type = {0} and mandant = {1} and beneficiary_id = {2})", PSQL_String(BeneType_LocalVendor), PSQL_Int(_mandant.id), PSQL_Int(vcRecord.id))

        Else
            ' local customer
            sql = sql_root & String.Format(" (beneficiary_type = {0} and mandant = {1} and beneficiary_id = {2})", PSQL_String(BeneType_LocalCustomer), PSQL_Int(_mandant.id), PSQL_Int(vcRecord.id))
        End If


        ' load resulting payment ways
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)
        For Each rawrec In rawData
            Dim rec As New PaymentWay
            rec._id = rawrec.AsInteger("id")
            rec.followed_by = rawrec.AsInteger("followed_by", True)
            rec.creation_date = rawrec.AsDate("creation_date")
            rec.validated_by = rawrec.AsString("validated_by")
            rec.mandant_id = rawrec.AsInteger("mandant", True)
            rec.beneficiary_type = rawrec.AsString("beneficiary_type")
            rec.beneficiary_id = rawrec.AsInteger("beneficiary_id")
            rec.xml = rawrec.AsString("xml")

            pws.Add(rec)
        Next

        Return pws
    End Function



    '
    ' Table Payment.plz
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("plz = {plz}")> _
    <DebuggerDisplay("subplz = {subplz}")> _
    <DebuggerDisplay("name = {name}")> _
    <DebuggerDisplay("longname = {longname}")> _
    <DebuggerDisplay("kanton = {kanton}")> _
    Public Class OrtPLZ
        ' record data
        Public id As Integer
        Public plz As String
        Public subplz As String
        Public name As String
        Public longname As String
        Public kanton As String
        ' only filled by GetOrtPLZ4Regex:
        Public regexName As String = Nothing
    End Class

    Private Shared _ortplzT As List(Of OrtPLZ) = Nothing
    Public Shared ReadOnly Property ortplzT As List(Of OrtPLZ)
        Get
            If _ortplzT Is Nothing Then
                _ortplzT = LoadOrtPLZ()
            End If
            Return _ortplzT
        End Get
    End Property

    Private Shared Function LoadOrtPLZ() As List(Of OrtPLZ)
        Dim opT As New List(Of OrtPLZ)

        Dim sql = "select * from plz"
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New OrtPLZ
            rec.id = rawrec.AsInteger("id")
            rec.plz = rawrec.AsString("plz")
            rec.subplz = rawrec.AsString("subplz")
            rec.name = rawrec.AsString("name")
            rec.longname = rawrec.AsString("longname")
            rec.kanton = rawrec.AsString("kanton")
            opT.Add(rec)
        Next
        Return opT
    End Function

    Public Shared Function GetOrtByPLZ(ByVal plz As String) As String
        Dim orte = ortplzT.FindAll(Function(_o) _o.plz = plz)
        If orte.Count = 0 Then
            Return Nothing
        End If

        orte.Sort(Function(o1, o2) o1.subplz < o2.subplz)
        Return orte(0).name
    End Function

    Private Shared _GetOrtPLZ4Regex As List(Of OrtPLZ) = Nothing
    Public Shared Function GetOrtPLZ4Regex() As List(Of OrtPLZ)
        ' returns a list of all OrtPLZ records with subplz = '00'
        ' regexName is filled with a string suitable for regex search (kinda fuzzy)

        If _GetOrtPLZ4Regex Is Nothing Then
            _GetOrtPLZ4Regex = (From _o In ortplzT Where _o.subplz = "00" Select _o).ToList
            For Each _o In _GetOrtPLZ4Regex
                Dim ms = Regex.Matches(_o.name.ToLower, "\w{3,}")
                For Each _m As Match In ms
                    If _m.Value <> "les" Then
                        _o.regexName = _m.Value
                        Continue For
                    End If
                Next
                If _o.regexName Is Nothing Then
                    _o.regexName = NoWhitespace(_o.name.ToLower)
                End If
            Next
        End If

        Return _GetOrtPLZ4Regex
    End Function


    '
    ' Table Payment.six_bankenstamm
    '
    ' contains all CH bank institutes

    <DebuggerDisplay("id = {id}")> _
    Public Class SIXBankStamm
        ' record data
        Public id As Integer
        Public gruppe As String
        Public bcnr As String
        Public filialId As String
        Public bcNrNeu As String
        Public sicNr As String
        Public hauptsitz As String
        Public bcArt As Char
        Public validFrom As Date
        Public sic As Char
        Public euroSic As Char
        Public sprache As Char
        Public kurzbez As String
        Public bank As String
        Public domizil As String
        Public adresse As String
        Public plz As String
        Public ort As String
        Public telefon As String
        Public fax As String
        Public vorwahl As String
        Public landcode As String
        Public postkonto As String
        Public swift As String
    End Class

    Private Shared Function LoadSIXBankStamm(Optional ByVal whereClause As String = Nothing) As List(Of SIXBankStamm)
        ' whereClause example:
        ' whereClause = "where ""BCNr"" = '700'"
        Dim sbT As New List(Of SIXBankStamm)

        Dim sql = "select * from six_bankenstamm"
        If whereClause IsNot Nothing Then
            sql &= " " & whereClause
        End If
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New SIXBankStamm
            rec.id = rawrec.AsInteger("id")
            rec.gruppe = rawrec.AsString("Gruppe")
            rec.bcnr = rawrec.AsString("BCNr")
            rec.filialId = rawrec.AsString("Filial-ID")
            rec.bcNrNeu = rawrec.AsString("BCNr neu")
            rec.sicNr = rawrec.AsString("SIC-Nr")
            rec.hauptsitz = rawrec.AsString("Hauptsitz")
            rec.bcArt = rawrec.AsChar("BC-Art")
            rec.validFrom = rawrec.AsDate("gültig ab")
            rec.sic = rawrec.AsChar("SIC")
            rec.euroSic = rawrec.AsChar("euroSIC")
            rec.sprache = rawrec.AsChar("Sprache")
            rec.kurzbez = rawrec.AsString("Kurzbez")
            rec.bank = rawrec.AsString("Bank")
            rec.domizil = rawrec.AsString("Domizil")
            rec.adresse = rawrec.AsString("Postadresse")
            rec.plz = rawrec.AsString("PLZ")
            rec.ort = rawrec.AsString("Ort")
            rec.telefon = rawrec.AsString("Telefon")
            rec.fax = rawrec.AsString("Fax")
            rec.vorwahl = rawrec.AsString("Vorwahl")
            rec.landcode = rawrec.AsString("Landcode")
            rec.postkonto = rawrec.AsString("Postkonto")
            rec.swift = rawrec.AsString("SWIFT")
            sbT.Add(rec)
        Next
        Return sbT
    End Function


    Public Shared Function GetSIXBankStammByBCNr(ByVal bcnr As String) As SIXBankStamm
        ' finds the bank record for a given bc (Bankleitzahl BLZ)

        If bcnr IsNot Nothing Then
            Dim m = Regex.Match(bcnr.Trim, "^0*(\d+)$")
            If m.Success Then
                Dim whereClause = String.Format("where ""BCNr"" = {0}", PSQL_String(m.Groups(1).Value))
                Dim sixes = LoadSIXBankStamm(whereClause)
                If sixes.Count > 0 Then
                    sixes.Sort(Function(s1, s2) s1.filialId < s2.filialId)
                    Return sixes(0)
                End If
            End If
        End If

        Return Nothing
    End Function


    Public Shared Function GetSIXBankStammByPC(ByVal pc As String) As SIXBankStamm
        ' finds the bank record for a given transaction account (pc), like '80-2-2' --> UBS

        If Not RMA2S.CheckPCNumber(pc) Then
            ' invalid pc..
            Return Nothing
        End If

        Dim bcnr = GetBCNRfromPCESR(pc)
        Return GetSIXBankStammByBCNr(bcnr)
    End Function


    Public Shared Function GetSIXBankStammByAccount(ByVal account As String) As SIXBankStamm
        ' finds the bank record for a given account number (iban, pc or esr transaction pc)

        Dim sbs = GetSIXBankStamm4CHIBAN(account)
        If sbs Is Nothing Then
            sbs = GetSIXBankStammByPC(account)
        End If
        Return sbs
    End Function



    '
    ' Table Payment.vendors_ch_stamm
    '
    '   --> this table should be renamed because it is used for vendors AND customers
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("uid = {uid}")> _
    <DebuggerDisplay("mwstnr = {mwstnr}")> _
    <DebuggerDisplay("name = {name}")> _
    <DebuggerDisplay("addl_name = {addl_name}")> _
    <DebuggerDisplay("co = {co}")> _
    <DebuggerDisplay("address = {address}")> _
    <DebuggerDisplay("number = {number}")> _
    <DebuggerDisplay("plz = {plz}")> _
    <DebuggerDisplay("city = {city}")> _
    <DebuggerDisplay("sl_mandant = {sl_mandant}")> _
    <DebuggerDisplay("zahlungsfrist = {zahlungsfrist}")> _
    Public Class VC_CH_Stamm
        ' record data
        Public id As Integer
        Public uid As String
        Public mwstnr As String
        Public name As String
        Public addl_name As String
        Public co As String
        Public address As String
        Public number As String
        Public plz As String
        Public city As String
        Public sl_mandant As String
        Public zahlungsfrist As Integer

        Public Overrides Function ToString() As String
            Return String.Format("{0}, {1} {2}", name, plz, city)
        End Function

        Public Function GetXMLId() As String
            Return String.Format("VC_Stamm¬{0} ({1})", id, name)
        End Function

        Public ReadOnly Property NormalizedSearchStr() As String
            Get
                Dim nss = String.Format("{0}¬{1}¬{2}¬{3}¬{4}¬{5}¬{6}", name, plz, city, uid, mwstnr, address, sl_mandant)
                Return NoWhitespace(RemoveDiacritics(nss)).ToLower
            End Get
        End Property

        Public Shared Operator =(ByVal a As VC_CH_Stamm, ByVal b As VC_CH_Stamm) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As VC_CH_Stamm, ByVal b As VC_CH_Stamm) As Boolean
            Return Not a = b
        End Operator
    End Class

    Private Shared _vcCHStammT As New List(Of VC_CH_Stamm)
    Public Shared ReadOnly Property vcCHStammT As List(Of VC_CH_Stamm)
        Get
            SyncLock _vcCHStammT
                If MustReload("VCCHStamm") Then
                    _vcCHStammT.Clear()
                    _vcCHStammT.AddRange(LoadVCCHStamm())
                End If
            End SyncLock
            Return _vcCHStammT
        End Get
    End Property

    Public Shared Sub ReLoadVCCHStamm()
        Reload("VCCHStamm")
    End Sub

    Private Shared Function LoadVCCHStamm() As List(Of VC_CH_Stamm)
        Dim vcsT As New List(Of VC_CH_Stamm)

        Dim sql = "select * from vendors_ch_stamm"
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New VC_CH_Stamm
            rec.id = rawrec.AsInteger("id")
            rec.uid = rawrec.AsString("uid")
            rec.mwstnr = rawrec.AsString("mwstnr")
            rec.name = rawrec.AsString("name")
            rec.addl_name = rawrec.AsString("addl_name")
            rec.co = rawrec.AsString("co")
            rec.address = rawrec.AsString("address")
            rec.number = rawrec.AsString("number")
            rec.plz = rawrec.AsString("plz")
            rec.city = rawrec.AsString("city")
            rec.sl_mandant = rawrec.AsString("sl_mandant")
            rec.zahlungsfrist = rawrec.AsInteger("zahlungsfrist")

            vcsT.Add(rec)
        Next
        Return vcsT
    End Function


    Public Shared Function VCCHStamm_CreateNewEntry(ByRef newEntry As VC_CH_Stamm) As String
        ' creates a new entry in payment.vendors_ch_stamm
        ' on success (Nothing returned), the table is reloaded & newEntry is set to the new table entry
        ' on failure, the exception text is returned

        Dim sql = "insert into vendors_ch_stamm (id, uid, mwstnr, name, addl_name, co, address, number, plz, city, sl_mandant, zahlungsfrist) " & _
            "values (default, " & _
                     String.Format("{0}, {1}, {2}, ", PSQL_String(newEntry.uid), PSQL_String(newEntry.mwstnr), PSQL_String(newEntry.name)) & _
                     String.Format("{0}, {1}, {2}, ", PSQL_String(newEntry.addl_name), PSQL_String(newEntry.co), PSQL_String(newEntry.address)) & _
                     String.Format("{0}, {1}, {2}, ", PSQL_String(newEntry.number), PSQL_String(newEntry.plz), PSQL_String(newEntry.city)) & _
                     String.Format("{0}, {1}", PSQL_String(newEntry.sl_mandant), newEntry.zahlungsfrist) & _
            ") returning id"

        Dim newId As Object = Nothing
        Try
            newId = DBAccess.dbPayment.SQL2O(sql)
            If newId Is Nothing Then
                ' insert failed
                Return "RMA2D.VCCHStamm_CreateNewEntry: 'insert into vendors_ch_stamm' failed."
            End If
        Catch ex As Exception
            Return ex.Message
        End Try

        ' reload table containing the new record and update newEntry parameter
        ReLoadVCCHStamm()
        newEntry = vcCHStammT.Find(Function(_chV) _chV.id = newId)
        If newEntry Is Nothing Then
            Return "RMA2D.VCCHStamm_CreateNewEntry: created entry not found."
        End If

        Return Nothing
    End Function


    Public Shared Function VCCHStamm_UpdateEntry(ByVal vc As VC_CH_Stamm) As String
        ' updates the given entry in payment.vendors_ch_stamm
        ' on success (Nothing returned), the table is reloaded
        ' on failure, the exception text is returned

        Dim sql = "update vendors_ch_stamm set " & _
                    String.Format("uid = {0}, mwstnr = {1}, name = {2}, ", PSQL_String(vc.uid), PSQL_String(vc.mwstnr), PSQL_String(vc.name)) & _
                    String.Format("addl_name = {0}, co = {1}, address = {2}, ", PSQL_String(vc.addl_name), PSQL_String(vc.co), PSQL_String(vc.address)) & _
                    String.Format("number = {0}, plz = {1}, city = {2}, ", PSQL_String(vc.number), PSQL_String(vc.plz), PSQL_String(vc.city)) & _
                    String.Format("sl_mandant = {0}, zahlungsfrist = {1}", PSQL_String(vc.sl_mandant), vc.zahlungsfrist) & _
                  " where id = " & vc.id

        Dim newId As Object = Nothing
        Try
            DBAccess.dbPayment.SQL2O(sql)
        Catch ex As Exception
            Return ex.Message
        End Try

        ' reload table containing the new record and update newEntry parameter
        ReLoadVCCHStamm()
        Return Nothing
    End Function



    '
    ' Table Payment.xref_localvendor_chvendor
    '
    '

    Private Shared lvxref_mandant As String = Nothing
    Private Shared lvxref_l2g As Dictionary(Of Integer, Integer) = Nothing
    Private Shared lvxref_g2l As Dictionary(Of Integer, Integer) = Nothing
    '
    Public Shared Sub ReloadVendorXREFDicts()
        lvxref_l2g = Nothing
        lvxref_g2l = Nothing
    End Sub
    '
    Public Shared Function LoadVendorXREFDicts(ByVal mandant As String, ByRef l2g As Dictionary(Of Integer, Integer), ByRef g2l As Dictionary(Of Integer, Integer),
                                               Optional ByVal reload As Boolean = False) As String
        ' loads the joined global vendor ids for all local vendors of the given mandant
        ' fills 2 dictionaries: local --> global ids, global --> local ids
        ' returns Nothing on success, the exception message on failure

        If Not reload AndAlso (lvxref_l2g IsNot Nothing And lvxref_g2l IsNot Nothing And mandant = lvxref_mandant) Then
            l2g = lvxref_l2g
            g2l = lvxref_g2l
            Return Nothing
        End If
        l2g = New Dictionary(Of Integer, Integer)
        g2l = New Dictionary(Of Integer, Integer)

        Dim localIds As New List(Of Object)
        For Each lv In slVendors.FindAll(Function(_lv) _lv.isActive)
            localIds.Add(lv.id)
        Next
        If localIds.Count() = 0 Then
            Return Nothing
        End If

        Dim sql = String.Format("select local_vendorid, vendors_ch_stamm_id from rma.xref_localvendor_chvendor " & _
                                "where mandant = '{0}' and local_vendorid in ({1})", mandant, RMA2S.EasyJoinA(", ", localIds.ToArray))

        Dim resultList As List(Of Object())
        Try
            resultList = DBAccess.dbPayment.SQL2LAO(sql)
        Catch ex As Exception
            Return "RMA2D.LoadVendorXREFDicts: " & ex.Message
        End Try

        If resultList IsNot Nothing Then
            For Each rec In resultList
                l2g(rec(0)) = rec(1)
                g2l(rec(1)) = rec(0)
            Next
        End If

        lvxref_mandant = mandant
        lvxref_l2g = l2g
        lvxref_g2l = g2l

        Return Nothing
    End Function



    '
    ' Table Payment.xref_customer_l2g
    '
    '

    Private Shared lcxref_mandant As String = Nothing
    Private Shared lcxref_l2g As Dictionary(Of Integer, Integer) = Nothing
    Private Shared lcxref_g2l As Dictionary(Of Integer, Integer) = Nothing
    Private Shared lcxref_isSwiss As Dictionary(Of Integer, Boolean) = Nothing
    '
    Public Shared Function LoadCustomerXREFDicts(ByVal mandant As String, ByRef l2g As Dictionary(Of Integer, Integer), ByRef g2l As Dictionary(Of Integer, Integer),
                                                 ByRef isSwiss As Dictionary(Of Integer, Boolean), Optional ByVal reload As Boolean = False) As String
        ' loads the joined global customer ids for all local customer of the given mandant
        ' fills 2 dictionaries: local --> global ids, global --> local ids
        ' returns Nothing on success, the exception message on failure

        If reload OrElse lcxref_l2g Is Nothing OrElse mandant <> lcxref_mandant Then
            Dim sql = String.Format("select * from rma.xref_customer_l2g where mandant = '{0}'", mandant)
            Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

            lcxref_l2g = New Dictionary(Of Integer, Integer)
            lcxref_g2l = New Dictionary(Of Integer, Integer)
            lcxref_isSwiss = New Dictionary(Of Integer, Boolean)
            For Each rec In rawData
                Dim l_id = rec.AsInteger("local_customerid")
                Dim g_id = rec.AsInteger("global_customerid", nullAsMinValue:=True)
                Dim customerIsSwiss = rec.AsBooleanNIL("nrlo_is_swiss")

                If g_id <> Integer.MinValue Then
                    lcxref_l2g(l_id) = g_id
                    lcxref_g2l(g_id) = l_id
                End If

                If customerIsSwiss.HasValue Then
                    lcxref_isSwiss(l_id) = customerIsSwiss.Value
                End If
            Next

            lcxref_mandant = mandant
        End If

        l2g = lcxref_l2g
        g2l = lcxref_g2l
        isSwiss = lcxref_isSwiss

        Return Nothing
    End Function

    Public Shared Sub SetCustomerIsSwiss(mandant As String, debitor As SLVCRecord, isSwiss As Boolean?)
        Dim sql = String.Format("update rma.xref_customer_l2g set nrlo_is_swiss = {0} where mandant = '{1}' and local_customerid = {2}",
                                PSQL_BooleanNIL(isSwiss), mandant, debitor.id)
        Dim affected = DBAccess.dbPayment.SQLExec(sql)

        If affected = 0 Then
            sql = String.Format("insert into xref_customer_l2g (id, mandant, local_customerid, global_customerid, nrlo_is_swiss) values (default, {0}, {1}, null, {2}) returning id",
                                    PSQL_String(mandant), debitor.id, PSQL_BooleanNIL(isSwiss))
            DBAccess.dbPayment.SQLExec(sql)
        End If

        If isSwiss.HasValue Then
            lcxref_isSwiss(debitor.id) = isSwiss.Value
        Else
            lcxref_isSwiss.Remove(debitor.id)
        End If
        debitor.nrlo_Is_Swiss = isSwiss
    End Sub



    '
    ' Table Payment.xref_pcesr_bcnr
    '
    '

    Public Shared Function GetBCNRfromPCESR(ByVal pcesr As String) As String

        If Not RMA2S.CheckPCNumber(pcesr) Then
            ' pc account seems to be invalid
            Return Nothing
        End If
        pcesr = Regex.Replace(pcesr, "\D", "")

        Dim sql = String.Format("select * from xref_pcesr_bcnr where pcesr = {0}", PSQL_String(pcesr))
        Dim rawData = DBAccess.dbPayment.SQL2RD(sql)

        If rawData.Count > 0 Then
            Return rawData(0).AsString("bcnr")
        End If

        Return Nothing
    End Function


    '
    ' Table Payment.swift_stamm
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("swift = {swift}")> _
    <DebuggerDisplay("addr1 = {addr1}")> _
    <DebuggerDisplay("addr2 = {addr2}")> _
    <DebuggerDisplay("addr3 = {addr3}")> _
    Public Class SwiftStamm
        ' record data
        Public id As Integer
        Public swift As String
        Public addr1 As String
        Public addr2 As String
        Public addr3 As String
    End Class

    Public Shared Function GetBankAddrBySWIFT(ByRef swift As String) As SwiftStamm
        ' loads the bank address for a given SWIFT code
        ' returns Nothing if not found/error

        If Not RMA2S.CheckSWIFTBICFormat(swift) Then
            Return Nothing
        End If

        Try
            Dim sql = String.Format("select * from rma.swift_stamm where swift = '{0}'", swift)
            Dim rawData = DBAccess.dbPayment.SQL2RD(sql)
            If rawData.Count = 0 Then
                Dim swiftXXX = swift.Substring(0, 8) & "XXX"
                sql = String.Format("select * from rma.swift_stamm where swift = '{0}'", swiftXXX)
                rawData = DBAccess.dbPayment.SQL2RD(sql)
            End If

            If rawData.Count > 0 Then
                Dim rawrec = rawData(0)
                Dim rec As New SwiftStamm
                rec.id = rawrec.AsInteger("id")
                rec.swift = rawrec.AsString("swift")
                swift = rec.swift
                rec.addr1 = rawrec.AsString("addr1")
                rec.addr2 = rawrec.AsString("addr2")
                rec.addr3 = rawrec.AsString("addr3")
                Return rec
            End If

        Catch ex As Exception
        End Try

        Return Nothing
    End Function




    '=============================================================================================================
    '
    '       Docca
    '
    '


    '
    ' qi_vendor (qualified items for vendor recognition)
    '
    '

    <DebuggerDisplay("mandant = {mandant}")> _
    <DebuggerDisplay("key = {key}")> _
    <DebuggerDisplay("targetId = {targetId}")> _
    Public Class QualifiedItem
        Public id As Integer = -1
        Public mandant As String = Nothing
        Public key As String = Nothing
        Public targetId As Integer = -1
        Public reliability As Integer
    End Class

    Public Shared Function qi_vendor(ByVal mandant As StammRecord, ByVal keys As List(Of String)) As List(Of QualifiedItem)
        Dim keyList = "'" & EasyJoinIf("', '", keys.ToArray) & "'"
        Dim sql As String
        If mandant Is Nothing Then
            sql = String.Format("Select * from qi_vendor where key in ({0})", keyList)
        Else
            sql = String.Format("Select * from qi_vendor where key in ({0}) and (mandant is NULL or mandant = {1})", keyList, PSQL_String(mandant.mandant))
        End If
        Dim rawData = DBAccess.dbDocca.SQL2RD(sql)

        Dim qil As New List(Of QualifiedItem)
        For Each rawrec In rawData
            Dim rec As New QualifiedItem
            rec.id = rawrec.AsInteger("id")
            rec.mandant = StringEmpty2Nothing(rawrec.AsString("mandant"))
            rec.key = rawrec.AsString("key")
            rec.targetId = rawrec.AsInteger("vendorid")
            rec.reliability = rawrec.AsInteger("reliability")
            qil.Add(rec)
        Next
        Return qil
    End Function


    Public Shared Function Get_All_QI_Vendor() As List(Of QualifiedItem)
        Dim result As New List(Of QualifiedItem)

        Dim sql = "Select * from qi_vendor"
        Dim rawData = DBAccess.dbDocca.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New QualifiedItem
            rec.id = rawrec.AsInteger("id")
            Dim _mandant = rawrec.AsString("mandant")
            If _mandant IsNot Nothing AndAlso _mandant.Length > 0 Then
                rec.mandant = _mandant
            End If
            rec.key = rawrec.AsString("key")
            rec.targetId = rawrec.AsInteger("vendorid")
            rec.reliability = rawrec.AsInteger("reliability")
            result.Add(rec)
        Next

        Return result
    End Function


    Public Shared Function AddUpdateDelete_QI_Vendor(ByVal qi As QualifiedItem, Optional ByVal deleteThisQI As Boolean = False) As Boolean

        If qi Is Nothing Then
            Return True
        End If

        Dim sql As String
        With qi
            If deleteThisQI Then
                sql = String.Format("delete from qi_vendor where id = {0}", .id)

            ElseIf .id <> -1 Then
                sql = String.Format("update qi_vendor set reliability = {0} where id = {1}", .reliability, .id)

            ElseIf .mandant Is Nothing Then
                sql = String.Format("insert into qi_vendor values (default, NULL, {0}, {1}, {2})",
                                            PSQL_String(.key), PSQL_Int(.targetId), PSQL_Int(.reliability))
            Else
                sql = String.Format("insert into qi_vendor values (default, {0}, {1}, {2}, {3})",
                                            PSQL_String(.mandant), PSQL_String(.key), PSQL_Int(.targetId), PSQL_Int(.reliability))
            End If
        End With
        Dim rowsAffected = DBAccess.dbDocca.SQLExec(sql)

        Return (rowsAffected = 1)
    End Function


    '=============================================================================================================
    '
    '       misc tables
    '
    '

    Public Class Misc_AppValue
        Public id As Integer
        Public app_id As String
        Public key As String
        Public value As String
    End Class


    Public Shared Function Misc_GetAppValues(app_id As String, Optional key As String = Nothing, Optional id As Integer = Integer.MinValue) As List(Of Misc_AppValue)
        Dim sql As String = "select * from app_values where "
        If id > 0 Then
            sql &= String.Format("id = {0}", id)

        Else
            sql &= String.Format("app_id = {0} {1}",
                                  PSQL_String(app_id),
                                  If(key Is Nothing, "", String.Format("and key = {0}", PSQL_String(key))))
        End If
        sql &= " order by id desc"

        Dim rawData = DBAccess.dbMisc.SQL2RD(sql)

        Dim result As New List(Of Misc_AppValue)
        For Each rawrec In rawData
            Dim rec As New Misc_AppValue
            rec.id = rawrec.AsInteger("id")
            rec.app_id = rawrec.AsString("app_id")
            rec.key = rawrec.AsString("key")
            rec.value = rawrec.AsString("value")
            result.Add(rec)
        Next

        Return result
    End Function

    Public Shared Function Misc_GetAppValue(app_id As String, Optional key As String = Nothing, Optional id As Integer = Integer.MinValue) As Misc_AppValue
        Dim values = Misc_GetAppValues(app_id, key, id)
        If values.Count = 0 Then
            Return Nothing
        End If
        Return values(0)
    End Function

    Public Shared Function Misc_WriteAppValue(app_id As String, key As String, value As String, Optional updateLastExisting As Boolean = False) As Integer

        If updateLastExisting Then
            Dim lastExistingRecord = Misc_GetAppValue(app_id, key)
            If lastExistingRecord IsNot Nothing Then
                Misc_OverwriteAppValue(lastExistingRecord.id, value)
                Return lastExistingRecord.id
            End If
        End If

        Dim sql = String.Format("insert into app_values (id, app_id, key, value) values (default, {0}, {1}, {2}) returning id",
                                PSQL_String(app_id), PSQL_String(key), PSQL_String(value))

        Dim newId As Object = DBAccess.dbMisc.SQL2O(sql)
        If newId Is Nothing Then
            ' insert failed
            Return Integer.MinValue
        End If

        Return CInt(newId)

    End Function

    Public Shared Sub Misc_OverwriteAppValue(overwriteId As Integer, value As String)
        Dim sql = String.Format("update app_values set value = {0} where id = {1}", PSQL_String(value), overwriteId)
        DBAccess.dbMisc.SQLExec(sql)
    End Sub





    '=============================================================================================================
    '
    '       Doc
    '
    '


    '
    ' L0 tables
    '
    '

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("status = {status}")> _
    <DebuggerDisplay("path = {path}")> _
    <DebuggerDisplay("checkout_time = {checkout_time}")> _
    Public Class L0_Attachment
        ' represents a single file attachment to the master document l0_doc
        ' path: the relative path in L0 and L1, where
        '   L0\path contains raw level 0 files (as received by Email/Dropbox etc)
        '   L1\path contains the files l1_000000001.PDF, l1_000000001.TIF and l1_000000001.XML (may change), where 1 is the id of this record
        Public id As Integer
        Public l0_doc As L0_Doc
        Public status As String
        Public path As String
        Public checkout_time As Date

        Public Function GetPathAndFileRoot(Optional ByVal ext As String = Nothing)
            Dim root = String.Format("{0}\l1_{1:D9}", path, id)
            If ext IsNot Nothing Then
                ext = Regex.Replace(ext, "\W", "")
                root &= "." & ext
            End If
            Return root
        End Function

    End Class

    Public Shared Function LoadL0Attachments(Optional ByVal idList As List(Of Integer) = Nothing) As List(Of L0_Attachment)
        Dim sql As String = Nothing
        If idList Is Nothing Then
            ' load all attachments in the 'ready' state
            sql = "select * from l0_doc_attachment where status = 'ready' order by id"
        Else
            ' load all attachments in the given list
            Dim strIds = From i In idList Select CStr(i)
            sql = String.Format("select * from l0_doc_attachment where id in ({0})", Join(strIds.ToArray, ","))
        End If
        Dim rawData = DBAccess.dbDoc.SQL2RD(sql)

        ' load matching attachments
        Dim la As New List(Of L0_Attachment)
        Dim doc_ids As New Dictionary(Of Integer, Integer)
        For Each rawrec In rawData
            Dim rec As New L0_Attachment
            rec.id = rawrec.AsInteger("id")
            doc_ids(rec.id) = rawrec.AsInteger("l0_doc_ref")
            rec.status = rawrec.AsString("status")
            rec.path = rawrec.AsString("path")
            la.Add(rec)
        Next

        ' load docs
        Dim uniqueIDs As New HashSet(Of Integer)
        For Each docid In doc_ids.Values
            uniqueIDs.Add(docid)
        Next
        Dim docsDict = LoadL0Docs(uniqueIDs.ToList).ToDictionary(Function(_l0d) _l0d.id)

        ' update attachment records
        For Each att In la
            att.l0_doc = docsDict(doc_ids(att.id))
        Next

        Return la
    End Function


    Public Shared Function LoadL0Attachment(ByVal id As Integer) As L0_Attachment
        Dim idList As New List(Of Integer) From {id}
        Dim l0atts = LoadL0Attachments(idList)
        If l0atts.Count = 0 Then
            Return Nothing
        Else
            Return l0atts(0)
        End If
    End Function



    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("doc_source_type_ref = {doc_source_type}")> _
    <DebuggerDisplay("created = {created}")> _
    <DebuggerDisplay("received = {received}")> _
    <DebuggerDisplay("source_id = {source_id}")> _
    <DebuggerDisplay("from = {from}")> _
    <DebuggerDisplay("subject = {subject}")> _
    <DebuggerDisplay("body = {body}")> _
    Public Class L0_Doc
        Public id As Integer
        Public doc_source_type As Integer
        Public created As Date
        Public received As Date
        Public source_id As String
        Public from As String
        Public subject As String
        Public body As String
    End Class


    Public Shared Function LoadL0Docs(ByVal idList As List(Of Integer)) As List(Of L0_Doc)
        Dim strIds = From i In idList Select CStr(i)
        Dim sql = String.Format("select * from l0_doc where id in ({0})", Join(strIds.ToArray, ","))
        Dim rawData = DBAccess.dbDoc.SQL2RD(sql)

        Dim recs As New List(Of L0_Doc)
        For Each rawrec In rawData
            Dim rec = New L0_Doc
            rec.id = rawrec.AsInteger("id")
            rec.doc_source_type = rawrec.AsInteger("doc_source_type_ref")
            rec.created = rawrec.AsDate("created")
            rec.received = rawrec.AsDate("received_datetime")
            rec.source_id = rawrec.AsString("source_id")
            rec.from = rawrec.AsString("from_str")
            rec.subject = rawrec.AsString("subject")
            rec.body = rawrec.AsString("body")
            recs.Add(rec)
        Next

        Return recs
    End Function


    Public Shared Function LoadL0Doc(ByVal id As Integer) As L0_Doc
        Dim idList As New List(Of Integer) From {id}
        Dim l0docs = LoadL0Docs(idList)
        If l0docs.Count = 0 Then
            Return Nothing
        Else
            Return l0docs(0)
        End If
    End Function




    '
    ' L2 tables
    '
    '

    Public Shared Function GetData_dbDoc(ByVal sql As String) As List(Of WPF_Roots.RawDataRecord)
        ' usable in Tasks...
        Return DBAccess.dbDoc.SQL2RD(sql)
    End Function



    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("created = {created}")> _
    <DebuggerDisplay("l0_att_id = {l0_att_id}")> _
    <DebuggerDisplay("current_node = {current_node}")> _
    <DebuggerDisplay("mandant = {mandant}")> _
    <DebuggerDisplay("checkout_time = {checkout_time}")> _
    <DebuggerDisplay("ext_doc_id = {ext_doc_id}")> _
    <DebuggerDisplay("path = {path}")> _
    <DebuggerDisplay("vvx = {vvx}")> _
    <DebuggerDisplay("mark = {mark}")> _
    <DebuggerDisplay("mark_text = {mark_text}")> _
    Public Class L2_Doc
        Public id As Integer
        Public created As Date
        Public current_node As Integer
        Public mandant As Integer
        Public checkout_time As Date
        Public ext_doc_id As String
        Public path As String
        Public vvx As String
        Public mark As String
        Public mark_text As String

        ' l0 items:
        Public l0doc_received As Date
        Public l0doc_from As String
        Public l0doc_subject As String

        Public data As List(Of L2_Data) = Nothing

        Public Function GetFullPathAndFileName(ByVal fileExtension As String) As String
            Dim fileName As String = id
            If fileExtension IsNot Nothing Then
                fileName &= "." & Regex.Replace(fileExtension, "^\W+", "")
            End If

            Return CombinePathsFile(AppConfig.GetItem("DocRoot"), "l2", path, fileName)
        End Function


        Public Function GetCurrentNode() As L2_Node
            Return L2Nodes.Find(Function(_n) _n.id = current_node)
        End Function


        Public Sub FillInVVItems(Optional ByVal mandantId As Integer = Integer.MinValue, Optional ByVal extDocId As String = Nothing)
            ' updates this record (memory & db) to the given values

            Dim sqlParts As New List(Of String)

            If mandantId <> Integer.MinValue Then
                Me.mandant = mandantId
                sqlParts.Add(String.Format("mandant = {0}", PSQL_Int(mandantId)))
            End If

            If extDocId IsNot Nothing Then
                Me.ext_doc_id = extDocId
                sqlParts.Add(String.Format("ext_doc_id = {0}", PSQL_String(extDocId)))
            End If

            If sqlParts.Count > 0 Then
                ' update l2_doc table
                Dim sql = "update l2_doc set " & Join(sqlParts.ToArray, ", ") & " where id = " & id
                DBAccess.dbDoc.SQLExec(sql)
            End If
        End Sub


        Public Sub SetMark(ByVal _mark As String, ByVal text As String)
            ' updates this record (memory & db) to the given values

            Dim sqlParts As New List(Of String)

            Me.mark = _mark
            Me.mark_text = text

            ' update l2_doc table
            Dim sql = String.Format("update l2_doc set mark = {0}, mark_text = {1} where id = {2}", PSQL_String(_mark, maxLength:=10), PSQL_String(text), id)
            DBAccess.dbDoc.SQLExec(sql)
        End Sub


        Public Sub LoadDataItems()
            ' load the document's data items
            data = New List(Of L2_Data)

            Dim sql As String = String.Format("select * from l2_data where doc_id = {0}", id)
            Dim rawData = DBAccess.dbDoc.SQL2RD(sql)

            For Each rawrec In rawData
                Dim rec As New L2_Data
                rec.id = rawrec.AsInteger("id")
                rec.data_type = rawrec.AsInteger("data_type")
                rec.doc_id = id
                rec.flow_id = rawrec.AsInteger("flow_id")
                rec.xml = rawrec.AsString("xml")
                data.Add(rec)
            Next
        End Sub

        Public Function CreateNewDataItem(ByVal data_type As L2DataType, ByVal xml As String,
                                          Optional ByVal flow_id As Integer = Integer.MinValue) As Boolean
            ' creates a new data entry for this document with the given parameters

            Dim sql = String.Format("insert into l2_data (doc_id, data_type, flow_id, xml) values ({0}, {1}, {2}, {3})",
                                    id, CInt(data_type), PSQL_Int(flow_id), PSQL_String(xml))

            Dim ok = DBAccess.dbDoc.SQLExec(sql)
            Return (ok = 1)
        End Function

        Public Function OverwriteDataItem(ByVal data_type As L2DataType, ByVal xml As String,
                                          Optional ByVal flow_id As Integer = Integer.MinValue) As Boolean
            ' creates a new data entry for this document with the given parameters, OVERWRITING any existing entry of the same type

            ' erase all entries of the same type
            Dim sql = String.Format("delete from l2_data where doc_id = {0} and data_type = {1}", id, CInt(data_type))
            DBAccess.dbDoc.SQLExec(sql)

            ' create new entry
            Return CreateNewDataItem(CInt(data_type), xml, flow_id)
        End Function

    End Class


    Public Enum L2DataType
        InternalNote = 1
        ExternalNote = 2
        Transaction = 3
        Barcodes = 4
        OpenIssue = 5
        ClosedIssue = 6
        DoccaData = 7
        PreProcessing = 8
    End Enum

    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("data_type = {data_type}")> _
    <DebuggerDisplay("doc_id = {doc_id}")> _
    <DebuggerDisplay("flow_id = {flow_id}")> _
    <DebuggerDisplay("xml = {xml}")> _
    Public Class L2_Data
        Public id As Integer
        Public data_type As Integer
        Public doc_id As Integer
        Public flow_id As Integer
        Public xml As String

        Public ReadOnly Property xDoc As XDocument
            Get
                Try
                    Dim _xDoc = XDocument.Parse(xml)
                    Return _xDoc
                Catch ex As Exception
                    Return Nothing
                End Try
            End Get
        End Property
    End Class


    Public Shared Function CreateNewL2Data(ByVal data_type As L2DataType, ByVal doc_id As Integer, ByVal node_id As Integer, xml As String,
                                           Optional ByVal byContact_id As Integer = 1487, Optional flow_id As Integer = Integer.MinValue) As Integer

        Dim sql = String.Format("insert into l2_data (id, data_type, doc_id, flow_id, xml, node_id, by_contact_id) values (default, {0}, {1}, {2}, {3}, {4}, {5}) returning id",
                                PSQL_Int(data_type), PSQL_Int(doc_id), PSQL_Int(flow_id), PSQL_String(xml), PSQL_Int(node_id), PSQL_Int(byContact_id))

        Dim newId As Object = Nothing
        newId = DBAccess.dbDoc.SQL2O(sql)
        If newId Is Nothing Then
            ' insert failed
            Return Integer.MinValue
        End If

        Return CInt(newId)
    End Function




    Public Shared Function GetL2DocCount(Optional ByVal nodes As List(Of L2_Node) = Nothing,
                                         Optional ByVal nodeType As DocFlo_NodeType = Integer.MinValue,
                                         Optional ByVal mandantIds As List(Of Integer) = Nothing) As Integer
        Dim wheres As New List(Of String)
        If nodes IsNot Nothing AndAlso nodes.Count > 0 Then
            Dim nodesIds = From _n In nodes Select String.Format("{0}", _n.id)
            If nodeType <> Integer.MinValue Then
                nodesIds = From _n In nodes Where _n.node_type = nodeType Select String.Format("{0}", _n.id)
            End If
            wheres.Add(String.Format("current_node_id in ({0})", Join(nodesIds.ToArray, ",")))

        ElseIf nodeType <> Integer.MinValue Then
            Dim nodesWithRightType = From _n In L2Nodes() Where _n.node_type = nodeType Select String.Format("{0}", _n.id)
            If nodesWithRightType.Count = 0 Then
                Return 0
            End If
            wheres.Add(String.Format("current_node_id in ({0})", Join(nodesWithRightType.ToArray, ",")))
        End If
        If mandantIds IsNot Nothing AndAlso mandantIds.Count > 0 Then
            wheres.Add(String.Format("mandant in ({0})", Join((From _mid In mandantIds Select CStr(_mid)).ToArray, ",")))
        End If

        Dim sql As String = "select count(*) from l2_doc "
        If wheres.Count > 0 Then
            sql &= String.Format("where {0}", Join((From _w In wheres Select "(" & _w & ")").ToArray, " and "))
        End If

        Dim count As Integer = DBAccess.dbDoc.SQL2O(sql)
        Return count
    End Function


    Public Shared Function LoadL2Docs(Optional ByVal idList As List(Of Integer) = Nothing,
                                      Optional ByVal barcodeList As List(Of String) = Nothing,
                                      Optional ByVal nodes As List(Of L2_Node) = Nothing,
                                      Optional ByVal nodeType As DocFlo_NodeType = Integer.MinValue,
                                      Optional ByVal allowNegativeDocIds As Boolean = False) As List(Of L2_Doc)
        ' loads either L2Docs whose ids are contained in the idList,
        ' or all L2Docs which are in the given node(s) and/or node type.
        ' This function is meant as an overview - no L2_Data items are loaded. Use doc.LoadDataItems() to get them.
        '
        ' set allowNegativeDocIds to get access to debug documents (--> l2.id < 0, directly routed to trash by workflow)
        Dim recs As New List(Of L2_Doc)

        Dim wheres As New List(Of String)
        If Not allowNegativeDocIds Then
            wheres.Add("l2_doc.id > 0")
        End If
        If nodes IsNot Nothing AndAlso nodes.Count > 0 Then
            Dim nodesIds = From _n In nodes Select String.Format("{0}", _n.id)
            If nodeType <> Integer.MinValue Then
                nodesIds = From _n In nodes Where _n.node_type = nodeType Select String.Format("{0}", _n.id)
            End If
            wheres.Add(String.Format("l2_doc.current_node_id in ({0})", Join(nodesIds.ToArray, ",")))

        ElseIf nodeType <> Integer.MinValue Then
            Dim nodesWithRightType = From _n In L2Nodes() Where _n.node_type = nodeType Select String.Format("{0}", _n.id)
            If nodesWithRightType.Count = 0 Then
                Return recs
            End If
            wheres.Add(String.Format("l2_doc.current_node_id in ({0})", Join(nodesWithRightType.ToArray, ",")))

        End If
        If idList IsNot Nothing AndAlso idList.Count > 0 Then
            Dim strIds = From i In idList Select CStr(i)
            wheres.Add(String.Format("l2_doc.id in ({0})", Join(strIds.ToArray, ",")))
        End If
        If barcodeList IsNot Nothing AndAlso barcodeList.Count > 0 Then
            wheres.Add(String.Format("l2_doc.ext_doc_id in ('{0}')", Join(barcodeList.ToArray, "','")))
        End If
        If wheres.Count = 0 Then
            Return recs
        End If

        Dim sql As String =
            String.Format("select l2_doc.*, l0_doc.from_str as l0doc_from, l0_doc.received_datetime as l0doc_received, l0_doc.subject as l0doc_subject " & _
                            "from l2_doc " & _
                            "     left outer join l0_doc_attachment on (l2_doc.l0_att_id = l0_doc_attachment.id) " & _
                            "     left outer join l0_doc on (l0_doc_attachment.l0_doc_ref = l0_doc.id) " & _
                            "where {0}", Join((From _w In wheres Select "(" & _w & ")").ToArray, " and "))
        Dim rawData = DBAccess.dbDoc.SQL2RD(sql)

        For Each rawrec In rawData
            Dim rec As New L2_Doc
            rec.id = rawrec.AsInteger("id")
            rec.created = rawrec.AsDate("created")
            rec.current_node = rawrec.AsInteger("current_node_id")
            rec.mandant = rawrec.AsInteger("mandant")
            rec.checkout_time = rawrec.AsDate("checkout_time")
            rec.ext_doc_id = rawrec.AsString("ext_doc_id")
            rec.path = rawrec.AsString("path")
            rec.vvx = rawrec.AsString("vvx")
            rec.mark = rawrec.AsString("mark")
            rec.mark_text = rawrec.AsString("mark_text")

            rec.l0doc_from = rawrec.AsString("l0doc_from")
            rec.l0doc_received = rawrec.AsDate("l0doc_received")
            rec.l0doc_subject = rawrec.AsString("l0doc_subject")

            rec.data = Nothing
            recs.Add(rec)
        Next
        Return recs
    End Function


    Public Shared Function LoadL2Doc(ByVal id As Integer) As L2_Doc
        ' Loads the specified L2_Doc
        Dim l2docs = LoadL2Docs(idList:=New List(Of Integer) From {id}, allowNegativeDocIds:=True)
        If l2docs.Count = 0 Then
            Return Nothing
        End If
        Return l2docs(0)
    End Function


    Public Shared Function LoadL2Doc(ByVal ext_doc_id As String) As L2_Doc
        ' Loads the specified L2_Doc
        Dim l2docs = LoadL2Docs(barcodeList:=New List(Of String) From {ext_doc_id})
        If l2docs.Count = 0 Then
            Return Nothing
        End If
        Return l2docs(0)
    End Function


    Public Shared Function CreateNewL2Doc(ByVal mandant As Integer, Optional ByVal ext_doc_id As String = Nothing, Optional ByVal l0_att As Integer = Integer.MinValue,
                                          Optional ByVal initialNode As Integer = 101, Optional path As String = Nothing, Optional vvx As String = Nothing) As L2_Doc
        ' creates a new L2 document entry with the given parameters
        ' for an initial node other than 'Vorverarbeitung' (= 101), overwrite the initialNode parameter

        Dim sql = String.Format("insert into l2_doc (mandant, ext_doc_id, l0_att_id, current_node_id, path, vvx) values ({0}, {1}, {2}, {3}, {4}, {5}) returning id",
                                PSQL_Int(mandant), PSQL_String(ext_doc_id), PSQL_Int(l0_att), PSQL_Int(initialNode), PSQL_String(path), PSQL_String(vvx))

        Dim newId As Object = Nothing
        newId = DBAccess.dbDoc.SQL2O(sql)
        If newId Is Nothing Then
            ' insert failed
            Return Nothing
        End If

        ' return the new record
        Return LoadL2Doc(CInt(newId))
    End Function


    ' DocFlow node types & node ids
    Public Enum DocFlo_NodeType
        Vorverarbeitung = 101
        Hauptverarbeitung = 103
    End Enum

    Public Enum DocFlo_NodeID
        Trash = -1000
        FatalError = -2
        Vorverarbeitung = 101
        Hauptverarbeitung = 103
        VV_Supervisor = 108
        HVPendenz = 111
    End Enum


    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("display_name = {display_name}")> _
    <DebuggerDisplay("display_shortname = {display_shortname}")> _
    <DebuggerDisplay("node_type = {node_type}")> _
    <DebuggerDisplay("mandant_id = {mandant_id}")> _
    Public Class L2_Node
        Public id As Integer
        Public display_name As String
        Public display_shortname As String
        Public node_type As DocFlo_NodeType
        Public mandant_id As Integer
        ' Public config As String                   used by certain nodes; contains script code
    End Class


    Private Shared L2N_nodeCache As New List(Of L2_Node)
    '
    Public Shared Function L2Nodes() As List(Of L2_Node)

        SyncLock L2N_nodeCache
            If MustReload("LoadL2Node") Then
                Dim sql = String.Format("select * from l2_node")
                Dim rawData = DBAccess.dbDoc.SQL2RD(sql)

                L2N_nodeCache.Clear()
                For Each rawrec In rawData
                    Dim rec As New L2_Node
                    rec.id = rawrec.AsInteger("id")
                    rec.display_name = rawrec.AsString("display_name")
                    rec.display_shortname = rawrec.AsString("display_shortname")
                    rec.node_type = rawrec.AsInteger("node_type")
                    rec.mandant_id = rawrec.AsInteger("mandant_id")
                    L2N_nodeCache.Add(rec)
                Next
            End If
        End SyncLock

        Return L2N_nodeCache
    End Function





    Public Enum DocFlo_EdgeType
        Break = 0
        Accept = 1
        Reject = 2
        Forward = 3
        Trash = 4
    End Enum

    Public Class L2_Edge
        Public id As Integer
        Public name As String
        Public edge_type As Integer
        Public src_node As Integer
        Public target_node As Integer

        Public ReadOnly Property edge_type_name As String
            Get
                Select Case edge_type
                    Case DocFlo_EdgeType.Break
                        Return "BREAK"
                    Case DocFlo_EdgeType.Accept
                        Return "ACCEPT"
                    Case DocFlo_EdgeType.Reject
                        Return "REJECT"
                    Case DocFlo_EdgeType.Forward
                        Return "FORWARD"
                    Case DocFlo_EdgeType.Trash
                        Return "TRASH"
                End Select
                Return Nothing
            End Get
        End Property

        Public Overrides Function ToString() As String
            Return name
        End Function

        Public Shared Operator =(ByVal a As L2_Edge, ByVal b As L2_Edge) As Boolean
            If (a Is Nothing) Xor (b Is Nothing) Then
                Return False
            End If
            Return (a Is Nothing And b Is Nothing) OrElse a.id = b.id
        End Operator

        Public Shared Operator <>(ByVal a As L2_Edge, ByVal b As L2_Edge) As Boolean
            Return Not a = b
        End Operator
    End Class



    <DebuggerDisplay("id = {id}")> _
    <DebuggerDisplay("created = {created}")> _
    <DebuggerDisplay("doc_id = {doc_id}")> _
    <DebuggerDisplay("from_node = {from_node}")> _
    <DebuggerDisplay("to_node = {to_node}")> _
    <DebuggerDisplay("by_contact_id = {by_contact_id}")> _
    Public Class L2_Flow
        Public id As Integer
        Public created As Date
        Public doc_id As Integer
        Public from_node As Integer
        Public to_node As Integer
        Public by_contact_id As Integer
    End Class


    Public Shared Function CreateNewL2Flow(ByVal doc_id As Integer, ByVal from_node As Integer, ByVal to_node As Integer,
                                      Optional ByVal contact_id As Integer = Integer.MinValue) As Boolean

        Dim sql = String.Format("insert into l2_flow (id, created, doc_id, from_node, to_node, by_contact_id) values (default, default, {0}, {1}, {2}, {3})",
                                doc_id, from_node, to_node, PSQL_Int(contact_id))

        Dim affectedRows = DBAccess.dbDoc.SQL2O(sql)
        Return affectedRows = 1
    End Function


End Class
