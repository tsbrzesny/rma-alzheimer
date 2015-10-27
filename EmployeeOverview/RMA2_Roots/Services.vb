Imports System.Text.RegularExpressions
Imports System.Text
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports System.Drawing
Imports System.Net.Mail
Imports System.Windows.Forms
Imports RMA_Roots


Public Module RMA2S
    ' container for common shared services


    ' Document Classes - don't change the item names! They're used literally in XML files.
    Public Enum RMADocClass
        Kreditor
        Hauptbuchung
        DMS
        Bankbeleg
        Spesenbeleg
        Spesenformular
        Debitor
        Lohnbeleg
        Lohnunterlagen
        Kreditkartenabrechnung
        PaperBuddy
    End Enum



    ' 
    '   string services
    '
    '


    ' Ensures that a string is not Nothing
    Public Function StringNothing2Empty(ByVal s As String) As String
        If s Is Nothing Then
            Return ""
        End If
        Return s
    End Function


    ' The opposite: converts an empty string into Nothing
    Public Function StringEmpty2Nothing(ByVal s As String, Optional ByVal allWhiteIsEmpty As Boolean = False) As String
        If s Is Nothing OrElse s.Length = 0 OrElse
            (allWhiteIsEmpty AndAlso NoWhitespace(s).Length = 0) Then
            Return Nothing
        End If
        Return s
    End Function


    ' Splits a multiline string into separate lines.. whatever the line delimiter is
    Public Function SplitIntoLines(ByVal str As String) As String()
        If str Is Nothing Then
            str = ""
        End If
        str = Replace(str, vbCrLf, vbLf)
        str = Replace(str, vbCr, vbLf)
        Return Split(str, vbLf)
    End Function

    ' Splits into pieces of a certain length
    Public Function SplitIntoMaxLength(ByVal s As String, ByVal maxLen As Integer, Optional ByVal pattern As String = "(?<=[a-zA-Z,]{2})\s+") As List(Of String)
        Dim los As New List(Of String)

        If s Is Nothing Then
            los.Add(Nothing)
            Return los
        End If

        While s.Length > maxLen
            ' try to split gently
            Dim m = Regex.Matches(s, pattern)
            Dim lastHit As Match = Nothing
            For Each hit As Match In m
                If hit.Index <= maxLen Then
                    lastHit = hit
                Else
                    Exit For
                End If
            Next
            If lastHit IsNot Nothing Then
                los.Add(s.Substring(0, lastHit.Index))
                s = s.Substring(lastHit.Index + lastHit.Length)
                Continue While
            End If

            ' no success.. just cut
            los.Add(s.Substring(0, maxLen))
            s = s.Substring(maxLen)
        End While
        If s.Length > 0 Or los.Count = 0 Then
            los.Add(s)
        End If

        Return los
    End Function


    Public Function EasyJoinIf(ByVal delim As String, ByVal ParamArray items() As String) As String
        ' Trims all items and joins them if not "" or Nothing
        Dim doMultiline As Boolean = (delim = vbLf Or delim = vbCr Or delim = vbCrLf)
        Dim parts As New List(Of String)

        If doMultiline Then
            For Each item In items
                ' items are divided into single lines..
                parts.AddRange(SplitIntoLines(item))
            Next
        Else
            parts.AddRange(items)
        End If

        Dim result As String = ""
        For Each item In items
            If item Is Nothing Then
                Continue For
            End If
            item = FullTrim(item)
            If item.Length = 0 Then
                Continue For
            End If
            If result.Length = 0 Then
                result = item
            Else
                result &= delim & item
            End If
        Next
        Return result
    End Function


    Public Function EasyJoin(ByVal delim As String, ByVal ParamArray items() As String) As String
        ' Trims all items and joins them
        Dim result As String = ""
        For Each item In items
            If item Is Nothing Then
                item = ""
            Else
                item = FullTrim(item)
            End If
            If result.Length = 0 Then
                result = item
            Else
                result &= delim & item
            End If
        Next
        Return result
    End Function


    Public Function EasyJoin(ByVal delim As String, ByVal ParamArray items() As Object) As String
        ' converts items into strings and joins them
        Dim result As String = ""
        For Each item In items
            If item IsNot Nothing Then
                If result.Length = 0 Then
                    result = item.ToString
                Else
                    result &= delim & item.ToString
                End If
            End If
        Next
        Return result
    End Function


    Public Function EasyJoinA(ByVal delim As String, ByVal items As Object()) As String
        ' converts items into strings and joins them
        Dim result As String = ""
        For Each item In items
            If item IsNot Nothing Then
                If result.Length = 0 Then
                    result = item.ToString
                Else
                    result &= delim & item.ToString
                End If
            End If
        Next
        Return result
    End Function


    Public Function EasyJoinAndEscape(ByVal delim As String, ByVal escapeThis As String, ByVal escapeBy As String, ByVal ParamArray items() As Object) As String
        ' replace each occurence of escapeThis by escapeBy in all items, then call EasyJoin
        Dim itemList As New List(Of String)
        For Each item In items
            Dim itemStr = StringNothing2Empty(item)
            itemStr = itemStr.Replace(escapeThis, escapeBy)
            itemList.Add(itemStr)
        Next
        Return EasyJoin(delim, itemList.ToArray)
    End Function


    Public Function EasyLeft(ByVal s As String, ByVal n As Integer, Optional ByRef wasCut As Boolean = False) As String
        wasCut = False
        s = StringNothing2Empty(s)
        If s.Length <= n Then
            Return s
        End If
        wasCut = True
        Return Left(s, n)
    End Function


    ' FullTrim:
    ' removes all leading & trailing whitespace
    Private FullTrim_rgx As New Regex("^\s*|\s*$")
    Public Function FullTrim(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Return FullTrim_rgx.Replace(str, "")
    End Function


    ' NoWhitespace:
    ' removes all whitespace
    Private NoWhitespace_rgx As New Regex("\s+")
    Public Function NoWhitespace(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Return NoWhitespace_rgx.Replace(str, "")
    End Function


    ' FirstLine:
    ' returns the first line or Nothing
    Private FirstLine_rgx As New Regex("^(.*)$", RegexOptions.Multiline)
    Public Function FirstLine(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Dim m As Match = FirstLine_rgx.Match(str)
        If m.Success Then
            Return m.Value
        Else
            Return Nothing
        End If
    End Function


    ' CleanUpSpaces:
    ' FullTrim + Replace multiple spaces with single spaces
    Private CleanUpSpaces_rgx As New Regex(" +")
    Public Function CleanUpSpaces(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Return CleanUpSpaces_rgx.Replace(str, " ").Trim
    End Function


    ' CleanUpWhitespace:
    ' FullTrim + Replace all whitespace blocks with single spaces
    Private CleanUpWhitespace_rgx As New Regex("\s+")
    Public Function CleanUpWhitespace(ByVal str As String) As String
        If str Is Nothing Then
            Return Nothing
        End If
        Return CleanUpWhitespace_rgx.Replace(str, " ").Trim
    End Function


    ' CleanUpLines:
    ' Removes empty lines from a list
    Public Sub CleanUpLines(ByVal lines As List(Of String))
        lines.RemoveAll(Function(line) line Is Nothing OrElse line.Trim.Length = 0)
    End Sub

    ' CleanUpLines:
    ' Removes empty lines from the given string
    Public Function CleanUpLines(ByVal str As String, Optional ByVal delim As String = vbLf) As String
        If str Is Nothing Then
            Return Nothing
        End If

        Dim lines As New List(Of String)
        lines.AddRange(SplitIntoLines(str))
        CleanUpLines(lines)

        Return EasyJoin(delim, lines.ToArray)
    End Function


    Public Function ExtractStr(ByVal input As String, ByVal pattern As String)
        ' applies pattern to input & returns .Groups(1).Value, if exists
        Dim m = Regex.Match(input, pattern)
        If m.Success AndAlso m.Groups.Count > 1 Then
            Return m.Groups(1).Value
        End If
        Return Nothing
    End Function



    Public Function StrLength(ByVal str As String, ByVal width As Integer) As String
        ' cleans up spaces & cuts to given length

        If str Is Nothing Then
            Return ""
        End If

        str = CleanUpWhitespace(str)
        If str.Length <= width Then
            Return str
        End If
        Return str.Substring(0, width)
    End Function


    Public Function CombinePathFile(ByVal path As String, ByVal file As String, Optional ByVal ext As String = Nothing) As String
        ' path & [\] & file [& ext] .. inserts an optional backslash if necessary
        Dim result As String = path
        If Not result.EndsWith("\") Then
            result &= "\"
        End If
        result &= file
        If ext IsNot Nothing Then
            If Not ext.StartsWith(".") Then
                result &= "."
            End If
            result &= ext
        End If
        Return result
    End Function


    Public Function CombinePathsFile(ByVal ParamArray parts() As String) As String
        ' last item must be the filename.. \'s are inserted where needed
        Dim n = parts.Length

        Dim result As String = ""
        For i = 0 To n - 2
            result &= parts(i)
            If Not (result.EndsWith("\") Or result.EndsWith("/")) Then
                result &= "\"
            End If
        Next
        result &= parts(n - 1)
        Return result
    End Function


    Public Function RemoveInvalidCharacters(ByVal s As String) As String
        If s Is Nothing Then
            Return Nothing
        End If
        s = New String(s.Where(Function(c) Not Char.IsControl(c) OrElse c = vbCr OrElse c = vbLf).ToArray())
        Return s
    End Function


    Public Function RemoveInvalidFilenameCharacters(ByVal fileName As String) As String
        If fileName Is Nothing Then
            fileName = ""
        End If
        fileName = Regex.Replace(fileName, "[^\w_\-]", "")
        Return fileName
    End Function


    Public Function ReplaceUmlaut(ByVal s As String) As String
        ' replaces german Umlauts with their 2 character form
        ' 'ä ö ü Ä Ö Ü' -> 'ae oe ue Ae Oe Ue'
        s = s.Replace("ä", "ae")
        s = s.Replace("ö", "oe")
        s = s.Replace("ü", "ue")
        s = s.Replace("Ä", "Ae")
        s = s.Replace("Ö", "Oe")
        s = s.Replace("Ü", "Ue")
        Return s
    End Function


    Public Function RemoveDiacritics(ByVal s As String) As String
        ' original C# code from
        ' http://blog.fredrikhaglund.se/blog/2008/04/16/how-to-remove-diacritic-marks-from-strings/
        ' converts 'Abc 123 ÅÄÕ åäõ Eé.!?'
        ' into     'Abc 123 AAO aao Ee.!?'
        Dim normalizedString As String = s.Normalize(System.Text.NormalizationForm.FormD)
        Dim stringBuilder As New StringBuilder
        Dim c As Char
        For i As Integer = 0 To normalizedString.Length - 1
            c = normalizedString(i)
            If CharUnicodeInfo.GetUnicodeCategory(c) <> UnicodeCategory.NonSpacingMark Then
                stringBuilder.Append(c)
            End If
        Next
        Return stringBuilder.ToString()
    End Function


    Private HTMLxlat(,) = _
        {{"&", "&amp;"}, {"""", "&quot;"}, {"<", "&lt;"}, {">", "&gt;"}, _
         {"¡", "&iexcl;"}, {"¢", "&cent;"}, {"£", "&pound;"}, {"¤", "&curren;"}, _
         {"¥", "&yen;"}, {"¦", "&brvbar;"}, {"§", "&sect;"}, {"¨", "&uml;"}, _
         {"©", "&copy;"}, {"ª", "&ordf;"}, {"«", "&laquo;"}, {"¬", "&not;"}, _
         {"®", "&reg;"}, {"¯", "&macr;"}, {"°", "&deg;"}, {"±", "&plusmn;"}, _
         {"²", "&sup2;"}, {"³", "&sup3;"}, {"´", "&acute;"}, {"µ", "&micro;"}, _
         {"¶", "&para;"}, {"·", "&middot;"}, {"¸", "&cedil;"}, {"¹", "&sup1;"}, _
         {"º", "&ordm;"}, {"»", "&raquo;"}, {"¼", "&frac14;"}, {"½", "&frac12;"}, _
         {"¾", "&frac34;"}, {"¿", "&iquest;"}, {"À", "&Agrave;"}, {"Á", "&Aacute;"}, _
         {"Â", "&Acirc;"}, {"Ã", "&Atilde;"}, {"Ä", "&Auml;"}, {"Å", "&Aring;"}, _
         {"Æ", "&AElig;"}, {"Ç", "&Ccedil;"}, {"È", "&Egrave;"}, {"É", "&Eacute;"}, _
         {"Ê", "&Ecirc;"}, {"Ë", "&Euml;"}, {"Ì", "&Igrave;"}, {"Í", "&Iacute;"}, _
         {"Î", "&Icirc;"}, {"Ï", "&Iuml;"}, {"Ð", "&ETH;"}, {"Ñ", "&Ntilde;"}, _
         {"Ò", "&Ograve;"}, {"Ó", "&Oacute;"}, {"Ô", "&Ocirc;"}, {"Õ", "&Otilde;"}, _
         {"Ö", "&Ouml;"}, {"×", "&times;"}, {"Ø", "&Oslash;"}, {"Ù", "&Ugrave;"}, _
         {"Ú", "&Uacute;"}, {"Û", "&Ucirc;"}, {"Ü", "&Uuml;"}, {"Ý", "&Yacute;"}, _
         {"Þ", "&THORN;"}, {"ß", "&szlig;"}, {"à", "&agrave;"}, {"á", "&aacute;"}, _
         {"â", "&acirc;"}, {"ã", "&atilde;"}, {"ä", "&auml;"}, {"å", "&aring;"}, _
         {"æ", "&aelig;"}, {"ç", "&ccedil;"}, {"è", "&egrave;"}, {"é", "&eacute;"}, _
         {"ê", "&ecirc;"}, {"ë", "&euml;"}, {"ì", "&igrave;"}, {"í", "&iacute;"}, _
         {"î", "&igrave;"}, {"ï", "&iuml;"}, {"ð", "&eth;"}, {"ñ", "&ntilde;"}, _
         {"ò", "&ograve;"}, {"ó", "&oacute;"}, {"ô", "&ocirc;"}, {"õ", "&otilde;"}, _
         {"ö", "&ouml;"}, {"÷", "&divide;"}, {"ø", "&oslash;"}, {"ù", "&ugrave;"}, _
         {"ú", "&uacute;"}, {"û", "&ucirc;"}, {"ü", "&uuml;"}, {"ý", "&yacute;"}, _
         {"þ", "&thorn;"}, {"ÿ", "&yuml;"} _
        }
    Public Function HTMLEscape(ByVal str As String) As String
        ' replaces all HTML special characters by their HTML equvalent

        For index0 = 0 To HTMLxlat.GetLength(0) - 1
            str = Replace(str, HTMLxlat(index0, 0), HTMLxlat(index0, 1))
        Next
        Return str
    End Function


    Public Function UrlEncode(ByVal str As String) As String
        ' replaces all URL special characters by their URL equivalent

        If str IsNot Nothing Then
            str = Regex.Replace(str, "[^a-zA-Z0-9\-_.~]", AddressOf UrlReplacement)
        End If
        Return str
    End Function

    Private Function UrlReplacement(ByVal m As Match) As String
        UrlReplacement = ""
        For Each b In System.Text.Encoding.UTF8.GetBytes(m.Groups(0).ToString)
            UrlReplacement &= "%" & Hex(b).PadLeft(2, "0")
        Next
    End Function

    Public Function ISO_8859_1_Encode(ByVal str As String) As String
        ' replaces all special characters by their iso-8859-1 equivalent
        ' the current version of SQLLedger expects stuff like that..

        If str IsNot Nothing Then
            str = Regex.Replace(str, "[^a-zA-Z0-9\-_.~]", AddressOf ISO_8859_1_Replacement)
        End If
        Return str
    End Function

    Private Function ISO_8859_1_Replacement(ByVal m As Match) As String
        ISO_8859_1_Replacement = ""
        For Each b In System.Text.Encoding.GetEncoding(28591).GetBytes(m.Groups(0).ToString)
            ISO_8859_1_Replacement &= "%" & Hex(b).PadLeft(2, "0")
        Next
    End Function


    ' 
    '   specialized (RMA) string services
    '
    '

    Public Function BooleanTo01String(ByVal value As Boolean)
        ' RMA DBs often use strings for booleans - "1" for True, "0" for False
        If value Then
            Return "1"
        Else
            Return "0"
        End If
    End Function


    Public Function IsSameCurrency(ByVal curr1 As String, ByVal curr2 As String) As Boolean
        ' forgiving currency compare
        Return curr1.Trim.ToLower = curr2.Trim.ToLower
    End Function


    Public Function Str2Double(ByVal str As String, ByRef d As Double) As Boolean
        ' converts various string representations into a Double
        ' remove this, as soon as Kofax is out!!
        ' replaced by AmountStr2Double

        str = NoWhitespace(str)
        str = Regex.Replace(str, "(?<=.)-", "0")
        str = Regex.Replace(str, "'", "")
        Dim m = Regex.Matches(str, "\.|,")
        If m.Count > 0 Then
            Dim lastSep = m.Item(m.Count - 1).Value
            If lastSep = "," Then
                str = Regex.Replace(str, "\.", "")
                str = Regex.Replace(str, ",", ".")
            Else
                str = Regex.Replace(str, ",", "")
            End If
        End If
        Return Double.TryParse(str, d)
    End Function


    Public Function AmountStr2Double(ByVal str As String, ByRef d As Double) As Boolean
        ' converts various string representations of amounts into a Double

        str = NoWhitespace(str)
        str = Regex.Replace(str, "(?<=.)--?", "0")
        str = Regex.Replace(str, "['‘´]", "")
        Dim m = Regex.Matches(str, "\.|,")
        If m.Count > 0 Then
            Dim lastSep = m.Item(m.Count - 1).Value
            Dim pos = m.Item(m.Count - 1).Index
            str = str.Remove(pos, 1)
            str = str.Insert(pos, "@")
            If lastSep = "," Then
                str = Regex.Replace(str, "\.", "")
            Else
                str = Regex.Replace(str, ",", "")
            End If
        End If
        str = str.Replace("@", ".")
        Return Double.TryParse(str, d)
    End Function



    '
    '   Boolean? service
    '
    '

    Public Function TriBooleanEquals(b1 As Boolean?, b2 As Boolean?) As Boolean
        If (b1.HasValue <> b2.HasValue) Then
            Return False
        End If
        Return (Not b1.HasValue OrElse b1.Value = b2.Value)
    End Function


    '
    '   Date services
    '
    '

    Public Function Str2Date(ByVal str As String, ByVal format As String) As Date
        ' currently implemented:
        '   'yymmdd'    20yy is assumed
        '   'yyyymmdd'  

        Dim dateStr As String = Nothing
        If format Like "yymmdd" Then
            Dim m = Regex.Match(str, "^(\d{2})(\d{2})(\d{2})$")
            If m.Success Then
                dateStr = String.Format("{0}/{1}/20{2}", m.Groups(3).Value, m.Groups(2).Value, m.Groups(1).Value)
            End If

        ElseIf format Like "yyyymmdd" Then
            Dim m = Regex.Match(str, "^(\d{4})(\d{2})(\d{2})$")
            If m.Success Then
                dateStr = String.Format("{0}/{1}/{2}", m.Groups(3).Value, m.Groups(2).Value, m.Groups(1).Value)
            End If

        Else
            Throw New ApplicationException(String.Format("Str2Date: unknown format '{0}'", format))
        End If

        If dateStr Is Nothing Then
            Throw New ApplicationException(String.Format("Str2Date: String '{0}' does not fit format '{1}'", str, format))
        End If

        Try
            Return DateTime.Parse(dateStr)
        Catch
            Return Nothing
        End Try
    End Function



    Public Function AnalyzeDateStr(ByVal str As String, ByRef thisDate As Date, Optional ByRef unsure As Boolean = False) As Boolean
        str = CleanUpWhitespace(str).ToLower
        unsure = False

        ' special cases..
        If str = "h" Then
            thisDate = Now.Date
            Return True
        End If


        Dim m As Match
        Do
            ' catch 2012-04-01T00:00:00
            m = Regex.Match(str, "^(?<y>\d{4})-(?<m>\d{2})-(?<d>\d{2})t\d{2}:\d{2}:\d{2}$")
            If m.Success Then
                Exit Do
            End If

            ' catch delimiter-less format
            m = Regex.Match(str, "^(?<d>\d{2})(?<m>\d{2})(?<y>\d{2,4})$")           ' ddmmyy[yy]
            If m.Success Then
                Exit Do
            End If

            ' catch explicitly written month names
            m = Regex.Match(str, "^(?<d>\d{1,2})[-., ]*(?<m>[a-zäöü]{3,})[-., ]*(?<y>\d{2,4})$")  ' dd "month" yy[yy]
            If m.Success Then
                Exit Do
            End If

            m = Regex.Match(str, "^(?<y>\d{4})(?<del>[-./])(?<m>\d{2})\k<del>(?<d>\d{2})$")           ' yyyy-mm-dd
            If m.Success Then
                Exit Do
            End If

            ' reject characters
            If Regex.IsMatch(str, "[a-z]") Then
                Return False
            End If

            str = NoWhitespace(str)
            str = Regex.Replace(str, ",", ".")

            ' reject multiple consecutive delimiters
            If Regex.IsMatch(str, "\D{2}") Then
                Return False
            End If

            'reject on different delimiters
            Dim dDict As New Dictionary(Of Char, Integer)
            Regex.Replace(str, "\D", Function(lambdaM)
                                         dDict(lambdaM.Value) = 0
                                         Return ""
                                     End Function)
            If dDict.Keys.Count <> 1 Then
                Return False
            End If

            ' American format is mm/dd/yy, European is dd/mm/yy --> unsure if delim is "/"
            unsure = (dDict.Keys(0) = "/")

            ' delim = "-"
            str = Regex.Replace(str, "\D+", "-")

            ' catch standard delimited format
            m = Regex.Match(str, "^(?<d>\d{1,2})-(?<m>\d{1,2})-(?<y>\d{2,4})$")
            If m.Success Then
                Exit Do
            End If

            ' unknown format
            Return False
        Loop

        Dim dayStr = m.Groups("d").Value
        Dim monthStr = m.Groups("m").Value
        Dim yearStr = m.Groups("y").Value

        ' treat year
        Dim year = Val(yearStr)
        If year < 100 Then
            year += 2000
        ElseIf year < 2000 Then
            Return False
        End If

        ' treat month
        Dim month = 0
        If Regex.IsMatch(monthStr, "[a-z]") Then
            ' try to convert month name
            monthStr = RemoveDiacritics(monthStr)
            If Regex.IsMatch(monthStr, "^jan") Then
                month = 1
            ElseIf Regex.IsMatch(monthStr, "^feb|^fev") Then
                month = 2
            ElseIf Regex.IsMatch(monthStr, "^mar") Then
                month = 3
            ElseIf Regex.IsMatch(monthStr, "^apr|^avr") Then
                month = 4
            ElseIf Regex.IsMatch(monthStr, "^mai|^may") Then
                month = 5
            ElseIf Regex.IsMatch(monthStr, "^jun|^jui") Then
                month = 6
            ElseIf Regex.IsMatch(monthStr, "^jul") Then
                month = 7
            ElseIf Regex.IsMatch(monthStr, "^aug") Then
                month = 8
            ElseIf Regex.IsMatch(monthStr, "^sep") Then
                month = 9
            ElseIf Regex.IsMatch(monthStr, "^okt|^oct") Then
                month = 10
            ElseIf Regex.IsMatch(monthStr, "^nov") Then
                month = 11
            ElseIf Regex.IsMatch(monthStr, "^dez|^dec") Then
                month = 12
            Else
                Return False
            End If

        Else
            month = Val(monthStr)
        End If

        If month = 0 Then
            Return False
        End If

        ' treat day
        Dim day = Val(dayStr)

        If day = 0 Then
            Return False
        End If

        Try
            thisDate = DateSerial(year, month, day)
            If thisDate.Year = year And thisDate.Month = month And thisDate.Day = day Then
                Return True
            End If
        Catch ex As Exception
        End Try

        ' if we're unsure about day/month order, give another try
        If unsure Then
            Try
                Dim tempDay = day
                day = month
                month = tempDay

                thisDate = DateSerial(year, month, day)
                If thisDate.Year = year And thisDate.Month = month And thisDate.Day = day Then
                    Return True
                End If
            Catch ex As Exception
            End Try
        End If

        Return False
    End Function



    Public Function CheckDatePeriod(ByVal pFrom As Date, ByVal pTo As Date, Optional ByVal thisDate As Date = Nothing) As Boolean
        ' Checks if pFrom <= thisDate < pTo.
        '   - if thisDate is not specified, Now is used
        '   - pTo may be Nothing --> open period (thisDate always < pTo)
        '   - if pFrom is Nothing, False is returned

        If pFrom = Nothing Then
            Return False
        ElseIf thisDate = Nothing Then
            thisDate = Now
        End If

        If pFrom > thisDate Then
            Return False
        ElseIf pTo <> Nothing And pTo <= thisDate Then
            Return False
        End If

        Return True
    End Function



    '
    '   Financial services
    '
    '

    ' list of known currencies
    ' the following short forms are introduced for our most used currencies:
    '   c   -->     CHF
    '   e   -->     EUR
    '   g   -->     GBP
    '   u   -->     USD

    Private _KnownCurrencies As List(Of Tuple(Of String, String)) = Nothing
    Public ReadOnly Property KnownCurrencies As List(Of Tuple(Of String, String))
        Get
            If (_KnownCurrencies Is Nothing) Then
                ' add all official currencies
                _KnownCurrencies = RMA2D.estvRatesT.Select(Function(_er) _er.currency).Distinct().Select(Function(cur) New Tuple(Of String, String)(cur.ToLower(), cur)).ToList()

                ' add additional keys to common currencies
                _KnownCurrencies.Add(New Tuple(Of String, String)("chf", "CHF"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("sfr", "CHF"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("franken", "CHF"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("fr.", "CHF"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("c", "CHF"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("€", "EUR"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("euro", "EUR"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("e", "EUR"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("£", "GBP"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("g", "GBP"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("yen", "JPY"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("us$", "USD"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("$", "USD"))
                _KnownCurrencies.Add(New Tuple(Of String, String)("u", "USD"))
            End If
            Return _KnownCurrencies
        End Get
    End Property



    Public Enum ESR_Type
        ESR_Invalid = 0
        ESR = &H1            ' ESR with amount - see 'belegart' to find exact sub-type
        ESR_Plus = &H2       ' ESR without amount - see 'belegart' to find exact sub-type
        ESR_IBAN = &H4       ' ESR with encoded IBAN, without amount
        ESR_PC = &H8         ' only last group of ESR found.. 
        ESR_Incomplete_IBAN = &H1000  ' used to encode the first group of an ESR_IBAN
    End Enum


    Public Class ESR
        Public _type As ESR_Type = ESR_Type.ESR_Invalid
        Private _amount As Double? = Nothing
        Private _currency As String = Nothing
        Private _iban As String = Nothing
        Private _pc As String = Nothing         ' formatted Teilnehmernummer, if available (XX-XXXXXX-X, PC account)
        Private _stn As String = Nothing        ' sub-TN (left 6 digits of p2)
        Private _ref As String = Nothing        ' reference number (=p2)
        Private _belegart As String = ""
        '   01 = ESR in CHF
        '   03 = N-ESR in CHF
        '   04 = ESR+ in CHF
        '   11 = ESR in CHF zur Gutschrift auf das eigene Konto (Ziffer 3.3.4)
        '   14 = ESR+ in CHF zur Gutschrift auf das eigene Konto (Ziffer 3.3.4)
        '   21 = ESR in EUR
        '   23 = ESR in EUR zur Gutschrift auf das eigene Konto (Ziffer 3.3.4)
        '   31 = ESR+ in EUR
        '   33 = ESR+ in EUR zur Gutschrift auf das eigene Konto (Ziffer 3.3.4)
        '   OLD = old format {15}{15}{9} NO PARITY!

        Private inhibit_p3 As New List(Of String)

        Public Shared Function CheckESR(ByRef esr As String) As Boolean
            Dim p1 As String = ""
            Dim p2 As String = ""
            Dim p3 As String = ""
            Dim ba As String = ""
            Dim et As ESR_Type
            Return CheckESRToParts(esr, et, ba, p1, p2, p3)
        End Function


        Private Shared Function CheckESRToParts(ByRef esr As String, ByRef type As ESR_Type, ByRef belegart As String, _
                                                ByRef p1 As String, ByRef p2 As String, ByRef p3 As String) As Boolean
            esr = RMA2S.NoWhitespace(esr)
            p1 = ""
            p2 = ""
            p3 = ""
            belegart = ""

            Dim m_esr = Regex.Match(esr, "^(\d{13})>(\d{5,27})\+(\d{9})>$")
            Dim m_esrplus = Regex.Match(esr, "^(\d{3})>(\d{5,27})\+(\d{9})>$")
            Dim m_iban = Regex.Match(esr, "^(\d{5,27})\+(\d{9})>$")
            Dim m_pc = Regex.Match(esr, "^(\d{9})>$")
            Dim m_old = Regex.Match(esr, "^<(\d{15})>(\d{15})\+(\d{5,9})>$")

            Dim isOldFormat As Boolean = False
            If m_old.Success Then
                p1 = m_old.Groups(1).Value
                p2 = m_old.Groups(2).Value
                p3 = m_old.Groups(3).Value
                type = ESR_Type.ESR_Plus
                belegart = "OLD"
                isOldFormat = True

            ElseIf m_esr.Success Then
                p1 = m_esr.Groups(1).Value
                p2 = m_esr.Groups(2).Value.PadLeft(27, "0")
                p3 = m_esr.Groups(3).Value
                belegart = p1.Substring(0, 2)
                type = ESR_Type.ESR

            ElseIf m_esrplus.Success Then
                p1 = m_esrplus.Groups(1).Value
                p2 = m_esrplus.Groups(2).Value.PadLeft(27, "0")
                p3 = m_esrplus.Groups(3).Value
                belegart = p1.Substring(0, 2)
                type = ESR_Type.ESR_Plus

            ElseIf m_iban.Success Then
                p2 = m_iban.Groups(1).Value.PadLeft(27, "0")
                p3 = m_iban.Groups(2).Value
                type = ESR_Type.ESR_Incomplete_IBAN

            ElseIf m_pc.Success Then
                p3 = m_pc.Groups(1).Value
                type = ESR_Type.ESR_PC

            Else
                ' no ESR pattern matches...
                Return False
            End If

            ' check parity info
            If Not isOldFormat Then
                If Not RMA2S.CheckStringParityModulo10(p3) OrElse _
                    ((type <> ESR_Type.ESR_PC) AndAlso Not RMA2S.CheckStringParityModulo10(p2)) OrElse _
                    ((type = ESR_Type.ESR Or type = ESR_Type.ESR_Plus) AndAlso Not RMA2S.CheckStringParityModulo10(p1)) Then
                    ' parity error
                    Return False
                End If
            End If

            ' create a beautified version
            If isOldFormat Then
                esr = String.Format("<{2}>{1}+{0}>", p3, p2, p1)
            Else
                Select Case type
                    Case ESR_Type.ESR_PC
                        esr = String.Format("{0}>", p3)
                    Case ESR_Type.ESR_Incomplete_IBAN
                        esr = String.Format("{1}+{0}>", p3, p2)
                    Case ESR_Type.ESR, ESR_Type.ESR_Plus
                        esr = String.Format("{2}>{1}+{0}>", p3, p2, p1)
                    Case Else
                        Throw New Exception("RMA2S.ESR.CheckESR: Unknown ESR-Type!")
                End Select
            End If
            Return True
        End Function


        Public Function GetESRKey() As String
            If Not IsValid() Then
                Return Nothing
            End If

            If RMA2D.GetBCNRfromPCESR(_pc) IsNot Nothing Then
                Return String.Format("{0}/{1}¬esr", _pc, _stn)

            Else
                Return String.Format("{0}¬pc", _pc)
            End If
        End Function


        Public ReadOnly Property GetPCAccount As String
            Get
                If Not IsValid() Then
                    Return Nothing
                End If
                Return _pc
            End Get
        End Property


        Private _esrDisplay As String = ""
        Public Function GetDisplayVersion(Optional ByVal esrCodeOnly As Boolean = True) As String
            If Not IsValid() Then
                Return "(ungültiger ESR)"

            ElseIf esrCodeOnly Then
                Return _esrDisplay

            End If

            ' produce a real display version
            Dim dStr As String = ""
            Select Case _type
                Case ESR_Type.ESR
                    If _amount.HasValue Then
                        dStr = String.Format("ESR über {0:F2} {1} auf PC {2}, Ref = {3}", _amount.Value, _currency, _pc, RefNr())
                    Else
                        dStr = String.Format("ESR in {0} auf PC {1}, Ref = {2}", _currency, _pc, RefNr())
                    End If

                Case ESR_Type.ESR_Plus
                    dStr = String.Format("ESR+ in {0} auf PC {1}, Ref = {2}", _currency, _pc, RefNr())

                Case ESR_Type.ESR_PC
                    dStr = String.Format("ESR auf PC {0}", _pc)

                Case Else
                    dStr = "(unbekannter ESR-Typ)"

            End Select

            Return dStr
        End Function


        Public Sub AddESRNumber(ByVal esrStr As String)
            ' split & check parity
            Dim p1 As String = ""
            Dim p2 As String = ""
            Dim p3 As String = ""
            Dim inputBelegart As String = Nothing
            Dim inputET As ESR_Type
            If Not CheckESRToParts(esrStr, inputET, inputBelegart, p1, p2, p3) Then
                Return
            End If

            ' check if we've seen that already (as part of a longer esr fragment, for example)
            ' this takes care of a malicious problem with double-part IBAN esrs
            If inhibit_p3.Contains(p3) Then
                Return
            End If
            inhibit_p3.Add(p3)

            ' check if we can use this additional info
            If _type = 0 OrElse
                inputET = ESR_Type.ESR OrElse
                (_type <> ESR_Type.ESR AndAlso inputET = ESR_Type.ESR) Then
                ' first input or higher rank
                _type = inputET
                _belegart = inputBelegart

            ElseIf (inputET = ESR_Type.ESR_PC And _type = ESR_Type.ESR_Incomplete_IBAN) Or _
                   (inputET = ESR_Type.ESR_Incomplete_IBAN And _type = ESR_Type.ESR_PC) Then
                ' ESR_Incomplete_IBAN needs a ESR_PC..
                _type = ESR_Type.ESR_IBAN

            Else
                ' can't use this..
                Return
            End If

            ' store a display version of the longest esr fragment
            If esrStr.Length > _esrDisplay.Length Then
                _esrDisplay = esrStr
            End If


            ' extract betrag
            If _belegart = "OLD" Then
                ' this is an old 15-15-5 format
                _amount = Val(p1.Substring(6, 9)) / 100

            ElseIf inputET = ESR_Type.ESR AndAlso _
                  (_belegart = "01" Or _belegart = "21") Then
                _amount = Val(p1.Substring(2, 10)) / 100
            End If

            ' currency
            _currency = "CHF"
            If (_belegart = "21" Or _belegart = "23" Or _belegart = "31" Or _belegart = "33") Then
                _currency = "EUR"
            End If

            ' account info
            If _belegart = "OLD" Then
                ' this is an old 15-15-5 format
                _pc = p3           ' take as-is

            ElseIf inputET = ESR_Type.ESR_Incomplete_IBAN Then
                Dim iban = p3.Substring(2, 5) & p2.Substring(14, 12)
                Dim ibanParity = RMA2S.ModuloBig(iban & "121700", 97)
                _iban = "CH" & (97 + 1 - ibanParity).ToString.PadLeft(2, "0") & iban
                If Not CheckIBANNumber(_iban, True) Then
                    Throw New Exception("CheckESR: IBAN creation error!")
                End If

            Else
                _pc = p3
                RMA2S.CheckPCNumber(_pc)  ' format XX-XXXXXX-X
            End If

            ' vendor info & reference number
            If inputET = ESR_Type.ESR Or inputET = ESR_Type.ESR_Plus Then
                _stn = p2.Substring(0, 6)
                _ref = p2
            End If
        End Sub

        Public Function Type(ByRef et As ESR_Type, ByRef belegArt As String) As Boolean
            et = _type
            belegArt = _belegart
            Return _type <> ESR_Type.ESR_Incomplete_IBAN
        End Function

        Public Function IsValid() As Boolean
            Return _type <> ESR_Type.ESR_Invalid And _type <> ESR_Type.ESR_Incomplete_IBAN
        End Function

        Public Sub IDs(ByRef account As String, ByRef subTN As String, ByRef refNr As String)
            If _type = ESR_Type.ESR_IBAN Then
                account = _iban
            Else
                account = _pc
            End If
            subTN = _stn
            refNr = _ref
        End Sub

        Public Function RefNr() As String
            If _ref Is Nothing Then
                Return ""
            Else
                ' return _ref, grouped into 5 chars from the right
                Dim refStr = _ref
                For pos = refStr.Length - 5 To 1 Step -5
                    refStr = refStr.Insert(pos, " ")
                Next
                Return refStr
            End If
        End Function

        Public Sub AmountAndCurrency(ByRef esrAmount As Double, ByRef currency As String)
            ' call this if type is 'ESR'
            currency = _currency
            If _amount.HasValue Then
                esrAmount = _amount.Value
            Else
                esrAmount = 0.0
            End If
        End Sub

        Public Function Currency() As String
            ' call this if type is <> 'ESR'
            Return _currency
        End Function

        Public Overrides Function ToString() As String
            Return GetDisplayVersion()
        End Function

        ' direct object identity
        Public Shared Operator =(ByVal a As ESR, ByVal b As ESR) As Boolean
            Return a.IsValid AndAlso b.IsValid AndAlso a._esrDisplay = b._esrDisplay
        End Operator
        Public Shared Operator <>(ByVal a As ESR, ByVal b As ESR) As Boolean
            Return Not a = b
        End Operator

    End Class


    Public Function CheckIBANNumber(ByRef iban As String, Optional ByVal formatIBAN As Boolean = False) As Boolean
        ' checks the parity of an IBAN number
        ' returns a formatted version if desired
        If iban Is Nothing Then
            Return False
        End If

        Dim localIBAN = Regex.Replace(iban.ToUpper, "[^A-Z0-9]", "")

        ' erste 4 Stellen Ländercode & Prüfziffer und dann 
        ' KontoIdent ... niemals kürzer als 15 (Norwegen)!
        If localIBAN.Length < 15 Then
            Return False
        End If

        ' Modulo check..
        Dim checkIBAN = localIBAN.Substring(4) & localIBAN.Substring(0, 4)
        checkIBAN = Regex.Replace(checkIBAN, "[A-Z]", Function(m) (Asc(m.Groups(0).Value) - 64 + 9).ToString)
        If ModuloBig(checkIBAN, 97) <> 1 Then
            Return False
        End If

        iban = localIBAN
        If Not formatIBAN Then
            Return True
        End If

        ' give back a formatted IBAN
        For i = ((iban.Length - 1) \ 4) To 1 Step -1
            iban = iban.Insert(4 * i, " ")
        Next
        Return True

        ' Checking algorithm:
        ' http://de.wikipedia.org/wiki/International_Bank_Account_Number#Validierung_der_Pr.C3.BCfsumme
        ' IBAN:       DE68 2105 0170 0012 3456 78
        ' Umstellung: 2105 0170 0012 3456 78DE 68
        ' Modulus:    210501700012345678131468 mod 97 = 1
    End Function


    Public Function GetSIXBankStamm4CHIBAN(ByVal iban As String) As RMA2D.SIXBankStamm
        ' Gets the associated Bank-Stamm record for the given Swiss/Liechtensteiner IBAN
        If iban Is Nothing OrElse Not CheckIBANNumber(iban) Then
            Return Nothing
        End If

        iban = NoWhitespace(iban).ToUpper
        Dim m = Regex.Match(iban, "^(?:CH|LI)\d{2}(\d{5})")
        If Not m.Success Then
            Return Nothing
        End If

        Return RMA2D.GetSIXBankStammByBCNr(m.Groups(1).Value)
    End Function


    Public Function CheckPCNumber(ByRef pc As String) As Boolean
        ' checks the parity of a given PC code & returns a fomatted version on success
        If pc Is Nothing Then
            Return False
        End If

        Dim localPC = NoWhitespace(pc).ToUpper

        Dim m = Regex.Match(localPC, "^(\d{2})-?(\d{1,6})-?(\d)$")
        If m.Success Then
            ' format is basically ok.. check parity
            Dim pc1 = m.Groups(1).Value
            Dim pc2 = m.Groups(2).Value.PadLeft(6, "0")
            Dim pcPar = m.Groups(3).Value

            If CheckStringParityModulo10(pc1 & pc2 & pcPar) Then
                pc = String.Format("{0}-{1}-{2}", pc1, pc2.PadLeft(6, "0"), pcPar)
                Return True
            End If
        End If

        Return False
    End Function


    Public Function CheckStringParityModulo10(ByVal str As String) As Boolean
        ' all chars must be digits - last char is considered to be the Mod10-parity

        str = NoWhitespace(str)
        If Not Regex.IsMatch(str, "^\d*$") Then
            Return False
        End If

        Dim parity = 0
        Dim parityTable() = {0, 9, 4, 6, 8, 2, 7, 1, 3, 5}
        For Each c In str
            parity = parityTable((parity + Val(c)) Mod 10)
        Next
        Return (parity = 0)

    End Function


    Public Function CheckSwissUID(ByRef uid As String) As Boolean
        If uid Is Nothing Then
            Return False
        End If

        Dim uidCheck = Regex.Replace(uid, "\D", "")
        If uidCheck.Length <> 9 Then ' OrElse Not CheckStringParityModulo11(uidCheck) Then
            ' swiss uid must have 9 numeric digits, Mod11 = 0
            Return False
        End If

        ' uid digits are ok
        uidCheck = uidCheck.Insert(6, ".")
        uidCheck = uidCheck.Insert(3, ".")
        uid = "CHE-" & uidCheck
        Return True
    End Function


    Public Function CheckStringParityModulo11(ByVal str As String) As Boolean
        ' all chars must be digits - last char is considered to be the Mod11-parity

        str = NoWhitespace(str)
        If Not Regex.IsMatch(str, "^\d*$") Then
            Return False
        End If

        Dim paritySum = 0
        Dim multiplier = 2
        For Each c In str
            paritySum += Val(c) * multiplier
            multiplier += 1
        Next
        Return (paritySum Mod 11 = 0)

    End Function



    Public Function CheckSWIFTBICFormat(ByRef bic As String) As Boolean
        ' accepts 8 digit and 11 digit BICs
        '  - digits 0-5 must be alphas
        '  - digits 6-10 must be alphanumeric

        Dim bicCopy = NoWhitespace(StringNothing2Empty(bic)).ToUpper
        If Regex.IsMatch(bicCopy, "^[A-Z]{6}[A-Z0-9]{2}(?:[A-Z0-9]{3})?$") Then
            bic = bicCopy
            Return True
        End If
        Return False
    End Function


    '
    '   Math
    '
    '


    Public Function ModuloBig(ByVal inputInteger As String, ByVal divisor As Long) As Long
        ' Modulo for big Integers
        Dim pos = 0
        Dim resultInteger = ""
        Dim reminder As Long = 0
        Dim part As Long = 0
        Dim notSeen As String = ""

        For pos = 0 To inputInteger.Length - 1
            part = Val(part.ToString & inputInteger(pos))

            If part >= divisor Then
                resultInteger &= (part \ divisor).ToString
                reminder = part Mod divisor
                part = reminder
                notSeen = ""
            Else
                resultInteger &= "0"
                notSeen &= inputInteger(pos)
            End If
        Next
        ' beautify result - erase unneeded leading zeroes
        resultInteger = Regex.Replace(resultInteger, "^0+(?=\d)", "")

        reminder = Val(reminder & notSeen)
        Return reminder
    End Function



    '
    '   Color
    '
    '


    Public Function MidColor(ByVal c1 As Color, ByVal c2 As Color, ByVal percent As Integer) As Color
        ' calculates a color between c1 and c2, percent being the relative position (should be 0..100)
        ' percent = 0:   returns pure c1
        ' percent = 100: returns pure c2
        If percent < 0 Then
            percent = 0
        ElseIf percent > 100 Then
            percent = 100
        End If
        Dim mid As Color = Color.FromArgb((c1.A * (100 - percent) + c2.A * percent) \ 100,
                                          (c1.R * (100 - percent) + c2.R * percent) \ 100,
                                          (c1.G * (100 - percent) + c2.G * percent) \ 100,
                                          (c1.B * (100 - percent) + c2.B * percent) \ 100)
        Return mid
    End Function



    ' 
    '   low level hacks
    '
    '


    Public Function CallMethod(ByVal name As String, ByVal o As Object, ByVal ParamArray params() As Object) As Object
        ' use reflection to call an object's protected methods
        Dim objectType As Type = o.GetType
        Dim flags As BindingFlags = BindingFlags.Instance Or _
                                    BindingFlags.Static Or _
                                    BindingFlags.NonPublic Or _
                                    BindingFlags.Public
        Dim methodInfo As MethodInfo = Nothing
        Dim types(params.Length - 1) As Type
        For index = 0 To params.Length - 1
            types(index) = params(index).GetType
        Next
        While Not objectType Is Nothing
            methodInfo = objectType.GetMethod(name, flags, Nothing, types, Nothing)
            If methodInfo Is Nothing Then
                objectType = objectType.BaseType
            Else
                Exit While
            End If
        End While
        If methodInfo Is Nothing Then
            Throw New MissingMemberException(o.GetType.FullName, name)
        Else
            Return methodInfo.Invoke(o, params)
        End If
    End Function


    Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Integer, _
                                                                    ByVal wMsg As Integer, _
                                                                    ByVal wParam As Integer,
                                                                    ByVal lParam As Integer) As Integer

    Private Const WM_SETREDRAW As Integer = 11

    ' Extension methods for Control
    Public Sub ResumeDrawing(ByVal Target As Control, ByVal Redraw As Boolean)
        SendMessage(Target.Handle, WM_SETREDRAW, 1, 0)
        If Redraw Then
            Target.Refresh()
        End If
    End Sub

    Public Sub SuspendDrawing(ByVal Target As Control)
        SendMessage(Target.Handle, WM_SETREDRAW, 0, 0)
    End Sub



    ' 
    '   HTML document creation
    '
    '

    Public Class HTMLDocCreator
        Private doc As New List(Of String)
        Private opentags As New Stack(Of String)

        Public Sub StartTag(ByVal tag As String)
            tag = Trim(tag.ToLower())
            doc.Add("<" & tag & ">")
            opentags.Push(tag)
        End Sub

        Public Sub StartTag(ByVal tag As String, ByVal ParamArray elements() As String)
            tag = Trim(tag.ToLower())
            Dim eList As String = ""
            For Each e In elements
                Dim eParts = Split(e, "=", 2)
                If eParts.Length > 1 Then
                    e = Trim(eParts(0)) & "=" & Chr(34) & Trim(eParts(1)) & Chr(34)
                End If
                eList &= " " & e & ";"
            Next
            doc.Add("<" & tag & eList & ">")
            opentags.Push(tag)
        End Sub

        Public Sub StartTags(ByVal ParamArray tags() As String)
            For Each tag In tags
                StartTag(tag)
            Next
        End Sub

        Public Sub CloseTag(Optional ByVal tag As String = Nothing)
            ' check if any tags open
            If opentags.Count = 0 Then
                Throw New ApplicationException("HTMLDocCreator.CloseTag: No more open tags!")
            End If

            ' if no tag is specified, close next one
            If tag Is Nothing Then
                tag = opentags.Peek
            End If

            tag = Trim(tag.ToLower())
            ' pop next open tag off the stack & compare to 
            Dim nextOpen = opentags.Pop()
            If tag <> nextOpen Then
                Throw New ApplicationException(String.Format("HTMLDocCreator.CloseTag: Can't close '{0}', because next open is '{1}'!", tag, nextOpen))
            End If
            doc.Add("</" & tag & ">")
        End Sub

        Public Sub CloseTags(ByVal ParamArray tags() As String)
            ' closes next specified tags
            For Each tag In tags
                CloseTag(tag)
            Next
        End Sub

        Public Sub CloseTags(ByVal n As Integer)
            ' closes n next tag(s)
            For i = 1 To n
                CloseTag()
            Next
        End Sub

        Public Sub Element(ByVal e As String)
            doc.Add(e)
        End Sub

        Public Sub CompleteTag(ByVal tag As String, ByVal content As String, ByVal ParamArray elements() As String)
            tag = Trim(tag.ToLower())
            Dim eList As String = ""
            For Each e In elements
                Dim eParts = Split(e, "=", 2)
                If eParts.Length > 1 Then
                    e = Trim(eParts(0)) & "=" & Chr(34) & Trim(eParts(1)) & Chr(34)
                End If
                eList &= " " & e & ";"
            Next

            If content Is Nothing Then
                doc.Add("<" & tag & eList & "/>")
            Else
                doc.Add("<" & tag & eList & ">" & RMA2S.HTMLEscape(content) & "</" & tag & eList & ">")
            End If
        End Sub

        Public Sub Ln(Optional ByVal n As Integer = 1)
            For i = 1 To n
                doc.Add("<br>")
            Next
        End Sub

        Public Sub TextLn(ByVal str As String, Optional ByVal n As Integer = 1)
            Dim lines = RMA2S.SplitIntoLines(str)
            For Each line In lines
                line = RMA2S.HTMLEscape(line) & "<br>"
                doc.Add(line)
            Next
            For i = 2 To n
                Ln()
            Next
        End Sub

        Public Sub Text(ByVal str As String, Optional ByVal raw As Boolean = False)
            Dim lines = RMA2S.SplitIntoLines(str)
            Dim addBR = lines.Length()
            For Each line In lines
                If Not raw Then
                    line = RMA2S.HTMLEscape(line)
                End If
                addBR -= 1
                If addBR > 0 Then
                    line &= "<br>"
                End If
                doc.Add(line)
            Next
        End Sub

        Public Sub LnText(ByVal str As String, Optional ByVal n As Integer = 1)
            For i = 1 To n
                Ln()
            Next
            Text(str)
        End Sub

        Public Sub TableRow(ByVal ParamArray columns As String())
            StartTag("TR")
            For Each fieldTxt In columns
                StartTag("TD")
                Text(fieldTxt)
                CloseTag("TD")
            Next
            CloseTag("TR")
        End Sub

        Public Sub Link(ByVal href As String, ByVal linkText As String)
            StartTag("a", String.Format("href={0}", href))
            Text(linkText)
            CloseTag()
        End Sub

        Public Function GetDoc() As String
            ' check for open tags
            If opentags.Count <> 0 Then
                Dim tagList = Join(opentags.ToArray, ", ")
                Dim msg = "HTMLDocCreator.GetDoc: not all tags are closed: " & tagList
                Throw New ApplicationException(msg)
            End If

            Return Join(doc.ToArray, vbLf)
        End Function
    End Class


    '
    '   Logging
    '
    '

    Private logFullLogFile As String = Nothing
    Private logStream As FileStream
    Private logStreamWriter As StreamWriter
    Public Function StartLogFile(ByVal path As String, ByVal appName As String) As Boolean
        Dim dateStr As String = Now.ToString("yyyyMMdd")
        appName = Regex.Replace(appName, "\s*", "")
        logFullLogFile = CombinePathFile(path, appName & "_" & dateStr, "log")
        logStream = New FileStream(logFullLogFile, FileMode.Append, FileAccess.Write)
        logStreamWriter = New StreamWriter(logStream)
        logStreamWriter.AutoFlush = True
    End Function

    Public Sub EndLogFile()
        logStreamWriter.Close()
        logStream.Close()
        logStream = Nothing
        logFullLogFile = Nothing
    End Sub

    Public Sub WriteLog(ByVal ParamArray lines() As String)
        If LogIsReady() Then
            For Each line In lines
                logStreamWriter.Write(line & vbCrLf)
            Next
        Else
            ' Throw New ApplicationException("Call StartLogFile before logging stuff.")
        End If
    End Sub

    Public Function LogIsReady() As Boolean
        Return (logStreamWriter IsNot Nothing)
    End Function



    ' 
    '   Debug management
    '
    ' Uses a configuration parameter "DEBUG":
    '   not defined or empty:   disable debug mode
    '   = "MsgBox" :            reroute all output to msg boxes
    '   = path :                reroute all output to the given file
    '


    Public Sub DoDEBUGOutput(ByVal caption As String, ByVal ParamArray lines() As String)
        ' reroutes the given output to either a MsgBox or a text file, depending on the configured DEBUG mode

        ' build the output text block
        Dim output As String = vbLf & caption
        For Each line In lines
            output &= vbLf & "  " & line
        Next

        ' get the debug configuration and send the output to the desired location
        Dim debug = AppConfig.GetItem("DEBUG")
        If debug Is Nothing Then
            Throw New ApplicationException("DoDEBUGOutput() called, but no Debug mode is configured.")

        ElseIf debug.ToLower = "msgbox" Then
            ' show a message box..
            My.Computer.Clipboard.SetText(output)
            MsgBox(output, MsgBoxStyle.OkOnly, "Debugging Output (copied to Clipboard)")

        Else
            ' send output to a (text) file.. 
            My.Computer.FileSystem.WriteAllText(debug, output, True, System.Text.Encoding.GetEncoding(28591))

        End If

    End Sub


    Public Property EnableManualDEBUGMode As Boolean
    '
    Public Function CheckDEBUGState() As Boolean
        ' returns True if a DEBUG mode is on
        Dim debug = AppConfig.GetItem("DEBUG")
        Return EnableManualDEBUGMode OrElse (debug IsNot Nothing)
    End Function


    Public Sub WriteFileAndDebug(ByVal fileType As String, ByVal file As String, ByVal text As String, ByVal append As Boolean, _
                                        Optional ByVal encoding As System.Text.Encoding = Nothing)
        ' calls FileSystem.WriteAllText unless the DEBUG flag is set --> reroute output to a MsgBox/debug file

        If encoding Is Nothing Then
            encoding = System.Text.Encoding.GetEncoding(28591)
        End If

        If Not CheckDEBUGState() Then
            ' no debugging active.. write the file as requested

            ' check if path exists - create otherwise
            Dim m = Regex.Match(file, "(.*)\\[^\\]+?$")
            If m.Success Then
                Dim path = m.Groups(1).Value
                If Not My.Computer.FileSystem.DirectoryExists(path) Then
                    My.Computer.FileSystem.CreateDirectory(path)
                End If
            End If

            ' write file
            My.Computer.FileSystem.WriteAllText(file, text & vbLf, append, encoding)

        Else
            ' format as text and send to Debug output
            If fileType Is Nothing Then
                fileType = "(unknown type)"
            End If
            DoDEBUGOutput(String.Format("Writing {0} file..", fileType), String.Format("File: {0}", file), _
                          String.Format("Append: {0}", append), String.Format("Encoding: {0}", encoding.EncodingName), _
                          String.Format("Content: {0}", text))
        End If
    End Sub


    '
    '   Error management
    '
    '

    Private ErrorManagement_caption As String
    Private ErrorManagement_ex As Exception
    Private ErrorManagement_str As String
    Public ErrorManagement_retry As Boolean

    Public Sub ErrorManagement_Prepare(ByVal caption As String)
        ErrorManagement_caption = caption
        ErrorManagement_ex = Nothing
        ErrorManagement_str = Nothing
        ErrorManagement_retry = True
    End Sub

    Public Sub ErrorManagement_SetEx(ByVal ex As Exception)
        ErrorManagement_ex = ex
    End Sub

    Public Sub ErrorManagement_SetStr(ByVal str As String)
        ErrorManagement_str = str
    End Sub

    Public Function ErrorManagement_GetText() As String
        Dim result As String = Nothing
        If ErrorManagement_str IsNot Nothing Then
            result = ErrorManagement_str
        End If
        If ErrorManagement_ex IsNot Nothing Then
            result = EasyJoinIf(vbCrLf, result, ErrorManagement_ex.Message)
        End If
        Return result
    End Function


    Public Sub EM_Check(ByVal ok As Boolean)
        If ok Then
            ' the checked operation succeeded.. 
            ErrorManagement_retry = False
            Return
        End If

        ' no success. Show MsgBox with Retry/Cancel
        Dim result As MsgBoxResult = MsgBox(ErrorManagement_GetText(), MsgBoxStyle.RetryCancel, ErrorManagement_caption)
        If result = MsgBoxResult.Retry Then
            ' User chose to retry
            Return
        End If

        ' User pressed Cancel.
        Throw New ApplicationException("Aktion abgebrochen")
    End Sub

End Module


