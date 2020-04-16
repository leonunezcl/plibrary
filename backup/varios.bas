Attribute VB_Name = "mVarios"
Option Explicit

Public gbMatchCase As Integer
Public gbWholeWord As Integer
Public gsFindText As String
Public gbLastPos As Integer
Public gsBuffer As String
Public glbFindSql As String
Public gsHtml As String
Public gsLastPath As String
Public glbColorizarCodigo As Boolean
Public glbRespaldarLibreria As Boolean

'carga las opciones miscelaneas
Public Sub CargaOpciones()

    Dim Valor As Variant
    
    Valor = LeeIni("opciones", "colorizar", C_INI)
    If Valor <> "" Then
        If Valor = 1 Then
            glbColorizarCodigo = True
        Else
            glbColorizarCodigo = False
        End If
    Else
        glbColorizarCodigo = True
    End If
        
    Valor = LeeIni("opciones", "respaldar", C_INI)
    If Valor <> "" Then
        If Valor = 1 Then
            glbRespaldarLibreria = True
        Else
            glbRespaldarLibreria = False
        End If
    Else
        glbRespaldarLibreria = True
    End If
    
End Sub


'genera un archivo .html
Public Function GuardarArchivoHtml(ByVal Archivo As String, ByVal Titulo As String) As Boolean

    On Local Error GoTo ErrorGuardarArchivoHtml
    
    Dim ret As Boolean
    Dim nFreeFile As Long
    
    ret = True
    
    nFreeFile = FreeFile
    
    Open Archivo For Output As #nFreeFile
        Print #nFreeFile, "<html>"
        Print #nFreeFile, "<head><title>" & Titulo & "</title></head>"
        Print #nFreeFile, "<body>"
        Print #nFreeFile, gsHtml
        Print #nFreeFile, "</body>"
        Print #nFreeFile, "</html>"
    Close #nFreeFile
    
    GoTo SalirGuardarArchivoHtml
    
ErrorGuardarArchivoHtml:
    ret = False
    MsgBox "GuardarArchivoHtml : " & Err & " " & Error$, vbCritical
    Resume SalirGuardarArchivoHtml
    
SalirGuardarArchivoHtml:
    GuardarArchivoHtml = ret
    Err = 0
    
End Function
Public Function RTF2HTML(strRTF As String, Optional strOptions As String, Optional strHeader As String, Optional strFooter As String) As String
    'Version 2.9

    'The current version of this function is available at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2html.zip

    'More information can be found at
    'http://www2.bitstream.net/~bradyh/downloads/rtf2htmlrm.html

    'Converts Rich Text encoded text to HTML format
    'if you find some text that this function doesn't
    'convert properly please email the text to
    'bradyh@bitstream.net

    'Options:
    '+H              add an HTML header and footer
    '+G              add a generator Metatag
    '+T="MyTitle"    add a title (only works if +H is used)
    Dim strHTML As String
    Dim l As Long
    Dim lTmp As Long
    Dim lTmp2 As Long
    Dim lTmp3 As Long
    Dim lRTFLen As Long
    Dim lBOS As Long                 'beginning of section
    Dim lEOS As Long                 'end of section
    Dim strTmp As String
    Dim strTmp2 As String
    Dim strEOS As String             'string to be added to end of section
    Dim strBOS As String             'string to be added to beginning of section
    Dim strEOP As String             'string to be added to end of paragraph
    Dim strBOL As String             'string to be added to the begining of each new line
    Dim strEOL As String             'string to be added to the end of each new line
    Dim strEOLL As String            'string to be added to the end of previous line
    Dim strCurFont As String         'current font code eg: "f3"
    Dim strCurFontSize As String     'current font size eg: "fs20"
    Dim strCurColor As String        'current font color eg: "cf2"
    Dim strFontFace As String        'Font face for current font
    Dim strFontColor As String       'Font color for current font
    Dim lFontSize As Integer         'Font size for current font
    Const gHellFrozenOver = False    'always false
    Dim gSkip As Boolean             'skip to next word/command
    Dim strCodes As String           'codes for ascii to HTML char conversion
    Dim strCurLine As String         'temp storage for text for current line before being added to strHTML
    Dim strColorTable() As String    'table of colors
    Dim lColors As Long              '# of colors
    Dim strFontTable() As String     'table of fonts
    Dim lFonts As Long               '# of fonts
    Dim strFontCodes As String       'list of font code modifiers
    Dim gSeekingText As Boolean      'True if we have to hit text before inserting a </FONT>
    Dim gText As Boolean             'true if there is text (as opposed to a control code) in strTmp
    Dim strAlign As String           '"center" or "right"
    Dim gAlign As Boolean            'if current text is aligned
    Dim strGen As String             'Temp store for Generator Meta Tag if requested
    Dim strTitle As String           'Temp store for Title if requested

    'setup HTML codes
    strCodes = "&nbsp;  {00}&copy;  {a9}&acute; {b4}&laquo; {ab}&raquo; {bb}&iexcl; {a1}&iquest;{bf}&Agrave;{c0}&agrave;{e0}&Aacute;{c1}"
    strCodes = strCodes & "&aacute;{e1}&Acirc; {c2}&acirc; {e2}&Atilde;{c3}&atilde;{e3}&Auml;  {c4}&auml;  {e4}&Aring; {c5}&aring; {e5}&AElig; {c6}"
    strCodes = strCodes & "&aelig; {e6}&Ccedil;{c7}&ccedil;{e7}&ETH;   {d0}&eth;   {f0}&Egrave;{c8}&egrave;{e8}&Eacute;{c9}&eacute;{e9}&Ecirc; {ca}"
    strCodes = strCodes & "&ecirc; {ea}&Euml;  {cb}&euml;  {eb}&Igrave;{cc}&igrave;{ec}&Iacute;{cd}&iacute;{ed}&Icirc; {ce}&icirc; {ee}&Iuml;  {cf}"
    strCodes = strCodes & "&iuml;  {ef}&Ntilde;{d1}&ntilde;{f1}&Ograve;{d2}&ograve;{f2}&Oacute;{d3}&oacute;{f3}&Ocirc; {d4}&ocirc; {f4}&Otilde;{d5}"
    strCodes = strCodes & "&otilde;{f5}&Ouml;  {d6}&ouml;  {f6}&Oslash;{d8}&oslash;{f8}&Ugrave;{d9}&ugrave;{f9}&Uacute;{da}&uacute;{fa}&Ucirc; {db}"
    strCodes = strCodes & "&ucirc; {fb}&Uuml;  {dc}&uuml;  {fc}&Yacute;{dd}&yacute;{fd}&yuml;  {ff}&THORN; {de}&thorn; {fe}&szlig; {df}&sect;  {a7}"
    strCodes = strCodes & "&para;  {b6}&micro; {b5}&brvbar;{a6}&plusmn;{b1}&middot;{b7}&uml;   {a8}&cedil; {b8}&ordf;  {aa}&ordm;  {ba}&not;   {ac}"
    strCodes = strCodes & "&shy;   {ad}&macr;  {af}&deg;   {b0}&sup1;  {b9}&sup2;  {b2}&sup3;  {b3}&frac14;{bc}&frac12;{bd}&frac34;{be}&times; {d7}"
    strCodes = strCodes & "&divide;{f7}&cent;  {a2}&pound; {a3}&curren;{a4}&yen;   {a5}...     {85}"

    'setup color table
    lColors = 0
    ReDim strColorTable(0)
    lBOS = InStr(strRTF, "\colortbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\red")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strColorTable(lColors)
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 4, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 5, 1)), Mid(strRTF, lBOS + 5, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\green")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 6, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 8, 1)), Mid(strRTF, lBOS + 8, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\blue")
                strTmp = Trim(Hex(Mid(strRTF, lBOS + 5, 1) & IIf(IsNumeric(Mid(strRTF, lBOS + 6, 1)), Mid(strRTF, lBOS + 6, 1), "") & IIf(IsNumeric(Mid(strRTF, lBOS + 7, 1)), Mid(strRTF, lBOS + 7, 1), "")))
                If Len(strTmp) = 1 Then strTmp = "0" & strTmp
                strColorTable(lColors) = strColorTable(lColors) & strTmp
                lBOS = InStr(lBOS, strRTF, "\red")
                lColors = lColors + 1
            Wend
        End If
    End If

    'setup font table
    lFonts = 0
    ReDim strFontTable(0)
    lBOS = InStr(strRTF, "\fonttbl")
    If lBOS <> 0 Then
        lEOS = InStr(lBOS, strRTF, ";}}")
        If lEOS <> 0 Then
            lBOS = InStr(lBOS, strRTF, "\f0")
            While ((lBOS <= lEOS) And (lBOS <> 0))
                ReDim Preserve strFontTable(lFonts)
                While ((Mid(strRTF, lBOS, 1) <> " ") And (lBOS <= lEOS))
                    lBOS = lBOS + 1
                Wend
                lBOS = lBOS + 1
                strTmp = Mid(strRTF, lBOS, InStr(lBOS, strRTF, ";") - lBOS)
                strFontTable(lFonts) = strFontTable(lFonts) & strTmp
                lBOS = InStr(lBOS, strRTF, "\f" & (lFonts + 1))
                lFonts = lFonts + 1
            Wend
        End If
    End If

    strHTML = ""
    lRTFLen = Len(strRTF)
    'seek first line with text on it
    lBOS = InStr(strRTF, vbCrLf & "\deflang")
    If lBOS = 0 Then GoTo finally Else lBOS = lBOS + 2
    lEOS = InStr(lBOS, strRTF, vbCrLf & "\par")
    If lEOS = 0 Then GoTo finally

    While Not gHellFrozenOver
        strTmp = Mid(strRTF, lBOS, lEOS - lBOS)
        l = lBOS
        While l <= lEOS
            strTmp = Mid(strRTF, l, 1)
            Select Case strTmp
                Case "{"
                    l = l + 1
                Case "}"
                    strCurLine = strCurLine & strEOS
                    strEOS = ""
                    l = l + 1
                Case "\"    'special code
                    l = l + 1
                    strTmp = Mid(strRTF, l, 1)
                    Select Case strTmp
                        Case "b"
                            If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                                'b = bold
                                strCurLine = strCurLine & "<B>"
                                strEOS = "</B>" & strEOS
                                If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                            ElseIf (Mid(strRTF, l, 7) = "bullet ") Then
                                strTmp = "•"     'bullet
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "c"
                            If ((Mid(strRTF, l, 2) = "cf") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                                'cf = color font
                                lTmp = Val(Mid(strRTF, l + 2, 5))
                                If lTmp <= UBound(strColorTable) Then
                                    strCurColor = "cf" & lTmp
                                    strFontColor = "#" & strColorTable(lTmp)
                                    gSeekingText = True
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            Else
                                gSkip = True
                            End If
                        Case "e"
                            If (Mid(strRTF, l, 7) = "emdash ") Then
                                strTmp = "—"
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "f"
                            If IsNumeric(Mid(strRTF, l + 1, 1)) Then
                                'f# = font
                                'first get font number
                                lTmp = l + 2
                                strTmp2 = Mid(strRTF, l + 1, 1)
                                While IsNumeric(Mid(strRTF, lTmp, 1))
                                    strTmp2 = strTmp2 & Mid(strRTF, lTmp2, 1)
                                    lTmp = lTmp + 1
                                Wend
                                lTmp = Val(strTmp2)
                                strCurFont = "f" & lTmp
                                If ((lTmp <= UBound(strFontTable)) And (strFontTable(lTmp) <> strFontTable(0))) Then
                                    'insert codes if lTmp is a valid font # AND the font is not the default font
                                    strFontFace = strFontTable(lTmp)
                                    gSeekingText = True
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            ElseIf ((Mid(strRTF, l + 1, 1) = "s") And (IsNumeric(Mid(strRTF, l + 2, 1)))) Then
                                'fs# = font size
                                'first get font size
                                lTmp = l + 3
                                strTmp2 = Mid(strRTF, l + 2, 1)
                                While IsNumeric(Mid(strRTF, lTmp, 1))
                                    strTmp2 = strTmp2 & Mid(strRTF, lTmp, 1)
                                    lTmp = lTmp + 1
                                Wend
                                lTmp = Val(strTmp2)
                                strCurFontSize = "fs" & lTmp
                                lFontSize = Int((lTmp / 5) - 2)
                                If lFontSize = 2 Then
                                    strCurFontSize = ""
                                    lFontSize = 0
                                Else
                                    gSeekingText = True
                                    If lFontSize > 8 Then lFontSize = 8
                                    If lFontSize < 1 Then lFontSize = 1
                                End If
                                'move "cursor" position to next rtf code
                                lTmp = l
                                While ((Mid(strRTF, lTmp, 1) <> " ") And (Mid(strRTF, lTmp, 1) <> "\"))
                                    lTmp = lTmp + 1
                                Wend
                                If (Mid(strRTF, lTmp, 1) = " ") Then
                                    l = lTmp
                                Else
                                    l = lTmp - 1
                                End If
                            Else
                                gSkip = True
                            End If
                        Case "i"
                            If ((Mid(strRTF, l + 1, 1) = " ") Or (Mid(strRTF, l + 1, 1) = "\")) Then
                                strCurLine = strCurLine & "<I>"
                                strEOS = "</I>" & strEOS
                                If (Mid(strRTF, l + 1, 1) = " ") Then l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "l"
                            If (Mid(strRTF, l, 10) = "ldblquote ") Then
                                'left doublequote
                                strTmp = "“"
                                l = l + 9
                                gText = True
                            ElseIf (Mid(strRTF, l, 7) = "lquote ") Then
                                'left quote
                                strTmp = "‘"
                                l = l + 6
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "p"
                            If ((Mid(strRTF, l, 6) = "plain\") Or (Mid(strRTF, l, 6) = "plain ")) Then
                                If (Len(strFontColor & strFontFace) > 0) Then
                                    If Not gSeekingText Then strCurLine = strCurLine & "</FONT>"
                                    strFontColor = ""
                                    strFontFace = ""
                                End If
                                If gAlign Then
                                    strCurLine = strCurLine & "</TD></TR></TABLE><BR>"
                                    gAlign = False
                                End If
                                strCurLine = strCurLine & strEOS
                                strEOS = ""
                                If Mid(strRTF, l + 5, 1) = "\" Then l = l + 4 Else l = l + 5    'catch next \ but skip a space
                            ElseIf (Mid(strRTF, l, 9) = "pnlvlblt\") Then
                                'bulleted list
                                strEOS = ""
                                strBOS = "<UL>"
                                strBOL = "<LI>"
                                strEOL = "</LI>"
                                strEOP = "</UL>"
                                l = l + 7    'catch next \
                            ElseIf (Mid(strRTF, l, 7) = "pntext\") Then
                                l = InStr(l, strRTF, "}")   'skip to end of braces
                            ElseIf (Mid(strRTF, l, 6) = "pntxtb") Then
                                l = InStr(l, strRTF, "}")   'skip to end of braces
                            ElseIf (Mid(strRTF, l, 10) = "pard\plain") Then
                                strCurLine = strCurLine & strEOS & strEOP
                                strEOS = ""
                                strEOP = ""
                                strBOL = ""
                                strEOL = "<BR>"
                                l = l + 3    'catch next \
                            Else
                                gSkip = True
                            End If
                        Case "q"
                            If ((Mid(strRTF, l, 3) = "qc\") Or (Mid(strRTF, l, 3) = "qc ")) Then
                                'qc = centered
                                strAlign = "center"
                                'move "cursor" position to next rtf code
                                If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                                l = l + 1
                            ElseIf ((Mid(strRTF, l, 3) = "qr\") Or (Mid(strRTF, l, 3) = "qr ")) Then
                                'qr = right justified
                                strAlign = "right"
                                'move "cursor" position to next rtf code
                                If (Mid(strRTF, l + 2, 1) = " ") Then l = l + 2
                                l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "r"
                            If (Mid(strRTF, l, 7) = "rquote ") Then
                                'reverse quote
                                strTmp = "’"
                                l = l + 6
                                gText = True
                            ElseIf (Mid(strRTF, l, 10) = "rdblquote ") Then
                                'reverse doublequote
                                strTmp = "”"
                                l = l + 9
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "s"
                            'strikethrough
                            If ((Mid(strRTF, l, 7) = "strike\") Or (Mid(strRTF, l, 7) = "strike ")) Then
                                strCurLine = strCurLine & "<STRIKE>"
                                strEOS = "</STRIKE>" & strEOS
                                l = l + 6
                            Else
                                gSkip = True
                            End If
                        Case "t"
                            If (Mid(strRTF, l, 4) = "tab ") Then
                                strTmp = "&#9;"   'tab
                                l = l + 2
                                gText = True
                            Else
                                gSkip = True
                            End If
                        Case "u"
                            'underline
                            If ((Mid(strRTF, l, 3) = "ul ") Or (Mid(strRTF, l, 3) = "ul\")) Then
                                strCurLine = strCurLine & "<U>"
                                strEOS = "</U>" & strEOS
                                l = l + 1
                            Else
                                gSkip = True
                            End If
                        Case "'"
                            'special characters
                            strTmp2 = "{" & Mid(strRTF, l + 1, 2) & "}"
                            lTmp = InStr(strCodes, strTmp2)
                            If lTmp = 0 Then
                                strTmp = Chr("&H" & Mid(strTmp2, 2, 2))
                            Else
                                strTmp = Trim(Mid(strCodes, lTmp - 8, 8))
                            End If
                            l = l + 1
                            gText = True
                        Case "~"
                            strTmp = " "
                            gText = True
                        Case "{", "}", "\"
                            gText = True
                        Case vbLf, vbCr, vbCrLf    'always use vbCrLf
                            strCurLine = strCurLine & vbCrLf
                        Case Else
                            gSkip = True
                    End Select
                    If gSkip = True Then
                        'skip everything up until the next space or "\" or "}"
                        While InStr(" \}", Mid(strRTF, l, 1)) = 0
                            l = l + 1
                        Wend
                        gSkip = False
                        If (Mid(strRTF, l, 1) = "\") Then l = l - 1
                    End If
                    l = l + 1
                Case vbLf, vbCr, vbCrLf
                    l = l + 1
                Case Else
                    gText = True
            End Select
            If gText Then
                If ((Len(strFontColor & strFontFace) > 0) And gSeekingText) Then
                    If Len(strAlign) > 0 Then
                        gAlign = True
                        If strAlign = "center" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""center""><TD>"
                        ElseIf strAlign = "right" Then
                            strCurLine = strCurLine & "<TABLE ALIGN=""left"" CELLSPACING=0 CELLPADDING=0 WIDTH=""100%""><TR ALIGN=""right""><TD>"
                        End If
                        strAlign = ""
                    End If
                    If Len(strFontFace) > 0 Then
                        strFontCodes = strFontCodes & " FACE=" & strFontFace
                    End If
                    If Len(strFontColor) > 0 Then
                        strFontCodes = strFontCodes & " COLOR=" & strFontColor
                    End If
                    If Len(strCurFontSize) > 0 Then
                        strFontCodes = strFontCodes & " SIZE = " & lFontSize
                    End If
                    strCurLine = strCurLine & "<FONT" & strFontCodes & ">"
                    strFontCodes = ""
                End If
                strCurLine = strCurLine & strTmp
                l = l + 1
                gSeekingText = False
                gText = False
            End If
        Wend

        lBOS = lEOS + 2
        lEOS = InStr(lEOS + 1, strRTF, vbCrLf & "\par")
        strHTML = strHTML & strEOLL & strBOS & strBOL & strCurLine & vbCrLf
        strEOLL = strEOL
        If Len(strEOL) = 0 Then strEOL = "<BR>"

        If lEOS = 0 Then GoTo finally
        strBOS = ""
        strCurLine = ""
    Wend

finally:
    strHTML = strHTML & strEOS
    'clear up any hanging fonts
    If (Len(strFontColor & strFontFace) > 0) Then strHTML = strHTML & "</FONT>" & vbCrLf

    'Add Generator Metatag if requested
    If InStr(strOptions, "+G") <> 0 Then
        strGen = "<META NAME=""GENERATOR"" CONTENT=""RTF2HTML by Brady Hegberg"">"
    Else
        strGen = ""
    End If

    'Add Title if requested
    If InStr(strOptions, "+T") <> 0 Then
        lTmp = InStr(strOptions, "+T") + 3
        lTmp2 = InStr(lTmp + 1, strOptions, """")
        strTitle = Mid(strOptions, lTmp, lTmp2 - lTmp)
    Else
        strTitle = ""
    End If

    'add header and footer if requested
    If InStr(strOptions, "+H") <> 0 Then strHTML = strHeader & vbCrLf _
            & strHTML _
            & strFooter
    RTF2HTML = strHTML
End Function


Sub FindText()

    On Local Error GoTo SalirFindText
    
    Dim lWhere, lPos As Long
    Dim sTmp As String
    Dim Sql As String
            
    Sql = UCase$(glbFindSql)
    
    If gbLastPos = 0 Or gbLastPos > Len(Sql) Then
        lPos = 1
    Else
        lPos = gbLastPos
    End If
        
    Do While lPos < Len(Sql)
        
        sTmp = Mid(Sql, lPos, Len(Sql))
        
        lWhere = InStr(sTmp, UCase$(gsFindText))
        lPos = lPos + lWhere
        
        If lWhere Then   ' If found,
            frmMain.SetFocus
            frmMain.rtbCodigo.SelStart = lPos - 2   ' set selection start and
            frmMain.rtbCodigo.SelLength = Len(gsFindText)   ' set selection length.   Else
            gbLastPos = lPos
            Exit Do
        Else
            gbLastPos = 0
            Exit Do 'we are ready
        End If
    Loop
    
    Exit Sub
    
SalirFindText:
    gbLastPos = 0
    Err = 0
    
End Sub

Function EmptyString(ByRef sText As String) As Boolean
   If IsNull(sText) Then
      EmptyString = True
   Else
      EmptyString = (Len(Trim(sText)) = 0)
   End If
End Function
Function AttachPath(sFileName As String, sPath As String) As String
   If Len(Trim(ExtractPath(sFileName))) = 0 Then
      AttachPath = FixPath(sPath) & sFileName
   Else
      AttachPath = sFileName
   End If
End Function
Function FixPath(ByVal sPath As String) As String
   If Len(Trim(sPath)) = 0 Then
      FixPath = ""
   ElseIf Right$(sPath, 1) <> "\" Then
      FixPath = sPath & "\"
   Else
      FixPath = sPath
   End If
End Function


Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Public Sub Copiar(ByVal hWnd As Long)

    Dim ret As Long
    
    ret = SendMessage(hWnd, WM_COPY, 0, 0)
    
End Sub

Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function

