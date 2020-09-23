Attribute VB_Name = "basTranslate"
'    C2VB  Converts C style definitions to VB
'    Copyright (C) 2000  Kimon Andreou (kimon@mindless.com)
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit
Option Base 0

Public IsPublic As Boolean
Public LibraryName As String
Public Const vbTab As String = "    "

'Splits a string into tokens based on the delimiter specified
'strString: The source string
'TokenID:   The token wanted from the string
'Token:     The token returned
'Delimiter: The delimiter to use (default is " ")
'returns an array of strings where each item is a token
Public Function GetToken(strString As String, TokenID As Integer, Token As String, Optional Delimiter As String = " ") As Variant
Dim Tokens() As String
Dim UpperLimit As Integer

Tokens = Split(strString, Delimiter)
UpperLimit = UBound(Tokens)

If UpperLimit = -1 Then
    Token = ""
    GetToken = Tokens
    Exit Function
End If
If TokenID > UpperLimit Then
    Token = Tokens(UpperLimit)
    GetToken = Tokens
Else
    GetToken = Tokens
    Token = Tokens(TokenID)
End If

End Function

'Joins an array of strings with a delimiter
'I wrote this in case I wanted to add functionality not available with Join()
Public Function ConcatString(strStrings() As String, Optional Delimiter As String = " ") As String
Dim cnt As Integer

If UBound(strStrings) = 0 Then
    ConcatString = strStrings(0)
    Exit Function
End If

For cnt = 0 To UBound(strStrings) - 1
    ConcatString = ConcatString & strStrings(cnt) & Delimiter
Next cnt
ConcatString = ConcatString & strStrings(UBound(strStrings))
End Function

'Adds spaces around "," "(" ")"
Public Function AddSpaces(strString As String) As String
Dim dummy1 As String
Dim dummy2 As String
Dim strArray() As String

strArray = GetToken(strString, 1, dummy1, "(")
dummy1 = ConcatString(strArray, " ( ")
strArray = GetToken(dummy1, 1, dummy2, ",")
dummy1 = ConcatString(strArray, " , ")
strArray = GetToken(dummy1, 1, dummy2, ")")
dummy1 = ConcatString(strArray, " ) ")
AddSpaces = dummy1

End Function

'Determines what is the VB equivalent data type
'If it cannot figure it out, it returns the C name (useful for UDTs)
'Also determines if it is an array or a pointer (ByRef)
'and if it is a function or a procedure
Public Function GetType(strArray() As String, Optional IsProcedure As Boolean, Optional IsByVal As Boolean, Optional IsArray As Boolean) As String
Dim Ctype As String
Dim Types() As String
Dim dummy As String
Dim cnt As Integer

'Get rid of in-between spaces
For cnt = 0 To UBound(strArray)
    strArray(cnt) = Trim(strArray(cnt))
    Ctype = Ctype & strArray(cnt) & IIf(Len(strArray(cnt)) = 0, "", " ")
Next cnt

Types = GetToken(Trim(Ctype), 0, dummy)

Ctype = Join(Types, " ")
IsProcedure = False

IsByVal = True

'Check if we should pass ByRef or not and whether it is an array
If InStr(Ctype, "*") <> 0 Then IsByVal = False
If InStr(Ctype, "**") <> 0 Then IsArray = True
If InStr(Ctype, "[") <> 0 Then
    IsByVal = False
    IsArray = True
End If
If InStr(Ctype, "const") <> 0 Then
    IsByVal = True
    Ctype = Trim(RemoveChar(Ctype, "const"))
End If
If InStr(Ctype, "CONST") <> 0 Then
    IsByVal = True
    Ctype = Trim(RemoveChar(Ctype, "CONST"))
End If
If Left(Ctype, 2) = "LP" Then IsByVal = False

'Get rid of "*", "[", "]", "WINAPI", "CALLBACK"
Ctype = Trim(RemoveChar(Ctype, "*"))
Ctype = Trim(RemoveChar(Ctype, "["))
Ctype = Trim(RemoveChar(Ctype, "]"))
Ctype = Trim(RemoveChar(Ctype, "WINAPI"))
Ctype = Trim(RemoveChar(Ctype, "CALLBACK"))


'This is where the actual checking happens
'maybe saving this in an external file would be be better...
'i'll have to think about it
Select Case Ctype
    Case "int":
        GetType = "Long"
    Case "time_t":
        GetType = "Long"
    Case "size_t":
        GetType = "Long"
    Case "unsigned int":
        GetType = "Long"
    Case "unsigned char":
        GetType = "Byte"
    Case "void":
        IsProcedure = True
        GetType = ""
    Case "VOID":
        IsProcedure = True
        GetType = ""
    Case "char":
        GetType = IIf((Not IsByVal) Or IsArray, "String", "Byte")
        
    Case "ATOM":
        GetType = "Long"
        
    Case "BOOL":
        GetType = "Long"
    Case "BOOLEAN":
        GetType = "Long"
    Case "BYTE":
        GetType = "Byte"
        
    Case "CHAR":
        GetType = "Byte"
    Case "COLORREF":
        GetType = "Long"
    Case "CRITICAL_SECTION":
        GetType = "Long"
    Case "CTRYID":
        GetType = "Long"
        
    Case "DWORD":
        GetType = "Long"
    Case "DWORD_PTR":
        GetType = "Double"
    Case "DWORD32":
        GetType = "Long"
    Case "DWORD64":
        GetType = "Currency"
        
    Case "FLOAT":
        GetType = "Double"
    Case "FILE_SEGMENT_ELEMENT":
        GetType = "Currency"
        
    Case "HACCEL":
        GetType = "Long"
    Case "HANDLE":
        GetType = "Long"
    Case "HBITMAP"
        GetType = "Long"
    Case "HBRUSH":
        GetType = "Long"
    Case "HCOLORSPACE":
        GetType = "Long"
    Case "HCONV":
        GetType = "Long"
    Case "HCONVLIST":
        GetType = "Long"
    Case "HCURSOR":
        GetType = "Long"
    Case "HDC":
        GetType = "Long"
    Case "HDDEDATA":
        GetType = "Long"
    Case "HDESK":
        GetType = "Long"
    Case "HDROP":
        GetType = "Long"
    Case "HDWP":
        GetType = "Long"
    Case "HENHMETAFILE":
        GetType = "Long"
    Case "HFILE":
        GetType = "Long"
    Case "HFONT"
        GetType = "Long"
    Case "HGDIOBJ":
        GetType = "Long"
    Case "HGLOBAL":
        GetType = "Long"
    Case "HHOOK":
        GetType = "Long"
    Case "HICON":
        GetType = "Long"
    Case "HIMAGELIST":
        GetType = "Long"
    Case "HIMC":
        GetType = "Long"
    Case "HINSTANCE":
        GetType = "Long"
    Case "HKEY":
        GetType = "Long"
    Case "HKL":
        GetType = "Long"
    Case "HLOCAL":
        GetType = "Long"
    Case "HMENU":
        GetType = "Long"
    Case "HMETAFILE":
        GetType = "Long"
    Case "HMODULE":
        GetType = "Long"
    Case "HMONITOR":
        GetType = "Long"
    Case "HPALETTE":
        GetType = "Long"
    Case "HPEN":
        GetType = "Long"
    Case "HRESULT":
        GetType = "Long"
    Case "HRGN":
        GetType = "Long"
    Case "HRSRC":
        GetType = "Long"
    Case "HSZ":
        GetType = "Long"
    Case "HWINSTA":
        GetType = "Long"
    Case "HWND":
        GetType = "Long"
        
    Case "INT":
        GetType = "Long"
    Case "INT_PTR":
        GetType = "Long"
    Case "INT32":
        GetType = "Long"
    Case "INT64":
        GetType = "Currency"
    Case "IPADDR":
        GetType = "Long"
    Case "IPMASK":
        GetType = "Long"
        
    Case "LANGID":
        GetType = "Long"
    Case "LCID":
        GetType = "Long"
    Case "LONG":
        GetType = "Long"
    Case "LONG_PTR":
        GetType = "Long"
    Case "LONG32":
        GetType = "Long"
    Case "LONG64":
        GetType = "Currency"
    Case "LONGLONG":
        GetType = "Currency"
    Case "LPARAM":
        GetType = "Any"
        
    Case "LPBOOL":
        GetType = "Long"
    Case "LPBYTE":
        GetType = "Long"
    Case "LPCOLORREF":
        GetType = "Long"
    Case "LPCRITICAL":
        GetType = "Long"
    Case "LPCSTR":
        GetType = "String"
        IsByVal = True
    Case "LPCTSTR":
        GetType = "String"
        IsByVal = True
    Case "LPCVOID":
        GetType = "Any"
        IsByVal = True
    Case "LPDWORD":
        GetType = "Long"
    Case "LPHANDLE"
        GetType = "Long"
    Case "LPINT":
        GetType = "Long"
    Case "LPLONG":
        GetType = "Long"
    Case "LPSTR":
        GetType = "String"
    Case "LPTSTR":
        GetType = "String"
        IsByVal = False
    Case "LPVOID":
        GetType = "Any"
        IsByVal = False
    Case "LPWORD":
        GetType = "Long"
        
    Case "LRESULT":
        GetType = "Long"
    Case "LUID":
        GetType = "Long"
        
    Case "PBOOL":
        IsByVal = False
        GetType = "Long"
    Case "PBOOLEAN":
        IsByVal = False
        GetType = "Long"
    Case "PBYTE":
        IsByVal = False
        GetType = "Long"
    Case "PCHAR":
        IsByVal = False
        GetType = "Long"
    Case "PCRITICAL_SECTION":
        IsByVal = False
        GetType = "Long"
    Case "PCSTR":
        GetType = "String"
        IsByVal = True
    Case "PCTSTR:"
        GetType = "String"
        IsByVal = True
    Case "PDWORD":
        IsByVal = False
        GetType = "Long"
    Case "PFLOAT":
        IsByVal = False
        GetType = "Long"
    Case "PHANDLE":
        IsByVal = False
        GetType = "Long"
    Case "PHKEY":
        IsByVal = False
        GetType = "Long"
    Case "PINT":
        IsByVal = False
        GetType = "Long"
    Case "PLCID":
        IsByVal = False
        GetType = "Long"
    Case "PLONG":
        IsByVal = False
        GetType = "Long"
    Case "PLUID":
        IsByVal = False
        GetType = "Long"
    Case "POINTER_32":
        IsByVal = False
        GetType = "Long"
    Case "POINTER_64":
        IsByVal = False
        GetType = "Currency"
    Case "PSHORT":
        IsByVal = False
        GetType = "Long"
    Case "PSTR":
        IsByVal = False
        GetType = "String"
    Case "PTBYTE":
        IsByVal = False
        GetType = "Long"
    Case "PTCHAR":
        IsByVal = False
        GetType = "Long"
    Case "PUCHAR":
        IsByVal = False
        GetType = "Long"
    Case "PUINT":
        IsByVal = False
        GetType = "Long"
    Case "PULONG":
        IsByVal = False
        GetType = "Long"
    Case "PUSHORT":
        IsByVal = False
        GetType = "Long"
    Case "PVOID":
        IsByVal = False
        GetType = "Any"
    Case "PWORD":
        IsByVal = False
        GetType = "Long"
    
    Case "SERVICE_STATUS_HANDLE":
        GetType = "Long"
    Case "SHORT":
        GetType = "Integer"
    Case "SIZE_T":
        GetType = "Long"
    Case "SSIZE_T":
        GetType = "Long"
    
    Case "TBYTE":
        GetType = "Byte"
    Case "TCHAR":
        GetType = "Byte"
    
    Case "UCHAR":
        GetType = "Byte"
    Case "UINT":
        GetType = "Long"
    Case "UINT_PTR":
        GetType = "Long"
    Case "UINT32":
        GetType = "Long"
    Case "UINT64":
        GetType = "Currency"
    Case "ULONG":
        GetType = "Long"
    Case "ULONG_PTR":
        GetType = "Long"
    Case "ULONG32":
        GetType = "Long"
    Case "ULONG64":
        GetType = "Currency"
    Case "ULONGLONG":
        GetType = "Currency"
    Case "USHORT":
        GetType = "Integer"
    
    Case "VOID":
        GetType = "Any"
    
    Case "WORD":
        GetType = "Integer"
    Case "WPARAM":
        GetType = "Long"
    
    Case Else:
        IsByVal = False
        GetType = Ctype
End Select
        
End Function

'Removes a specific character from a string
Public Function RemoveChar(strString As String, strChar As String) As String
Dim dummy As String
Dim Tokens() As String

Tokens = GetToken(strString, 0, dummy, strChar)
RemoveChar = Join(Tokens)


End Function

'Converts from C style Octal notation to VB style Octal notation
Public Function Coct2VBoct(strCoct As String) As String

Coct2VBoct = "&O" & Mid(strCoct, 3)

End Function

'Converts from C style Hexadecimal notation to VB style Hexadecimal notation
Public Function Chex2VBhex(strChex As String) As String

Chex2VBhex = "&H" & UCase(Mid(strChex, 3))

End Function
'Gets rid of C style comments
Private Function StripComments(Cstring As String) As String
Dim cnt As Long
Dim strOut As String
Dim dummy As String

'Flags to see whether we are in a comment
Dim InEOLComment As Boolean     'In a "//" style comment
Dim InCComment As Boolean       'In a "/*  */"  style comment

'Flags to check if a comment is starting or ending
Dim CheckingEOLComment As Boolean
Dim CheckingCComment As Boolean
Dim CheckingEndCComment As Boolean


strOut = ""                                 'Initialize
Cstring = Trim(Cstring)                     'Trim spaces

'Loop through the characters
For cnt = 1 To Len(Cstring)
    dummy = Mid(Cstring, cnt, 1)
    
    Select Case dummy
        Case vbCr:
            
        Case vbLf:
            If InEOLComment Then
                InEOLComment = False
                strOut = strOut & vbCrLf
            End If
            
        Case "/":
            If CheckingEOLComment Then
                CheckingCComment = False
                CheckingEOLComment = False
                InEOLComment = True
            Else
                If CheckingEndCComment Then
                    CheckingEndCComment = False
                    InCComment = False
                    dummy = ""
                Else
                    CheckingCComment = True
                    CheckingEOLComment = True
                End If
            End If
            
            
        Case "*":
            If CheckingCComment Then
                CheckingCComment = False
                InCComment = True
                CheckingEOLComment = False
            Else
                CheckingEndCComment = InCComment
            End If
            
        Case Else:
            CheckingCComment = False
            CheckingEndCComment = False
            CheckingEOLComment = False
    End Select
    
    If Not (CheckingCComment Or CheckingEOLComment Or CheckingEndCComment Or _
        InCComment Or InEOLComment) Then strOut = strOut & dummy
Next cnt

StripComments = strOut

End Function

'Finds all the tabs in a string and replaces them with spaces
Private Function ReplaceTabs(strString As String) As String
Dim dummies() As String
Dim dummy As String

dummies = GetToken(strString, 0, dummy, vbTab)
ReplaceTabs = Join(dummies, " ")
End Function

'Determines what type of C instruction we're getting and processes it
'accordingly.
Private Function Translate(ByVal Cstring As String) As String
Dim TypeOfDeclaration As String
Dim strC As String
Dim dummy As String

strC = ReplaceTabs(Cstring)             'Get rid of the tabs
strC = Trim(StripComments(strC))        'Get rid of the comments
strC = RemoveChar(strC, "/*")           'Get rid of any comment identifiers
strC = RemoveChar(strC, "*/")           'Same as above

'I could've done a RemoveChar(strC, vbCrLf) but *NIX style text files
'don't have CRs.
strC = RemoveChar(strC, vbCr)           'Get rid of Carriage Returns
strC = RemoveChar(strC, vbLf)           'Get rid of Line Feeds


strC = Trim(strC)
If strC = "" Then
    Translate = ""
    Exit Function
End If

dummy = ""
GetToken strC, 0, TypeOfDeclaration         'Get the identifier

If Left(TypeOfDeclaration, 1) = "#" Then    'Is it a preprocessor instr.?
    dummy = ProcessPreProcessor(strC)
Else
    Select Case TypeOfDeclaration       'The main checker ;-)
        Case "":
            'Don't do anything
        
        Case "typedef":                     'Is it a type definition?
            dummy = ProcessTypeDef(strC) & vbCrLf
            
        Case "extern":                      'Is it an external var definition?
            dummy = "#If False Then" & vbCrLf
            dummy = dummy & "'The following is an externally defined global variable" & vbCrLf
            dummy = dummy & strC & vbCrLf
            
        Case "enum":                        'An enumeration
            GetToken strC, 1, dummy
            strC = "typedef " & Left(strC, Len(strC) - 1) & dummy & ";"
            dummy = ProcessTypeDef(strC) & vbCrLf
            
        Case "struct":                      'A structure
            GetToken strC, 1, dummy
            strC = "typedef " & Left(strC, Len(strC) - 1) & dummy & ";"
            dummy = ProcessTypeDef(strC) & vbCrLf
            
        Case "union":                       'A union
            Dim parts() As String
            parts = Split(strC, " ")
            parts(0) = "typedef struct"     'For VB, unions are structures
            parts(UBound(parts)) = Left(parts(UBound(parts)), Len(parts(UBound(parts))) - 1) & _
                parts(1) & ";"
            strC = Join(parts, " ")
            dummy = ProcessTypeDef(strC) & vbCrLf
            
        Case Else:                          'Anything else is a function (right?)
            strC = RemoveChar(strC, "WINAPI")   'We don't really need this
            Dim cnt As Integer
            
            'If we get a __declspec(dllimport)  get rid of it
            parts = Split(strC, "__declspec")
            For cnt = 0 To UBound(parts)
                parts(cnt) = Trim(parts(cnt))
            Next cnt
            strC = Trim(Join(parts, ""))
            parts = Split(strC, "dllimport")
            For cnt = 0 To UBound(parts)
                parts(cnt) = Trim(parts(cnt))
            Next cnt
            strC = Trim(Join(parts, ""))
            parts = Split(strC, "dllexport")
            For cnt = 0 To UBound(parts)
                parts(cnt) = Trim(parts(cnt))
            Next cnt
            strC = Trim(Join(parts, ""))
            parts = Split(strC, "()")
            strC = Trim(Join(parts, ""))
                
            'Ok, we're ready to translate now
            dummy = ProcessFunction(strC) & vbCrLf
    End Select
End If

Translate = dummy
End Function

'Manage the overall "translation"
'Arguments:
'strString:         Input string or filename
'strOut:            Output string
'IsFile:            Flag setting if processing a file or not
'strOutputFilename: What is the output filename
Public Sub Process(strString As String, Optional strOut As String, Optional IsFile As Boolean = False, Optional strOutputFileName As String)
Dim cnt As Integer
Dim CStrings() As String
Dim BasStrings() As String
Dim dummy As String

'Split string or file into individual instructions
CStrings = SplitInstructions(strString, IsFile)
ReDim BasStrings(UBound(CStrings))

'Translate every C instruction into its VB equivalent
For cnt = 0 To UBound(CStrings)
    dummy = Trim(CStrings(cnt))
    BasStrings(cnt) = IIf(dummy = "", "", Translate(dummy))
    
    'Yeah, yeah, I know... this isn't the best way to do it, but it works fine
    DoEvents
Next cnt


If IsFile Then
    OutputToFile strOutputFileName, BasStrings
Else
    strOut = Join(BasStrings, vbCrLf)
End If
End Sub

'Translates a C variable definition into a VB variable definition
'Arguments
'strArgument:   The variable definition in C
'Returns
'ArgName:       The variable name
'ArgType:       The VB data type
'IgnoreByVal:   If we don't care whether the variable should be passed ByVal
'Incrementer:   If we want to append an increment number, or any other postfix
'
'e.g. If strArgument="int argc"
'ArgName = "argc"
'ArgType = " As Long"
Public Sub ProcessArg(strArgument As String, ArgName As String, ArgType As String, Optional IgnoreByVal As Boolean = False, Optional Increment As String = "")
Dim parts() As String
Dim cnt As Integer
Dim dummy As String
Dim TypePart() As String
Dim IsByVal As Boolean
Dim IsArray As Boolean

'Flags that determine whether the variable should be passed ByVar or is an array
Dim GetTypeIsByVal As Boolean
Dim GetTypeIsArray As Boolean


'Split the definition
parts = GetToken(strArgument, 0, dummy)
IsByVal = True      'Initialize

'This is where we find out if the variable should be passed ByRef and/or is an array
If InStr(strArgument, "*") <> 0 Then IsByVal = IIf(InStr(strArgument, "const") <> 0, True, False)
If InStr(strArgument, "**") <> 0 Then IsArray = True
If InStr(strArgument, "[") <> 0 Then
    IsByVal = False
    IsArray = True
End If

'If we don't have a variable name, or no arguments handle it accordingly
If UBound(parts) = 0 Then
    If parts(0) = "void" Then
        ArgName = ""
        ArgType = ""
    Else
        ArgType = GetType(parts, , GetTypeIsByVal, GetTypeIsArray)
        IsByVal = IsByVal Or GetTypeIsByVal
        IsArray = IsArray Or GetTypeIsArray
        ArgType = " As " & IIf(((Not IsByVal) Or IsArray) And (ArgType = "Byte"), "String", ArgType)
        ArgName = IIf(IgnoreByVal, "", IIf(IsByVal, "ByVal ", "ByRef ")) & _
            "Arg" & Increment & IIf(IsArray, "()", "")
    End If
    Exit Sub
End If


ReDim TypePart(UBound(parts) - 1)
For cnt = 0 To UBound(TypePart)
    TypePart(cnt) = parts(cnt)
Next cnt
ArgType = GetType(TypePart, , GetTypeIsByVal, GetTypeIsArray)

ArgName = parts(UBound(parts))
ArgName = RemoveChar(ArgName, "*")
ArgName = RemoveChar(ArgName, "[")
ArgName = RemoveChar(ArgName, "]")
ArgName = Trim(ArgName)

IsByVal = IsByVal And GetTypeIsByVal
IsArray = IsArray Or GetTypeIsArray
ArgType = " As " & IIf(((Not IsByVal) Or IsArray Or GetTypeIsByVal) And (ArgType = "Byte"), "String", ArgType)
ArgName = IIf(IgnoreByVal, "", IIf(IsByVal, "ByVal ", "ByRef ")) & _
    IIf(ArgName = "", "Arg" & Increment, ArgName) & _
    IIf(IsArray, "()", "")

End Sub

'This is the function that splits a string or a text file into individual instructions
'Arguments
'strSource:     The string/filename to process
'IsFileName:    Flag setting.  True=Is file, False=Not file
'
'Returns
'An array of strings each item being an individual instruction

'Remember that the purpose is to split the text into individual instructions
'and not to handle them here.  That is done elsewhere.
Public Function SplitInstructions(strSource As String, Optional IsFileName = False) As Variant
Dim cnt As Integer
Dim Instructions As Integer
Dim SectionStarts As Integer
Dim dummy As String
Dim Instrc() As String

'Flags
Dim InInstr As Boolean
Dim InSection As Boolean
Dim InPreProcessor As Boolean
Dim InCComment As Boolean
Dim CheckingCComment As Boolean
Dim CheckingEndCComment As Boolean


'If it's a file, process accordingly
If IsFileName Then
    Instrc = ProcessFile(strSource)
Else  'Otherwise
    ReDim Instrc(1)                     'Resize output array
    For cnt = 1 To Len(strSource)       'Loop through every character
        dummy = Mid(strSource, cnt, 1)
        
        'Get rid of initial CRs and LFs
        If Len(Instrc(Instructions)) = 0 Then
            If (dummy = vbCr) Or (dummy = vbLf) Then dummy = ""
        End If
        
        'Append character to string
        Instrc(Instructions) = Instrc(Instructions) & dummy
        
        'Handle character
        Select Case dummy
            Case "/":                             'Start or end of comment
                If CheckingCComment Then          'Are we expecting a comment start?
                    CheckingCComment = False
                Else
                    If CheckingEndCComment Then   'Are we expecting a cooment end?
                        InCComment = False
                        CheckingEndCComment = False
                    Else                          'Expect a comment start.
                        CheckingCComment = True
                    End If
                End If
                
            Case "*":                              'Same as "/" more or less
                If CheckingCComment Then
                    CheckingCComment = False
                    InCComment = True
                Else
                    If InCComment Then
                        CheckingEndCComment = True
                    End If
                End If
                    
            Case "#":                             'Is it a preprocessor instr?
                'If we are within a comment, it doesn't matter, does it?
                If Not InCComment Then InPreProcessor = True
                
            Case vbLf:                            'Has line ended?
                'This matters, because preprocessor instr. end at a LF
                If Not InCComment Then
                    If InPreProcessor Then
                        InInstr = False
                        Instructions = Instructions + 1
                        ReDim Preserve Instrc(Instructions)
                        InPreProcessor = False
                        Instrc(Instructions) = ""
                    End If
                End If
                
            Case ";":                             'This is how C instr. end
                If Not InCComment Then
                    If (Not InSection) And InInstr Then
                        InInstr = False
                        Instructions = Instructions + 1
                        ReDim Preserve Instrc(Instructions)
                        Instrc(Instructions) = ""
                    End If
                End If
                
            Case "{":
            'We mustn't forget the times we have multiple instr. within
            '{ and }
                If Not InCComment Then
                    SectionStarts = SectionStarts + 1
                    InSection = True
                    InInstr = True
                End If
                
            Case "}":       'Has a segment ended?
                If Not InCComment Then
                    'Just checks to see if we have as many }'s as we do {'s
                    SectionStarts = SectionStarts - 1
                    If SectionStarts = 0 Then
                        InInstr = True
                        InSection = False
                    End If
                End If
                
            Case Else:      'Handle everything else
                If Not InCComment Then InInstr = True
                CheckingCComment = False
                CheckingEndCComment = False
        End Select
    Next cnt
End If


SplitInstructions = Instrc
End Function
