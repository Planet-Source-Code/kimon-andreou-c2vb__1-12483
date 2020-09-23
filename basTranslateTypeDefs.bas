Attribute VB_Name = "basTranslateTypeDefs"
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

'Convert a C style enumeration to a VB style enumeration
Private Function ProcessEnum(Cstring As String) As String
Dim Members() As String
Dim dummy As String
Dim cnt As Integer
Dim MemberName As String
Dim MemberValue As String

'Initialize
Cstring = RemoveChar(Cstring, " ")
Members = GetToken(Cstring, 0, dummy, ",")
dummy = ""

'For each member of the enumeration, translate to VB
For cnt = 0 To UBound(Members)
    'Check if the members have assigned values or not
    If InStr(Members(cnt), "=") Then
        GetToken Members(cnt), 0, MemberName, "="
        GetToken Members(cnt), 1, MemberValue, "="
        MemberValue = Trim(MemberValue)
        If InStr(MemberValue, "0x") <> 0 Then
            MemberValue = Chex2VBhex(MemberValue)
        End If
        If InStr(MemberValue, "0o") <> 0 Then
            MemberValue = Coct2VBoct(MemberValue)
        End If
        dummy = dummy & vbTab & Trim(MemberName) & " = " & Trim(MemberValue) & vbCrLf
    Else
        dummy = dummy & vbTab & Trim(Members(cnt)) & vbCrLf
    End If
Next cnt
ProcessEnum = dummy
End Function

'Convert a C style structure to a VB style UDT
Private Function ProcessStruct(Cstring As String) As String
Dim Members() As String
Dim MemberNames() As String
Dim MemberTypes() As String
Dim dummy As String
Dim cnt As Integer

If Len(Cstring) = 0 Then Exit Function

'Initialize
Members = GetToken(Cstring, 0, dummy, ";")

ReDim MemberNames(UBound(Members))
ReDim MemberTypes(UBound(Members))

'Loop through the struct members and convert
For cnt = 0 To UBound(Members)
    ProcessArg Members(cnt), MemberNames(cnt), MemberTypes(cnt), True
    If MemberTypes(cnt) = " As " Then MemberNames(cnt) = ""
Next cnt

dummy = ""
'Create the VB declaration
For cnt = 0 To UBound(Members)
    If Len(MemberNames(cnt)) > 0 Then
        dummy = dummy & vbTab & MemberNames(cnt) & MemberTypes(cnt) & vbCrLf
    End If
Next cnt
ProcessStruct = dummy
End Function

'In C you can have unions within structures, eg:
'typedef struct tagHELLO {
'  int MemberA;
'  int MemberB;
'  union {
'    char *MemberC;
'    char *MemberD;
'  } somename;
'  DWORD MemberE;
' } HELLO;
'
'This function gets rid of the union declaration, but not its members
Private Function RemoveUnion(strString As String) As String
Dim pos As Integer
Dim dummy As String

pos = InStr(strString, "union")

If pos = 0 Then
    RemoveUnion = strString
    Exit Function
End If

dummy = Left(strString, pos - 1)
dummy = dummy & Mid(strString, pos + Len("union"))
RemoveUnion = dummy
End Function

'Figures out if we're dealing with an "enum" or a "struct"
Public Function ProcessTypeDef(Cstring As String) As String
Dim parts() As String
Dim dummy As String
Dim dummy2 As String
Dim TypeNames() As String
Dim TypeName As String
Dim cnt As Integer
Dim IsStruct As Boolean
Dim Definition As String

'First of all, get rid of unions
Cstring = RemoveUnion(Cstring)

'Get the individual sections now
parts = GetToken(Cstring, 0, dummy, "{")

'Is it a structure?
IsStruct = (InStr(dummy, "struct") <> 0)

dummy2 = ""

For cnt = 1 To UBound(parts)
    dummy2 = dummy2 & parts(cnt)
Next cnt

'Prepare tokens for processing
parts = GetToken(dummy2, 0, dummy, "}")
TypeName = parts(UBound(parts))
TypeName = RemoveChar(TypeName, " ")
TypeName = RemoveChar(TypeName, ";")
GetToken TypeName, 0, dummy, ","
TypeName = dummy
TypeName = Trim(RemoveChar(TypeName, "*"))

dummy2 = ""
For cnt = 0 To UBound(parts) - 1
    dummy2 = dummy2 & parts(cnt)
Next cnt

'Remove irrelevant characters
dummy2 = RemoveChar(dummy2, vbCr)
dummy2 = RemoveChar(dummy2, vbLf)
dummy2 = RemoveChar(dummy2, "{")
dummy2 = RemoveChar(dummy2, "}")

'Process accordingly
If IsStruct Then
    Definition = ProcessStruct(dummy2)
Else
    Definition = ProcessEnum(dummy2)
End If

'Create the declaration
ProcessTypeDef = IIf(IsPublic, "Public ", "Private ") & _
    IIf(IsStruct, "Type ", "Enum ") & TypeName & vbCrLf & Definition & _
    "End " & IIf(IsStruct, "Type", "Enum")
    

End Function

