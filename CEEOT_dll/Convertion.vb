Option Strict Off
Option Explicit On
Imports System.IO

Friend Class Convertion
	Private mvarvar As Object 'local copy
	Private mvarFormato As String 'local copy
	Private mvarfileName As Object 'local copy
	Private mvarLineNum As Short 'local copy
	Private mvarInicia As Short 'local copy
	Private mvarLeng As Short 'local copy
    Private mvarCondition As String 'local copy
	Private mvarcol1 As Short 'local copy
	Private mvarcol2 As Short 'local copy
	'local variable(s) to hold property value(s)
	Private mvarDescription As Object 'local copy
Public Property Description() As String
  Get
   Description = mvarDescription
  End Get
  Set(ByVal Value As String)
   mvarDescription = Value
  End Set
 End Property
	Public WriteOnly Property col2() As Short
		Set(ByVal Value As Short)
			mvarcol2 = Value
		End Set
	End Property
	Public WriteOnly Property col1() As Short
		Set(ByVal Value As Short)
			mvarcol1 = Value
		End Set
	End Property
Public WriteOnly Property Condition() As String
  Set(ByVal Value As String)
   mvarCondition = Value
  End Set
 End Property
	Public WriteOnly Property Leng() As Short
		Set(ByVal Value As Short)
			mvarLeng = Value
		End Set
	End Property
	Public WriteOnly Property Inicia() As Short
		Set(ByVal Value As Short)
			mvarInicia = Value
		End Set
	End Property
	Public WriteOnly Property LineNum() As Short
		Set(ByVal Value As Short)
			mvarLineNum = Value
		End Set
	End Property
Public WriteOnly Property filename() As String
Set(ByVal Value As String)
  If IsReference(Value) And Not TypeOf Value Is String Then
    mvarfileName = Value
  Else
    mvarfileName = Value
  End If
 End Set
End Property
	
Public Function Convert(ByRef mvarvar As String, ByRef mvarFormato As String) As String
    Dim newval As String
        Dim formated As String
        Dim mvarvarnum As Double
    Dim Leng, lenfor As Integer
    Dim i As Short

    On Error GoTo goError

        Convert = ""
        mvarvarnum = Double.Parse(mvarvar)
        formated = String.Format("{0:" & mvarFormato & "}", mvarvarnum)
        'formated = String.Format(mvarvar, mvarFormato)
    Leng = Len(formated)
    lenfor = Len(mvarFormato)

    For i = Leng + 1 To lenfor
        Convert = Convert & " "
    Next

    If Left(mvarFormato, 1) = "#" Then
        Convert = Convert & formated
    Else
        Convert = formated & Convert
    End If

    Return Convert

    Exit Function
goError:
    MsgBox(Err.Description)

End Function
	
Public Function value() As String
Dim i As Object
Dim mypos As Object
Dim COND1 As Object
Dim z As Object
Dim fs As Object
Dim values() As String

On Error GoTo goError

    fs = CreateObject("Scripting.FileSystemObject")
    z = fs.OpenTextFile(mvarfileName)

    If (Trim(mvarCondition) <> "") Then
        COND1 = z.ReadLine
        mypos = InStr(1, COND1, mvarCondition)
        Do While mypos = 0
            COND1 = z.ReadLine
            mypos = InStr(1, COND1, mvarCondition)
        Loop
        value = Mid(COND1, mvarInicia, 16)
    Else
        For i = 1 To mvarLineNum - 1
            z.ReadLine()
        Next
        value = Mid(z.ReadLine, mvarInicia, mvarLeng)
        If value.Contains("|") Then
            values = Split(value, "|")
            value = values(0)
        End If
    End If

Exit Function
goError:
If Err.Number = 53 Then
MsgBox(Err.Description & mvarfileName)
Else
MsgBox(Err.Description)
End If
End Function
End Class