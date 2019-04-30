Option Strict Off
Option Explicit On
Module SWT_P
	'    PROGRAM APEX_SWAT
	' Name: APEX_SWAT.FOR +
	'      DESCRIPTION: THIS PROGRAM ADD ALL OF THE VALUES IN *.SWT FILES AND
	'    CREATE A NEW ONE WITH TOTALS.
	' Date:             OCTUBER 10, 2003
	' AUTHOR:           OSCAR GALLEGO
	'                 ALI SALEH
	' **************************************************************************
	' VARIABLE DEFINITION SECTION
	' **************************************************************************
	' *** LOCAL VARIABLES
	' filename : Name of *.swt file working
	' title    : Take the 5 lines of title. These are not used
	' Control  : Verify if there are not more files in the list.
	' **************************************************************************
	'UPGRADE_NOTE: Control was upgraded to Control_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Dim currsw() As Object
	Dim filename, Control_Renamed As Object
	Dim filenm As String
	Dim fileN As Object
	Dim lines As Short
	' **************************************************************************
	' OPEN FILES SECTION - OPEN FILE CONTAINING *.SWT FILES NAMES
	' **************************************************************************
	'UPGRADE_NOTE: Main was upgraded to Main_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
	Public Sub Main_Renamed()
		
		'UPGRADE_WARNING: Couldn't resolve default property of object fileN. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		fileN = 0
		ReDim currsw(0)
		' **************************************************************************
		' READ SECTION
		' THIS SECTION TAKE FROM SWTF.TXT THE FILES NAMES.
		' *******************************************************************************
		' READ FIRST FIVE LINES FROM FILE BECAUSE ARE NOT NEEDED.
		
		' READ FILE UNTIL TEN FIRST CHARACTERS ARE BLANKS
		'UPGRADE_WARNING: Couldn't resolve default property of object Initial.Output_files. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
		'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		filename = Dir(Initial.Output_files & "\*.swt")
		'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Do While filename <> ""
			' CALL swtadd SUBROUTINE TO CREATE NEW DATA IN NEW *.SWT FILES
			Call swtadd()
			'UPGRADE_WARNING: Dir has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
			'UPGRADE_WARNING: Couldn't resolve default property of object filename. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			filename = Dir()
			
		Loop 
		'****************************************************************
		' end of program.
		'****************************************************************
	End Sub
	
	Public Sub swtadd()
		Dim Y As Object
		Dim temp As Object
		Dim i As Object
		Dim b As Object
		Dim fs As Object
		' Name: swtadd
		'      DESCRIPTION: THIS PROGRAM VERIFY EACH *.SWT FILE AND DEFINE IN WHICH OTHER
		'    FILE IT HAS TO BE ADD.
		' Date:             OCTUBER 10, 2003
		' AUTHOR:           OSCAR GALLEGO
		'                   ALI SALEH
		' **************************************************************************
		' VARIABLE DEFINITION SECTION
		' **************************************************************************
		' *** LOCAL VARIABLES ***
		' line(6) : Take first line in *.swt file working.
		' line6   : This line contains name of the subarea file. This file helps to
		'           to accummulate the input files into an output file by group.
		' allval  : Used when file is the first in the group. This variable take all
		'           Record in the input file in order to save in the output file.
		' filenum(2): Take group number.
		' Filename, filenam : Take names of the file in order to compare from the current list.
		'           if name exist accumulate, if not the list is updated and new file is created.
		' Filenm : This is the file to be create. It has '.swt' as its extention.
		' Foundit: This variable indicates if current file exist in the list.
		' Lines  : Indicate how many lines each file has.
		' Areades: Description of Area
		' Area   : Area in Ha.
		' Ha     : Ha simble. Part of description.
		' a      : Serve to add seven 0 to the last part of each record.
		' date   : Take day and year only for the first line of data.
		Dim line1(6) As Object
		Dim ha, filenam, line6, allval, areades, a As Object
		Dim filenum(2) As Object
		Dim date1 As String
		Dim foundit As Short
		Dim area As Double
		
		fs = CreateObject("Scripting.FileSystemObject")
        b = fs.OpenTextFile(Initial.Output_files & "\" & filename)
        '      Wait_Form.Label1(2).Text = "Actual File Working  =  " & filename
        'Wait_Form.Show()
        a = " 0"
		For i = 1 To 6
            a = a & " 0"
		Next 
		' **************************************************************************
		' OPEN FILES SECTION - OPEN FILE CONTAINING *.SWT FILES NAMES
		' **************************************************************************
		' ===== FILES USED IN THIS SUBROUTINE ARE OPENED IN APEX_SWAT PROGRAM.======
		'      FILES OPENED ARE IN VARIABLE filename.
		' **************************************************************************
		' READ SECTION
		' THIS SECTION TAKE FROM SWTF.TXT THE FILES NAMES.
		' *******************************************************************************
		' READ FIRST SIX LINES FROM EACH FILE. THESE ARE ONLY TITLES NOT NEEDED
		For i = 1 To 6
            line1(i) = b.ReadLine
		Next 
        line6 = line1(6)
		
        Select Case Initial.Version
            ' FOR VERSION APEX0806
            Case "1.0.0", "1.1.0"
                filenum(1) = Mid(line6, 13, 1)
                filenum(2) = Mid(line6, 14, 1)
                ' FOR VERSION APEX2110, APEX0604
            Case "2.1.0", "2.0.0", "2.3.0", "4.1.0", "4.0.0", "4.2.0", "4.3.0"
                filenum(1) = Mid(line6, 21, 1)
                filenum(2) = Mid(line6, 22, 1)
        End Select
		' READ NEXT TWO LINES AND REPLACE SOME OF THE LAST LINES
        line1(4) = line6
        line1(3) = b.ReadLine
		' LINE 7 IS NOT USED SO IT IS REPLACED BY LINE 8
        temp = b.ReadLine
        areades = Left(temp, 27)
        area = CDbl(Mid(temp, 28, 10))
        ha = Mid(temp, 38, 3)
        line1(6) = b.ReadLine
        filenam = " "
		' ADD A 'P' AT LAST CHARACTER IN THE FILE NAME.
        If (filenum(1) = "0") Then
            filenam = filenum(2) & "P"
        Else
            filenam = filenum(1) & filenum(2) & "P"
        End If
		
		foundit = 0
		' VERIFY IF AT LEAST ONE OF THIS GROUP FILES EXISTS.
		
		For i = 1 To UBound(currsw)
			If (currsw(i - 1) = filenam) Then
				foundit = 1
				Exit For
			Else
				foundit = 0
			End If
		Next 
		' IF THIS IS THE FIRST FILE IN ITS GROUP IT IS ADDED TO CONTROL FILE.
		If (foundit = 0) Then
			currsw(fileN) = filenam
			fileN = fileN + 1
			ReDim Preserve currsw(fileN)
		End If
		
		' DEFINE THE REAL NAME FILE TO WORK
		filenm = Trim(filenam) & ".dat"
		
		If (foundit = 1) Then
			' IF AT LEAST ONE GROUP FILE EXISTS VALUES ARE ADDED IN SUBROUTINE accswt
			Call accmswt(area)
		Else
			' IF THIS IS THE FIRST FILE IN ITS GROUP NEW FILE IS CREATED.
			Y = fs.createtextfile(Initial.Output_files & "\" & filenm)
			
			For i = 1 To 6
				If (i = 5) Then
					Y.WriteLine(areades & area & ha)
				Else
					Y.WriteLine(line1(i))
				End If
			Next 
			
			' UPDATE NEW FILE IN THE GROUP WITH THE FIRST INPUT FILE AND COUNT No. OF LINES
			lines = 0
			Do While b.AtEndOfStream <> True
				' IF IT IS THE FIRST LINE FIELD 1 AND 2 HAVE TO BE BLANKS.
				If (lines = 0) Then
					temp = b.ReadLine
					date1 = Left(temp, 9)
					allval = Mid(temp, 10, 111)
					Y.WriteLine("   0    0" & allval & a)
				Else
					allval = b.ReadLine
					allval = Left(allval, 120) & a
					Y.WriteLine(allval)
				End If
				
				lines = lines + 1
			Loop 
			
			' DEFINE NUMBER OF LINES IN THIS FILE. ALL FILES IN THE SAME GROUP HAVE
			' THE SAME NUMBER OF LINES.
			Y.Close()
		End If
		
		'****************************************************************
		' end of program.
		'****************************************************************
	End Sub
	
    Public Sub accmswt(ByRef area As Object)
        Dim j As Object
        Dim Offset As Object
        Dim X2 As Object
        Dim temp As Object
        Dim z As Object
        Dim Y As Object
        Dim fs As Object
        Dim i As Object
        ' Name: accmswt.FOR
        '      DESCRIPTION: THIS PROGRAM ADD ONE SWT FILE TO ANOTHER IN THE SAME GROUP
        ' Date:             OCTUBER 10, 2003
        ' Author:           OSCAR GALLEGO
        '                 ALI SALEH
        ' **************************************************************************
        ' VARIABLE DEFINITION SECTION
        ' **************************************************************************
        ' Filenm   : Name of *.swt to update.
        ' lin(7)   : Keep the first seven lines in the file. These lines are titles.
        ' Lines    : Number in the lines to add and update in output file.
        ' val2, Val1, Valtot: Contain the values to add and total for new file.
        ' jda, jda1: Day of the year simulated
        ' year,year1: Year simulated.
        ' nothing  : Blank character from file.
        ' Areades: Description of Area
        ' Area1  : Area in Ha.
        ' Area   : Area in Ha receuved from last file readed.
        ' Ha     : Ha simble. Part of description.
        ' **************************************************************************
        ' INPUT VARIABLES
        Dim lin(6) As String
        ' LOCAL VARIABLES
        Dim val2(6) As Object
        Dim Val1(6) As Object
        Dim Valtot(,) As Double
        Dim jda() As Object
        Dim year_Renamed() As Object
        Dim areades As String = String.Empty
        Dim ha As String = String.Empty
        Dim a As String = String.Empty
        Dim area1 As String = String.Empty

        ReDim Valtot(lines, 6)
        ReDim year_Renamed(lines)
        ReDim jda(lines)

        For i = 1 To 7
            a = a & " 0"
        Next
        ' **************************************************************************
        ' OPEN FILES SECTION - OPEN FILE CONTAINING *.SWT FILES NAMES
        ' **************************************************************************
        fs = CreateObject("Scripting.FileSystemObject")
        Y = fs.OpenTextFile(Initial.Output_files & "\" & filenm)
        z = fs.OpenTextFile(Initial.Output_files & "\" & filename)
        ' **************************************************************************
        ' READ SECTION
        ' THIS SECTION READ BOTH *.SWT FILES AND ADD ALL OF THE VALUES AND UPDATE
        ' OLDER *.SWT FILE WITH NEW VALUES.
        ' *******************************************************************************
        ' READ TITLES LINES BEFORE VALUES IN THE CURRENT FILE
        For i = 1 To 6
            If (i = 5) Then
                temp = Y.ReadLine
                areades = Left(temp, 27)
                area1 = Mid(temp, 28, 10)
                ha = Mid(temp, 38, 3)
            Else
                lin(i) = Y.ReadLine
            End If
        Next
        X2 = InStr(1, area1, Space(1))
        area = area + Val(Left(area1, X2 - 1))

        ' READ TITLES LINES BEFORE VALUES IN THE CUMULATIVE FILE
        '   READ (20, 10) lin(6)
        ' READ ALL THE LINES IN THE FILE AND TAKE CURRENT VALUES TO UPDATE WITH NEW VALUES
        For i = 1 To lines
            Offset = 0
            temp = Y.ReadLine
            jda(i) = Left(temp, 4)
            year_Renamed(i) = Mid(temp, 5, 5)
            For j = 1 To 6
                val2(j) = Mid(temp, 12 + Offset, 16)
                Offset = Offset + 17
            Next

            Offset = 0
            temp = z.ReadLine
            For j = 1 To 6
                Val1(j) = Mid(temp, 12 + Offset, 16)
                Offset = Offset + 17
            Next
            ' ACUMMULATE VALUES FOR BOTH FILES.
            For j = 1 To 6
                Valtot(i, j) = val2(j) + Val1(j)
            Next
        Next

        Y.Close()

        Y = fs.OpenTextFile("filenm")
        ' WRITE TITLES IN THE TOTALS FILE
        For i = 1 To 6
            If (i = 5) Then
                Y.WriteLine(areades & area & ha)
            Else
                Y.WriteLine(lin(i))
            End If
        Next
        ' TRANFER VALUES FROM VALTOT MATRIX TO A VECTOR VARIABLE AND THEN
        ' UPDATE FILE WITH TOTALS.
        For i = 1 To lines
            For j = 1 To 6
                'val2 = val2 & Str(Valtot(i, j))
            Next
            'y.WriteLine jda(i) & year(i) & val2 & a
        Next
        '****************************************************************
        ' end of program.
        '****************************************************************
        Y.Close()
    End Sub
End Module