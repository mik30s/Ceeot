Option Strict Off
Option Explicit On
Imports excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Data.Common

Public Class Create_Files
    'version description:
    'X.S.F  X=APEX.S=SWAT.F=FEM
    '1=APEX0806. 4=APEX0604
    '1=SWAT2005, 2=SWAT2009, 3=SWAT2012
    Inherits System.Windows.Forms.Form
    Dim Dataclass As General
    Dim StrTmp As Object
    Dim PFiles(1000) As Object
    Dim subs(10000) As String
    Dim seq, stgadd As Object
    Dim stgnew(10000) As Short
    Dim filet, fileo, name1 As String
    Dim filet1 As String
    Dim SWTNames(1000) As String
    Dim SWTArea As Double
    Dim APEX_bat As String
    Dim Sw_bat As String
    Dim Swat_bat As String

    Private Structure STARTUPINFO
        Dim cb As Integer
        Dim lpReserved As String
        Dim lpDesktop As String
        Dim lpTitle As String
        Dim dwX As Integer
        Dim dwY As Integer
        Dim dwXSize As Integer
        Dim dwYSize As Integer
        Dim dwXCountChars As Integer
        Dim dwYCountChars As Integer
        Dim dwFillAttribute As Integer
        Dim dwFlags As Integer
        Dim wShowWindow As Short
        Dim cbReserved2 As Short
        Dim lpReserved2 As Integer
        Dim hStdInput As Integer
        Dim hStdOutput As Integer
        Dim hStdError As Integer
    End Structure

    Private Structure Sequence 'Create a structure for the sequence in Fig File
        Dim old As Short
        Dim new_Renamed As Short
        Dim route As Short
    End Structure

    Private Structure PROCESS_INFORMATION
        Dim hProcess As Integer
        Dim hThread As Integer
        Dim dwProcessID As Integer
        Dim dwThreadID As Integer
    End Structure

    Private Declare Function FatalAppExit Lib "Kernel32" (ByVal u As Short, ByVal p As String) As String
    Private Declare Function WaitForSingleObject Lib "Kernel32" (ByVal hHandle As Integer, ByVal dwMilliseconds As Integer) As Integer
    Private Declare Function GetLastError Lib "Kernel32" () As Object
    Private Declare Function CreateProcessA Lib "Kernel32" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Integer, ByVal lpThreadAttributes As Integer, ByVal bInheritHandles As Integer, ByVal dwCreationFlags As Integer, ByVal lpEnvironment As Integer, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Integer
    Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject As Integer) As Integer
    Private Declare Function GetExitCodeProcess Lib "Kernel32" (ByVal hProcess As Integer, ByRef lpExitCode As Integer) As Integer
    Private Const NORMAL_PRIORITY_CLASS As Integer = &H20
    Private Const INFINITE As Short = -1
    Dim same As Boolean
    Dim currentDir As String
    Dim convertFormat As Convertion
    Private mvarsuba As String
    Private mvarsoil As String
    Private mvarsite As String
    Private mvarwpm1 As String
    Private mvaropcs As String
    Private mvarparm As String
    Private mvarbsn As String
    Private mvarcod As String
    Private mvarcont As String
    Private mvarReachLenght As Short
    Private mvarsublines As Short
    Private mvarslrgages As String
    Private mvarhmdgatges As String
    Private mvarwndgages As String
    Private mvarpcpgages As Object
    Private mvartmpgages As Object
    Private mvarcntrl As Object
    Private mvarcntrl1 As Object
    Private mvarcntrl3 As Object
    Private mvarincludeNo As String
    Private limit As Short
    Dim totalArea As Double

    Public ReadOnly Property limitx() As Short
        Get
            limitx = limit
        End Get
    End Property
    Public Property includeNo() As Short
        Get
            includeNo = CShort(mvarincludeNo)
        End Get
        Set(ByVal Value As Short)
            mvarincludeNo = CStr(Value)
        End Set
    End Property
    Public Property cntrl() As Object
        Get
            cntrl = mvarcntrl
        End Get
        Set(ByVal Value As Object)
            mvarcntrl = Value
        End Set
    End Property
    Public Property cntrl3() As Object
        Get
            cntrl3 = mvarcntrl3
        End Get
        Set(ByVal Value As Object)
            mvarcntrl3 = Value
        End Set
    End Property
    Public ReadOnly Property cntrl1() As Object
        Get
            cntrl1 = mvarcntrl1
        End Get
    End Property
    Public ReadOnly Property pcpgages() As String
        Get
            pcpgages = mvarpcpgages
        End Get
    End Property
    Public ReadOnly Property tmpgages() As String
        Get
            tmpgages = mvartmpgages
        End Get
    End Property
    Public ReadOnly Property slrgages() As String
        Get
            slrgages = mvarslrgages
        End Get
    End Property
    Public ReadOnly Property hmdgages() As String
        Get
            Dim mvarhmdgages As String = String.Empty
            hmdgages = mvarhmdgages
        End Get
    End Property
    Public ReadOnly Property Wndgages() As String
        Get
            Dim mvarwndgages As String = String.Empty
            Wndgages = mvarwndgages
        End Get
    End Property
    Public ReadOnly Property sublines() As Object
        Get
            sublines = mvarsublines
        End Get
    End Property
    Public ReadOnly Property bsnFile() As String
        Get
            bsnFile = mvarbsn
        End Get
    End Property
    Public ReadOnly Property codFile() As String
        Get
            codFile = mvarcod
        End Get
    End Property
    Public ReadOnly Property ReachLenght() As String
        Get
            ReachLenght = mvarReachLenght
        End Get
    End Property
    Public ReadOnly Property parm() As String
        Get
            parm = mvarparm
        End Get
    End Property
    Public ReadOnly Property suba() As String
        Get
            suba = mvarsuba
        End Get
    End Property
    Public ReadOnly Property Soil() As String
        Get
            Soil = mvarsoil
        End Get
    End Property
    Public ReadOnly Property Site_Renamed() As String
        Get
            Site_Renamed = mvarsite
        End Get
    End Property
    Public ReadOnly Property Opcs() As String
        Get
            Opcs = mvaropcs
        End Get
    End Property
    Public ReadOnly Property wpm1() As String
        Get
            wpm1 = mvarwpm1
        End Get
    End Property
    Public ReadOnly Property cont() As String
        Get
            cont = mvarcont
        End Get
    End Property

    '    Private Sub Create_all_Click()
    '        Dim index As Integer
    '        Dim Pest_File As String
    '        Dim p, g, d, h, e As StreamWriter

    '        On Error GoTo goError
    '        'cwg Me.MousePointer = 11 'or vbHourglass
    '        'Me.ProgressBar1.Visible = True
    '        Me.Refresh()

    '        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
    '        With Wait_Form
    '            .Height = 300
    '            .Label1(0).Text = "The APEX General Files are Being Copied"
    '            .Show()
    '            Call cpyApex()
    '            'Me.ProgressBar1.Value = 10

    '            .Label1(1).Text = "The APEX Control File is Being Created"
    '            .Check1(0).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(0).Visible = True
    '            .Show()
    '            .Refresh()
    '            Control.Apexcont((0))
    '            Pest_File = Dir(Initial.Input_Files & "\Pest.dat")
    '            If Pest_File <> "" And Pest_File <> " " Then Control.Pesticide()
    '            Control.Fertilizer()
    '            'Me.ProgressBar1.Value = 14

    '            .Label1(2).Text = "The APEX Operation Files are Being Created"
    '            .Check1(1).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(1).Visible = True
    '            .Refresh()
    '            Create_Files_Form.Tag = "4"
    '            g = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Opcs))
    '            g.Close()
    '            g.Dispose()
    '            g = Nothing
    '            Sitefiles.FEMFiles(3)
    '            Sitefiles.SiteFiles(3)
    '            'Me.ProgressBar1.Value = 56

    '            .Label1(3).Text = "The APEX Subarea Files are Being Created"
    '            .Check1(2).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(2).Visible = True
    '            .Refresh()
    '            Create_Files_Form.Tag = "2"
    '            d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.suba))
    '            d.Close()
    '            d.Dispose()
    '            d = Nothing
    '            Sitefiles.SiteFiles(2)
    '            'Me.ProgressBar1.Value = 28

    '            .Label1(4).Text = "The APEX Soil Files are Being Created         "
    '            .Check1(3).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(3).Visible = True
    '            .Refresh()
    '            e = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Soil))
    '            e.Close()
    '            e.Dispose()
    '            e = Nothing
    '            Create_Files_Form.Tag = "3"
    '            Sitefiles.SiteFiles(4)
    '            'Me.ProgressBar1.Value = 42

    '            .Label1(5).Text = "The APEX Site Files are Being Created"
    '            .Check1(4).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(4).Visible = True
    '            Wait_Form.Refresh()
    '            h = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Site))
    '            h.Close()
    '            h.Dispose()
    '            h = Nothing
    '            Create_Files_Form.Tag = "5"
    '            Sitefiles.SiteFiles(5)
    '            'Me.ProgressBar1.Value = 70

    '            .Label1(6).Text = "The APEX Weather Files are Being Created"
    '            .Check1(5).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(5).Visible = True
    '            .Refresh()
    '            Create_Files_Form.Tag = "6"
    '            If CDbl(Create_Files_Form.pcpgages) <> 0 Then
    '                If Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
    '                    Dataclass.Weather1()
    '                Else
    '                    Sitefiles.SiteFiles(7)
    '                End If
    '            End If
    '            'Me.ProgressBar1.Value = 86

    '            .Label1(7).Text = "The APEX .wpm Files are Being Created"
    '            .Check1(6).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(6).Visible = True
    '            .Refresh()
    '            p = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.wpm1))
    '            p.Close()
    '            p.Dispose()
    '            p = Nothing
    '            Create_Files_Form.Tag = "7"
    '            Sitefiles.SiteFiles(7)
    '            'Me.ProgressBar1.Value = 98
    '            Wait_Form.Label1(8).Text = "The SWAP General Files are Being Copied"
    '            Wait_Form.Check1(7).CheckState = System.Windows.Forms.CheckState.Checked
    '            Wait_Form.Check1(7).Visible = True
    '            Wait_Form.Refresh()
    '            Call cpyApexSwat()
    '            Wait_Form.Label1(9).Text = "The SWAT General Files are Being Copied"
    '            Wait_Form.Check1(8).CheckState = System.Windows.Forms.CheckState.Checked
    '            Wait_Form.Check1(8).Visible = True
    '            Wait_Form.Refresh()
    '            Call cpySwat()
    '            Wait_Form.Check1(9).CheckState = System.Windows.Forms.CheckState.Checked
    '            Wait_Form.Check1(9).Visible = True
    '            Wait_Form.Refresh()

    '            'Me.ProgressBar1.Visible = False
    '            'cwg Me.MousePointer = 0 'or Default
    '        End With

    '        MsgBox("APEX Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")

    '        Exit Sub
    'goError:
    '        MsgBox(Err.Description)

    '    End Sub

    '    Private Sub Copy_Click(ByRef index As Short)

    '        On Error GoTo goError

    '        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
    '        With Wait_Form
    '            Select Case index
    '                Case 1
    '                    Wait_Form.Label1(0).Text = "The APEX General Files are Being Copied"
    '                    Wait_Form.Show()
    '                    Call cpyApex()
    '                    Wait_Form.Label1(1).Text = "The SWAP General Files are Being Copied"
    '                    Wait_Form.Check1(0).CheckState = System.Windows.Forms.CheckState.Checked
    '                    Wait_Form.Check1(0).Visible = True
    '                    Wait_Form.Refresh()
    '                    Call cpyApexSwat()
    '                    Wait_Form.Label1(2).Text = "The SWAT General Files are Being Copied"
    '                    Wait_Form.Check1(1).CheckState = System.Windows.Forms.CheckState.Checked
    '                    Wait_Form.Check1(1).Visible = True
    '                    Wait_Form.Refresh()
    '                    Call cpySwat()
    '                Case 2
    '                    Wait_Form.Label1(2).Text = "The APEX General Files are Being Copied"
    '                    Wait_Form.Show()
    '                    Wait_Form.Refresh()
    '                    Call cpyApex()
    '                Case 3
    '                    Wait_Form.Label1(2).Text = "The SWAP General Files are Being Copied"
    '                    Wait_Form.Show()
    '                    Wait_Form.Refresh()
    '                    Call cpyApexSwat()
    '                Case 4
    '                    Wait_Form.Label1(2).Text = "The SWAT General Files are Being Copied"
    '                    Wait_Form.Show()
    '                    Wait_Form.Refresh()
    '                    Call cpySwat()
    '            End Select
    '        End With
    '        Wait_Form.Close()
    '        Exit Sub
    'goError:
    '        MsgBox(Err.Description)

    '    End Sub

    Public Sub Apex_Option_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Apex_Option.Click
        Dim index As Short = Apex_Option.GetIndex(eventSender)

        Select Case index
            Case 1
                If Not IsAvailable(Out_Files_Form) Then Out_Files_Form = New Out_Files()
                Out_Files_Form.Show()
        End Select

    End Sub

    'Private Sub APEX_Out_Click(ByRef index As Short)

    '    Select Case index
    '        Case 1
    '            If Not IsAvailable(SWT_Files_Form) Then SWT_Files_Form = New SWT_Files()
    '            SWT_Files_Form.Show()
    '        Case 2
    '    End Select

    'End Sub

    Public Sub Calc_EValue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Calc_EValue.Click
        Dim adoSites As DataTable

        adoSites = New DataTable
        adoSites = getLocalDataTable("SELECT DISTINCT site FROM Measured ORDER BY Site", Initial.Output_files)

        If Not adoSites Is Nothing Then
            If adoSites.Rows.Count = 0 Then
                MsgBox("There is not monthly measured values - Upload monthly muesured values and try again")
            Else
                If Not IsAvailable(frmEvalues_Form) Then frmEvalues_Form = New frmEvalues()
                frmEvalues_Form.ShowDialog()
            End If
        End If
    End Sub

    Public Sub Calc_EValue_Year_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Calc_Evalue_Year.Click
        Dim adoSites As DataTable

        adoSites = New DataTable
        adoSites = getLocalDataTable("SELECT DISTINCT site FROM MeasuredYear ORDER BY Site", Initial.Output_files)

        If Not adoSites Is Nothing Then
            If adoSites.Rows.Count = 0 Then
                MsgBox("There is not annual measured values - Upload annual muesured values and try again")
            Else
                If Not IsAvailable(frmEvaluesYear_Form) Then frmEvaluesYear_Form = New frmEvaluesYear()
                frmEvaluesYear_Form.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Command1_Do_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Call Command1_Click()
    End Sub
    Private Sub Command1_Click()
        If Not IsAvailable(esriMap_Form) Then esriMap_Form = New esriMap()
        esriMap_Form.Show()
    End Sub

    Public Sub compareMeasured_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles compareMeasured.Click
        If Not IsAvailable(MeasuredData_Form) Then MeasuredData_Form = New MeasuredData()
        MeasuredData_Form.Tag = "Month"
        MeasuredData_Form.Show()
    End Sub

    Public Sub compareMeasuredYear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles CompareMeasuredYear.Click
        If Not IsAvailable(MeasuredData_Form) Then MeasuredData_Form = New MeasuredData()
        MeasuredData_Form.Tag = "Year"
        MeasuredData_Form.Show()
    End Sub

    Public Sub Evalue_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Evalue.Click
        If Not IsAvailable(Measured_Predicted_EValues_Form) Then Measured_Predicted_EValues_Form = New Measured_Predicted_EValues()
        Measured_Predicted_EValues_Form.Tag = "Month"
        Measured_Predicted_EValues_Form.ShowDialog()
    End Sub

    Public Sub EvalueYear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Evalue_Year.Click
        If Not IsAvailable(Measured_Predicted_EValues_Form) Then Measured_Predicted_EValues_Form = New Measured_Predicted_EValues()
        Measured_Predicted_EValues_Form.Tag = "Year"
        Measured_Predicted_EValues_Form.ShowDialog()
    End Sub

    Public Sub FarmGeneral_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FarmGeneral.Click
        Dim DataFEMfrm As New Object
        Initial.FEMdbChoice = "Farm General"
        DataFEMfrm.Show(0)
    End Sub

    Public Sub FeedPrice_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FeedPrice.Click
        Dim DataFEMfrm As New Object
        Initial.FEMdbChoice = "feedsAugmented"
        DataFEMfrm.Show()
    End Sub

    Public Sub LivestockGeneralPrices_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles LivestockGeneralPrices.Click
        Dim DataFEMfrm As New Object
        Initial.FEMdbChoice = "livestockgenAugmented"
        DataFEMfrm.Show(0)
    End Sub

    Public Sub loadWatershed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles LoadWatershed.Click
        If Not IsAvailable(SelectFile_Form) Then SelectFile_Form = New SelectFile()
        SelectFile_Form.Combo1.Items.Add("All Picure Files")
        SelectFile_Form.Combo1.Items.Add("Bitmaps (*.bmp; *.dib)")
        SelectFile_Form.Combo1.Items.Add("GIF (*.gif)")
        SelectFile_Form.Combo1.Items.Add("JPEG Images (*.jpg)")
        SelectFile_Form.Combo1.Items.Add("Metafiles (*.wmf; *.emf)")
        SelectFile_Form.Combo1.Items.Add("Icons (*.ico; *.cur)")
        SelectFile_Form.Combo1.SelectedIndex = 0
        SelectFile_Form.Show()
    End Sub

    Public Sub Machines_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Machines.Click
        Dim DataFEMfrm As New Object
        Initial.FEMdbChoice = "machinAugmented"
        DataFEMfrm.Show(0)
    End Sub

    Public Sub MeasuredVal_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MeasuredVal.Click
        If Not IsAvailable(MeasuredValues_Form) Then MeasuredValues_Form = New MeasuredValues()
        MeasuredValues_Form.Tag = "Month"
        MeasuredValues_Form.ShowDialog()
    End Sub

    Public Sub MeasuredValYear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MeasuredValYear.Click
        If Not IsAvailable(MeasuredValues_Form) Then MeasuredValues_Form = New MeasuredValues()
        MeasuredValues_Form.Tag = "Year"
        MeasuredValues_Form.ShowDialog()
    End Sub

    Public Sub Nutrients_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Nutrients.Click
        Dim DataFEMfrm As New Object
        Initial.FEMdbChoice = "nutrientAugmented"

        DataFEMfrm.Show(1)
    End Sub


    Public Sub SWATFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles SWATFiles.Click
        If Not IsAvailable(Edit_Subbasins_Inputs_Form) Then Edit_Subbasins_Inputs_Form = New Edit_Subbasins_Inputs()
        Edit_Subbasins_Inputs_Form.Tag = "Swat"
        Edit_Subbasins_Inputs_Form.Show()

    End Sub

    Public Sub APEXFiles_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles APEXFiles.Click
        If Not IsAvailable(Edit_APEX_inputs1_Form) Then Edit_APEX_inputs1_Form = New Edit_APEX_Inputs1()
        Edit_APEX_inputs1_Form.Show()
    End Sub

    Public Sub Close_Renamed_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Close_Renamed.Click
        Dim index As Short = Close_Renamed.GetIndex(eventSender)
        Me.Close()
    End Sub

    Public Sub Exclude_Consult_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Exclude_Consult.Click
        Call S_Include_Click(1)
        Call UpdateEnvironmentVariables()
    End Sub

    Private Sub cpyP()
        Dim subbasinRead As String
        Dim Subbasin As String
        Dim i As UShort
        Dim Code As String
        Dim none, filet As String
        Dim filetmp As String
        Dim srFile As StreamReader = Nothing

        On Error GoTo goError

        'fs = CreateObject("Scripting.FileSystemObject")

        Select Case Initial.Version
            Case "1.0.0", "2.0.0", "4.0.0"
                name1 = mvarcntrl1
                'a = fs.OpenTextFile(Initial.Output_files & mvarcntrl1)
                srFile = New StreamReader(File.OpenRead(Initial.Output_files & mvarcntrl1))
                For i = 1 To 14
                    srFile.ReadLine()
                Next

                Code = Strings.Left(srFile.ReadLine, 5)
                Do While srFile.EndOfStream <> True
                    none = srFile.ReadLine
                    If srFile.EndOfStream <> True Then
                        Code = Strings.Left(srFile.ReadLine, 5)
                    End If
                Loop
            Case "1.1.0", "2.1.0"
                name1 = Initial.Output_files & Initial.figsFile
                'a = fs.OpenTextFile(Initial.Output_files & "\" & Initial.figsFile)
                srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\" & Initial.figsFile))
                Subbasin = "subbasin"
                subbasinRead = srFile.ReadLine
                Do While Subbasin = Strings.Left(subbasinRead, 8)
                    none = srFile.ReadLine
                    Code = Mid(subbasinRead, 18, 5)
                    subbasinRead = srFile.ReadLine
                Loop

                Call CreateP2003()

                name1 = currentDir & "\1p_2003.p"
                fileo = currentDir & "\1p_2003.p"
        End Select

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing
        filetmp = Initial.Output_files & "\"
        fileo = Initial.Swat_Output & "\*.*"
        filet = Initial.Output_files & "\"
        name1 = fileo
        File.Copy(fileo, filet)
        name1 = Initial.Input_Files & "\crop.dat"
        fileo = Initial.Input_Files & "\crop.dat"
        filet = filet & "crop.dat"
        File.Copy(fileo, filet)

        Exit Sub
goError:
        If Err.Number = 53 Then
            MsgBox(Err.Description & " " & name1, , "Create_File Form (cpyP Function)")
        Else
            MsgBox(Err.Description, , "Create_File Form (cpyP Function)")
        End If

    End Sub
    Private Sub cpyApex()
        Dim ans As Integer
        Dim rs As DataTable
        Dim i As Integer
        Dim z As StreamWriter
        Dim tmp As String

        On Error GoTo goError

        rs = New DataTable
        convertFormat = New Convertion
        rs = getDBDataTable("SELECT * FROM Input_Files WHERE Version=" & "'" & Initial.Version & "'")

        With rs
            For i = 0 To .Rows.Count - 1
                Wait_Form.Pbar_Scenarios.Value = (1 - (.Rows.Count - i) / .Rows.Count) * 100
                Wait_Form.Refresh()
                name1 = Initial.OrgDir & "\" & .Rows(i).Item("File")
                fileo = Initial.OrgDir & "\" & .Rows(i).Item("File")

                If (Strings.Left(.Rows(i).Item("file"), 1) = "*") Then
                    filet = Initial.Output_files & "\"
                Else
                    filet = Initial.Output_files & "\" & .Rows(i).Item("File")
                End If

                If .Rows(i).Item("File") = "APEX2110_2000.EXE" Then
                    name1 = Initial.OrgDir & "\" & .Rows(i).Item("File")
                    fileo = Initial.OrgDir & "\" & .Rows(i).Item("File")
                    filet = Initial.Output_files & "\" & "APEX2110.EXE"
                End If
                File.Copy(fileo, filet, True)
            Next

        End With

        rs = getDBDataTable("SELECT * FROM apexfile_dat WHERE Version=" & "'" & Initial.Version & "'" & " ORDER BY apexfile_dat.Order")
        With rs
            If Initial.Version = "3.0.0" Or Initial.Version = "3.1.0" Then
                z = New StreamWriter(File.Create(Initial.Output_files & "\EPICfile.DAT"))
            Else
                z = New StreamWriter(File.Create(Initial.Output_files & "\Apexfile.DAT"))
            End If
            For i = 0 To .Rows.Count - 1
                z.Write(" ")
                tmp = String.Format("{0, -5}", .Rows(i).Item("FileCode").ToString.ToUpper)
                'b = convertFormat.Convert(.Rows(i).Item("FileCode"), "AAAAA")
                z.Write(tmp)
                z.Write("    ")
                z.WriteLine(.Rows(i).Item("FileName"))
            Next
        End With
        z.Close()

        Exit Sub

goError:
        If Err.Number = 53 Then
            ans = MsgBox(Err.Description & " " & name1, MsgBoxStyle.OkCancel, "Create_File Form (cpyApex Function53)")
        Else
            ans = MsgBox(Err.Description, MsgBoxStyle.OkCancel, "Create_File Form (cpyApex Function)")
        End If

        If ans <> 1 Then Exit Sub

    End Sub

    Public Sub cpyApexSwat()
        Dim ans As Integer
        Dim i As Integer = 0
        Dim SWATF As DataTable

        On Error GoTo goError
        SWATF = New DataTable

        SWATF = getDBDataTable("SELECT * FROM SwatApexF WHERE Version=" & "'" & Initial.Version & "'")

        With SWATF
            For i = 0 To .Rows.Count - 1
                Wait_Form.Pbar_Scenarios.Value = (1 - (.Rows.Count - i) / .Rows.Count) * 100
                name1 = Initial.OrgDir & "\" & .Rows(i).Item("File")
                fileo = Initial.OrgDir & "\" & .Rows(i).Item("File")
                If (Strings.Left(.Rows(i).Item("file"), 1) = "*") Then
                    filet = Initial.Output_files & "\"
                    filet1 = Initial.Swat_Output & "\"
                Else
                    filet = Initial.Output_files & "\" & .Rows(i).Item("File")
                    filet1 = Initial.Swat_Output & "\" & .Rows(i).Item("File")
                End If

                File.Copy(fileo, filet, True)
                File.Copy(fileo, filet1, True)
            Next
        End With

        SWATF.Dispose()
        SWATF = Nothing

        Exit Sub

goError:
        If Err.Number = 53 Then
            ans = MsgBox(Err.Description & " " & name1, MsgBoxStyle.OkCancel, "Create_File Form (cpyApexSwat Function)")
        Else
            ans = MsgBox(Err.Description, MsgBoxStyle.OkCancel, "Create_File Form (cpyApexSwat Function)")
        End If

        If ans <> 1 Then Exit Sub

        SWATF.Dispose()
        SWATF = Nothing

    End Sub
    Public Sub cpySwat()
        '********************************************************************************************************************************
        '*this routine copy all of the files from SWAT input or original SWAT folder to the SWAT_Output folder, the SWAT project files.**
        '********************************************************************************************************************************
        Dim ans As Integer
        Dim rs As DataTable
        Dim i As Integer
        Dim files() As String
        Dim lastIndexOf As Integer

        If Initial.Version = "3.0.0" Or Initial.Version = "3.1.0" Then Exit Sub

        rs = New DataTable

        rs = getDBDataTable("SELECT * FROM Swat_Files WHERE Version=" & "'" & Initial.Version & "'")
        With rs
            On Error GoTo goError
            For i = 0 To .Rows.Count - 1
                Wait_Form.Pbar_Scenarios.Value = (1 - (.Rows.Count - i) / .Rows.Count) * 100
                name1 = Initial.Input_Files & "\" & .Rows(i).Item("File")
                fileo = Initial.Input_Files & "\" & .Rows(i).Item("File")
                If (Strings.Left(.Rows(i).Item("file"), 1) = "*") Then
                    Dataclass.copyAllFiles(Initial.Input_Files, Initial.Swat_Output, "*.*", True)
                Else
                    filet = Initial.Swat_Output & "\" & .Rows(i).Item("File")
                    File.Copy(fileo, filet, True)
                End If

            Next

            Dataclass.copyAllFiles(Initial.New_Swat, Initial.Swat_Output, "*.*", True)
        End With

        changeFileCio()  'change the rch print method to 1 (daily)

        rs.Dispose()
        rs = Nothing

        Exit Sub

goError:
        If Err.Number = 53 Then
            ans = MsgBox(Err.Description & " " & name1, 2, "Create_File Form (cpySwat Function)")
        Else
            ans = MsgBox(Err.Description, MsgBoxStyle.OkCancel, "Create_File Form (cpySwat Function)")
        End If

        Select Case ans
            Case 3
                Exit Sub
            Case 4
                Resume Next
            Case Else
                Resume Next
        End Select
        rs.Dispose()
        rs = Nothing

    End Sub

    '*************************************************************************
    '*this routine change the print option in the file.cio file to 1 (daily **
    '*************************************************************************
    Public Sub changeFileCio()
        Dim srFile As StreamReader
        Dim swfile As StreamWriter
        Dim tempRead As String = String.Empty
        Dim pos As Short = 0

        swfile = New StreamWriter(File.Create(Initial.Swat_Output & "\" & "filecio.bk"))
        srFile = New StreamReader(File.OpenRead(Initial.Swat_Output & "\" & "file.cio"))

        Do While srFile.EndOfStream <> True
            tempRead = srFile.ReadLine
            If tempRead.Contains("IPRINT") Then
                If tempRead.Contains("0") Then
                    pos = tempRead.IndexOf("0")
                    If pos < 0 Then pos = tempRead.IndexOf("2")
                    If pos > 0 Then Mid(tempRead, pos + 1, 1) = "1"
                End If
            End If
            swfile.WriteLine(tempRead)
        Loop

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing
        swfile.Close()
        swfile.Dispose()
        swfile = Nothing

        File.Copy(Initial.Swat_Output & "\" & "file.cio", Initial.Swat_Output & "\org_" & "file.ciobk", True)
        File.Copy(Initial.Swat_Output & "\" & "filecio.bk", Initial.Swat_Output & "\" & "file.cio", True)
    End Sub

    Private Function doAPEXProcess2(ByVal sRunBat As String) As String
        Dim myProcess As System.Diagnostics.Process = New System.Diagnostics.Process
        Dim i As Integer
        Dim sReturn As String = ""

        Try
            ' set the file name and the command line args
            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.Arguments = "/C " & sRunBat & " " & Microsoft.VisualBasic.Chr(34) & " && exit"

            ' start the process in a hidden window
            'myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
            myProcess.StartInfo.CreateNoWindow = True

            ' allow the process to raise events
            myProcess.EnableRaisingEvents = True
            ' add an Exited event handler
            AddHandler myProcess.Exited, AddressOf processAPEXExited

            myProcess.Start()
            For i = 0 To 100000000
                If myProcess.HasExited Then
                    Exit For
                End If
            Next i

            If myProcess.ExitCode = 0 Then
                sReturn = "OK"
            Else
                sReturn = "Calculation Process Failed, Check Your Inputs and Try Again"
            End If

            myProcess.Close()
            myProcess.Dispose()
            Return sReturn

        Catch ex As System.Exception
            Return ex.Message
        End Try
    End Function

    Private Sub processAPEXExited(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bProcessAPEXExited As Boolean = False

        Try
            bProcessAPEXExited = True
        Catch ex As Exception

        End Try
    End Sub

    Public Sub run_process(ByRef cmdline As String)
        Dim myProcess As New Process

        Try
            myProcess.StartInfo.UseShellExecute = False
            '// You can start any process, HelloWorld is a do-nothing example.
            myProcess.StartInfo.FileName = cmdline
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.Start()
            '// This code assumes the process you are starting will terminate itself. 
            '// Given that is is started without a window so you cannot terminate it 
            '// on the desktop, it must terminate itself or you can do it programmatically
            '// from this application using the Kill method.
        Catch
        End Try
    End Sub

    Public Function ExecCmd(ByRef cmdline As String, ByRef direxe As String) As Integer
        Dim ret As Integer
        Dim CREATE_DEFAULT_ERROR_MODE As Integer
        Dim curdrive As String

        Dim proc As PROCESS_INFORMATION
        proc = New PROCESS_INFORMATION
        Dim start As STARTUPINFO
        start = New STARTUPINFO
        Dim exitCode As Integer

        Try
            ExecCmd = 0
            currentDir = CurDir()
            curdrive = Strings.Left(direxe, 2)
            ChDrive(CStr(curdrive))
            ChDir(direxe) 'define the current directory.
            start.cb = Len(start) ' Initialize the STARTUPINFO structure:
            ret = CreateProcessA(cmdline, vbNullString, 0, 0, 1, CREATE_DEFAULT_ERROR_MODE, 0, vbNullString, start, proc)
            ' Wait for the shelled application to finish:
            ret = WaitForSingleObject(proc.hProcess, INFINITE)

            exitCode = GetExitCodeProcess(proc.hProcess, ExecCmd)
            Call CloseHandle(proc.hThread)
            Call CloseHandle(proc.hProcess)

            'curdrive = Strings.Left(currentDir, 2)
            'ChDrive(CStr(curdrive))
            'ChDir(currentDir) 'define the current directory.
            If ExecCmd <> 0 Then
                MsgBox("MS-DOS " & cmdline & " process did not finish properly, please check it out", , "Error Message")
                Return ExecCmd
            End If

        Catch ex As Exception
            MsgBox(Err.Description & " " & cmdline)
        Finally
            curdrive = Strings.Left(currentDir, 2)
            ChDrive(CStr(curdrive))
            ChDir(currentDir) 'define the current directory.
        End Try
    End Function

    Private Sub Exit_Click(ByRef index As Short)
        'cwg Me.HelpContextID = 0
    End Sub

    Public Sub Execute_Epic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Execute_Epic.Click
        Dim index As Short = Execute_Epic.GetIndex(eventSender)
        Dim Sw_bat As String
        Dim myval As Integer
        Dim msg As String = "OK"

        On Error GoTo goError

        myval = 0

        SelectVersion()

        msg = doAPEXProcess2(Initial.Output_files & "\" & APEX_bat)
        If msg <> "OK" Then
            Throw New Global.System.Exception("Error runnig apex - " & msg)
        End If

        'Call validation
        'If same <> True Then

        'If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
        'With Wait_Form
        '    .Label1(1).Text = "The EPIC Program is Being Executed"
        '    .Show()
        '    myval = ExecCmd(Initial.Output_files & "\" & APEX_bat, Initial.Output_files)

        '    If myval <> 0 Then
        '        Exit Sub
        '    Else
        '        MsgBox("EPIC Program Was Successfully Executed", , "Confirmation")
        '    End If

        'End With
        'End If
        Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description, , "Function = Execute_EPIC")

    End Sub
    ' Double click "Execute FEM" menu
    'Updated by Yang on 2/21/2014
    'Public Sub Execute_FEM_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Execute_FEM.Click
    Public Function Execute_FEM_Click() As String
        'Dim index As Short = Execute_FEM.GetIndex(eventSender)
        Dim myval As Single
        Dim currdir As String
        Dim a As Short
        Dim msg As String = "OK"

        Call Control.YearSimulation()
        a = Initial.YearSim
        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
        Wait_Form.Label1(3).Text = "FEM Data is Being Transfer From APEX/SWAT, Please Wait"
        Wait_Form.Show()
        Wait_Form.Refresh()

        If Not IsAvailable(NewScenario1_Form) Then NewScenario1_Form = New NewScenario1()

        Call NewScenario1_Form.FEM_ControlFile_Add(0, 15.24, "N")
        Call ReadLatLong()
        Call takeHRUpcp()

        Wait_Form.Close()
        currdir = CurDir()
        ChDir(Initial.OrgDir & "\fem")

        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
        With Wait_Form
            .Label1(1).Text = "The FEM Program is Being Executed"
            .Show()
            .Refresh()
            'myval = ExecCmd(Initial.OrgDir & "\fem\fem.bat", Initial.OrgDir & "\fem")
            myval = Me.ExecCmd(Initial.OrgDir & "\fem\fem.bat", Initial.OrgDir & "\fem")
            Wait_Form.Close()
        End With
        ChDir(currdir)
        If myval <> 0 Then
            msg = "Error running FEM"
        End If

        Return msg
    End Function

    Public Sub Form_Initials()
        Dim mvarhmdgages As Object
        Dim i, j As Integer
        Dim tmpfiles As String = String.Empty
        Dim pCpfiles As Single
        Dim value As Object
        Dim current_line As Object
        Dim TakeField As Object
        Dim rs As DataTable

        On Error GoTo goError

        rs = New DataTable
        mvarReachLenght = 1 'Reach lenght Calculation
        Initial.subareafile = 2 'Subarea files created with all of the HRU's belonged to the same subarea.

        TakeField = New Convertion
        j = 0
        rs = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'file.cio' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")
        With rs
            If .Rows.Count = 0 Then Exit Sub
            current_line = .Rows(0).Item("Line")
            Do While j < .Rows.Count
                'cwg Done DBNull check added
                If Not IsDBNull(.Rows(j).Item("SwatFile")) AndAlso (.Rows(j).Item("SwatFile") <> "") Then
                    Select Case .Rows(j).Item("SwatFile")
                        Case "file.cio"
                            TakeField.filename = Initial.Input_Files & "\file.cio"
                    End Select

                    TakeField.Leng = .Rows(j).Item("Leng")
                    TakeField.LineNum = .Rows(j).Item("Lines")
                    TakeField.Inicia = .Rows(j).Item("Inicia")
                    value = Trim(TakeField.value())
                Else
                    value = Trim(.Rows(j).Item("Value"))
                End If

                Select Case .Rows(j).Item("Field")
                    Case 1
                        Initial.figsFile = Trim(value)
                    Case 2
                        If IsDBNull(value) OrElse value = "" Then
                            mvarcod = "basins.cod"
                        Else
                            mvarcod = Trim(value)
                        End If
                    Case 3
                        mvarbsn = Trim(value)
                    Case 4
                        pCpfiles = Val(value)
                    Case 5
                        tmpfiles = Val(value)
                    Case 6
                        Create_Files_Form.mvarpcpgages = Val(value)
                        'mvarpcpgages = Val(value)
                    Case 7
                        'mvartmpgages = Val(value)
                        Create_Files_Form.mvartmpgages = Val(value)
                    Case 8
                        limit = pCpfiles
                        If pCpfiles > 6 Then limit = 6
                        For i = 1 To limit
                            Initial.prpfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                        Next
                    Case 9
                        If pCpfiles > 6 Then
                            limit = pCpfiles
                            If pCpfiles > 12 Then limit = 12
                            For i = 7 To limit
                                Initial.prpfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                            Next
                        End If
                    Case 10
                        If pCpfiles > 18 Then
                            MsgBox("There are more than 18 pcp files in the file.cio. Please check it out")
                            Stop
                        Else
                            If pCpfiles > 12 Then limit = pCpfiles
                            For i = 13 To limit
                                Initial.prpfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                            Next
                        End If
                    Case 11
                        limit = tmpfiles
                        If tmpfiles > 6 Then limit = 6
                        For i = 1 To limit
                            Initial.temfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                        Next
                    Case 12
                        If tmpfiles > 6 Then
                            limit = tmpfiles
                            If tmpfiles > 12 Then limit = 12
                            For i = 7 To limit
                                Initial.temfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                            Next
                        End If
                    Case 13
                        If tmpfiles > 18 Then
                            MsgBox("There are more than 18 tmp files in the file.cio. Please check it out")
                            Stop
                        Else
                            If tmpfiles > 12 Then limit = tmpfiles
                            For i = 13 To limit
                                Initial.temfiles(i) = Mid(value, (i - 1) * 13 + 1, 13)
                            Next
                        End If
                    Case 14
                        mvarslrgages = CStr(Val(value))
                    Case 15
                        mvarhmdgages = Val(value)
                    Case 16
                        mvarwndgages = CStr(Val(value))
                    Case 17
                        Initial.slrfiles = value
                    Case 18
                        Initial.hmdfiles = value
                    Case 19
                        Initial.wndfiles = value
                End Select

                j += 1
            Loop
        End With

        Exit Sub
goError:
        MsgBox(Err.Description & " Create_Files (Function: Form_Load)")

    End Sub

    Private Sub Data1_Error(ByRef DataError As Short, ByRef Response As Short)
        Select Case DataError
            'If database file not found.
            Case 3022
                'Display an Open dialog box.
                CommonDialog1Open.ShowDialog()
        End Select
    End Sub

    Public Sub FEMResults_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles FEMResults.Click
        Dim index As Short = FEMResults.GetIndex(eventSender)
        Initial.FEMRes = index
        If Not IsAvailable(ResultsFEMfrm_Form) Then ResultsFEMfrm_Form = New ResultsFEMfrm()
        ResultsFEMfrm_Form.Show()
    End Sub

    Private Sub Create_Files_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dataclass = New General
        'fs = CreateObject("Scripting.FileSystemObject")

        ' cwg change Initial.OrgDir = My.Application.Info.DirectoryPath
        ' should get ceeot-swapp plugin folder
        'Initial.OrgDir = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()(0).FullyQualifiedName)
        Initial.OrgDir = Directory.GetCurrentDirectory
        'Const HelpCNT As Integer = &HB
        'cwg Create_Files_Form.Projects(0).HelpContextID = 21
        'cwg Create_Files_Form.Projects(1).HelpContextID = 22
        'cwg Create_Files_Form.Project.HelpContextID = 20
        'cwg Create_Files_Form.Exclude_Consult.HelpContextID = 30
        'cwg Create_Files_Form.Create_File.HelpContextID = 41
        'cwg System.Drawing.ColorTranslator.FromOle(Create_Files_Form.APEX.HelpContextID = 50
        'cwg Create_Files_Form.Results_.HelpContextID = 60
        'cwg Create_Files_Form.Files(0).HelpContextID = 61
        'cwg Create_Files_Form.Apex_Option(1).HelpContextID = 62
        'cwg Create_Files_Form.FEM_Results.HelpContextID = 63
        'cwg Create_Files_Form.ScenariosCom.HelpContextID = 64
        'cwg Create_Files_Form.APEXFiles.HelpContextID = 71
        'cwg Create_Files_Form.Parameters_.HelpContextID = 72
        'cwg Create_Files_Form.Tools_.HelpContextID = 81
        'cwg Create_Files_Form.Tools(0).HelpContextID = 81
        '< Updated by Yang on 6/23/2014
        'Label2(2).Text = "Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision & "(BETA)"
        'Label2(3).Text = My.Application.Info.Description
        Label2(2).Text = "Version 06.2016"
        Label2(3).Text = "June, 2016"
        '>
        Initial.Errors = 0
        'cwg TODO
        '		With CommonDialog1
        '			' You must set the Help file name.
        '			.HelpFile = "SWAPP.hlp"
        '			' Display the Table of Contents. Note that the
        '			' HelpCNT contstant is not an intrinsic
        '			' constant. The cdlHelpSetContents ensures that
        '			' only the Table of Contents (not Index or Find)
        '			' shows.
        '			'.HelpCommand = HelpCNT Or cdlHelpSetContents
        '			'.ShowHelp
        '		End With
        enable_Menu()
    End Sub

    Private Sub Create_Files_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        'cwg done Infinite recursion! Me.Close()
    End Sub

    Public Sub Help_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs)

    End Sub

    Private Sub Manure_Code_Click(ByRef index As Short)
        Dim n As Object
        Dim p As Object
        Dim SR As Object
        Dim rs As DataTable
        rs = New DataTable
        SR = "UPDATE Apexfiles SET apexfiles.value = " & "'" & index - 1 & "'" & " WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Apexcont.dat' AND Apex_Field = 'MNUL'"
        modifyRecords(SR)

        Select Case index
            Case 1
                SR = "UPDATE Apexfiles SET Apexfiles.value = '0' WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Site' AND Apex_Field = 'UPR'"
                modifyRecords(SR)
                SR = "UPDATE Apexfiles SET Apexfiles.value = '0' WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Site' AND Apex_Field = 'UNR'"
                modifyRecords(SR)
            Case 2, 3, 4
                p = InputBox("Enter Manure Application Rate to Supply P Uptake Rate (kg/ha/yr)", "Create_Files Form - Manure_Code Function")
                n = InputBox("Enter Manure Application Rate to Supply N Uptake Rate (kg/ha/yr)", "Create_Files Form - Manure_Code Function")
                SR = "UPDATE Apexfiles SET Apexfiles.value = " & "'" & p & "'" & " WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Site' AND Apex_Field = 'UPR'"
                modifyRecords(SR)
                SR = "UPDATE Apexfiles SET Apexfiles.value = " & "'" & n & "'" & " WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Site' AND Apex_Field = 'UNR'"
                modifyRecords(SR)
            Case 5
                p = InputBox("Enter Manure Application Rate to Supply P Uptake Rate (kg/ha/yr)", "Create_Files Form - Manure_Code Function")
                n = InputBox("Enter Manure Application Rate to Supply N Uptake Rate (kg/ha/yr)", "Create_Files Form - Manure_Code Function")
                SR = "UPDATE Apexfiles SET Apexfiles.value = " & "'" & p & "'" & " WHERE version = " & "'" & Initial.Version & "'" & " AND Apexfile = 'Site' AND Apex_Field = 'UPR'"
                modifyRecords(SR)
        End Select
    End Sub

    'Public Sub Parameters_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Parameters.Click
    '    Dim index As Short = Parameters.GetIndex(eventSender)
    '    Select Case index
    '        Case 1
    '            If Not IsAvailable(Apex_Files1_Form) Then Apex_Files1_Form = New Apex_Files1()
    '            Apex_Files1_Form.ShowDialog()
    '        Case 2
    '            If Not IsAvailable(Crops1_Form) Then Crops1_Form = New Crops1()
    '            Crops1_Form.ShowDialog()
    '        Case 3
    '            If Not IsAvailable(Till1_Form) Then Till1_Form = New Till1()
    '            Till1_Form.ShowDialog()
    '        Case 4
    '            If Not IsAvailable(Operations_Parameters_Form) Then Operations_Parameters_Form = New Operations_Parameters()
    '            Operations_Parameters_Form.ShowDialog()
    '    End Select
    'End Sub

    Private Sub Reach_Code_Click(ByRef index As Short)
        Select Case index
            Case 0
                mvarReachLenght = 3 ' keep zero
            Case 1
                mvarReachLenght = 0 ' do not apply any calculation
            Case 2
                mvarReachLenght = 1 ' Aply pitagoras
            Case 3
                mvarReachLenght = 2 ' Calculate depending on subarea area in the total area.
        End Select
    End Sub

    Public Sub Projects_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Projects.Click
        Dim index As Short = Projects.GetIndex(eventSender)
        Dim cons As String
        Dim Check_Local As Object
        Dim Database As String
        Dim myfile As String
        'Dim cn1 As ADODB.Connection
        Dim proj1 As DataTable

        On Error GoTo goError

        If Not IsAvailable(Initial1_Form) Then Initial1_Form = New Initial1()
        Select Case index
            Case 0
                Initial1_Form.CreateFolder.Visible = True
                proj1 = New DataTable
                'cwg Initial1_Form.HelpContextID = 21
                Call CloseFrames()
                Call OpenFrames(7)
                Initial1_Form.ShowDialog()
                If Initial.cancel1 = 1 Then Exit Sub
                Call CloseFrames()
                Call OpenFrames(1)
                Initial1_Form.ShowDialog()
                If Initial.cancel1 = 1 Then Exit Sub
                Call CloseFrames()
                Call OpenFrames(2)
                Initial1_Form.ShowDialog()
                If Initial.cancel1 = 1 Then Exit Sub
                Call CloseFrames()
                myfile = Dir(Initial.Input_Files & "\file.cio")
                If myfile <> "file.cio" Then
                    MsgBox("The Path Entered for SWAT Original Files is not Correct", , "Initial - Subdirectories Selection")
                    Exit Sub
                End If
                If Initial.MsgBox_Answer = 6 Then Call Convert00to03()
                Initial.Scenario = "Baseline"
                Initial.Dir2 = Initial.Dir1
                Initial.Scenario = "Baseline"
                Initial.CurrentOption = 10
                Initial.Output_files = Initial.Dir2 & "\APEX"
                Initial.New_Swat = Initial.Dir2 & "\New_SWAT"
                Initial.FEM = Initial.Dir2 & "\FEM"
                Initial.Swat_Output = Initial.Dir2 & "\SWAT_Output"

                If Initial.Errors = 1 Then
                    Initial.Errors = 0
                    Exit Sub
                End If
                myfile = Dir(Initial.OrgDir & "\FEM", FileAttribute.Directory)
                If myfile <> "FEM" Then
                    MkDir((Initial.OrgDir & "\FEM"))
                End If

                Database = Dir(Initial.Dir1 & "\Project_Parameters.mdb")
                If Database = "" Then File.Copy(Initial.OrgDir & "\Project_Parameters_Org.mdb", Initial.Dir1 & "\Project_Parameters.mdb")
                Call enable_Menu()

                Call Initial1_Form.Assign_Version()
                Call Form_Initials()
                Call Initial1_Form.Assign_Version()
                Call SaveEnvironmentVariables()
                Call FEM_ControlFile(0, 49.99, "Manure")
                Check_Local = Dir(Initial.Output_files & "\Local.mdb")
                If Check_Local = "" Then File.Copy(Initial.OrgDir & "\Local.mdb", Initial.Output_files & "\Local.mdb")

            Case 1
                Initial1_Form.CreateFolder.Visible = False
                'cwg Initial1_Form.HelpContextID = 22
                Call OpenFrames(9)
                Initial1_Form.ShowDialog()
                'cwg xx Initial1_Form.Width = 5385
                If Initial.cancel1 = 1 Then Exit Sub
                'If Initial.NumProj = 0 Then
                '    MsgBox("There are not Projects Created. Select Project --> New to Create a New One")
                '    Exit Sub
                'End If
                Call CloseFrames()
                Call enable_Menu()
                Call OpenEnvironmentVariables(Initial.Scenario)
                If Not IsAvailable(Create_Files_Form) Then Create_Files_Form = New Create_Files
                Call Form_Initials()
                Call Initial1_Form.Assign_Version()
                'Call FEM_ControlFile(0, 49.99, "P")
            Case 2
                'cn1 = New ADODB.Connection
                proj1 = New DataTable
                'cn1.Open("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Initial.Dir1 & "\Project_Parameters.mdb;Persist Security Info=False")

                Call OpenFrames(8)
                Call ProjectList()
                'If Initial1_Form.List1.Items.Count = 0 Then
                'MsgBox("There are not Projects Created. Select Project New to Create a New One")
                'Exit Sub
                'End If
                Initial1_Form.Label7(1).Visible = False
                Initial1_Form.List3.Visible = False
                Initial1_Form.ShowDialog()
                Initial1_Form.Label7(1).Visible = True
                Initial1_Form.List3.Visible = True
                If Initial.cancel1 = 1 Then Exit Sub
                Call CloseFrames()
                cons = "DELETE ProjectName FROM Paths WHERE ProjectName = " & "'" & Initial.deleted & "'"
                modifyRecords(cons)
                MsgBox("Project " & Initial.deleted & " Has Been Deleted", MsgBoxStyle.OkOnly, "Delete Projects  ")
        End Select

        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        Exit Sub

goError:
        MsgBox(Err.Description & " --> Subroutine Projects")

    End Sub

    Private Sub SWAT_ADD()
        If Not IsAvailable(FilesToAddFrm_Form) Then FilesToAddFrm_Form = New FilesToAddFrm()
        FilesToAddFrm_Form.Lib_Renamed(0).Path = Initial.Output_files
        FilesToAddFrm_Form.Lib_Renamed(1).Path = Initial.Output_files
        FilesToAddFrm_Form.ShowDialog()
    End Sub

    '*************************************************************************************
    'Name:          FigFile                                                             **
    'Date:          01/02/05                                                            **
    'Description:   Modified the fig file if there is a problem in the sequence about   **
    '               input source files. they have to be before the route command and not**
    '               after as it is shown initialy. If the sequence is correct this sub- **
    '               does nothing.                                                       **
    '*************************************************************************************

    Private Sub FigFile()
        Dim l, i As Integer
        Dim newSeq() As Sequence
        Dim record(0) As String
        Dim temp() As String
        Dim srFile As StreamReader
        Dim swFile As StreamWriter

        On Error GoTo goError

        'fs = CreateObject("Scripting.FileSystemObject")
        name1 = Initial.Output_files & Initial.figsFile
        swFile = New StreamWriter(File.OpenWrite(Initial.Output_files & "\Basinsbk.fig"))
        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Trim(Initial.figsFile)))
        convertFormat = New Convertion

        If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then GoTo Fig2003
        i = 0
        Do While srFile.EndOfStream <> True
            ReDim Preserve record(i)
            record(i) = srFile.ReadLine
            Select Case Mid(record(i), 15, 2)
                Case " 2"
                    i = i + 1
                    ReDim temp(i)
                    ReDim Preserve record(i)
                    temp(i - 1) = record(i - 1)
                    If srFile.EndOfStream <> True Then temp(i) = srFile.ReadLine
                    If (Mid(temp(i), 15, 2) = "10") Then
                        i = i + 2
                        ReDim Preserve temp(i)
                        ReDim Preserve record(i)
                        record(i - 3) = Strings.Left(temp(i - 2), 16) & Mid(temp(i - 3), 17, 6) & Mid(temp(i - 2), 23, 57)
                        temp(i - 1) = srFile.ReadLine
                        temp(i) = srFile.ReadLine
                        record(i - 2) = temp(i - 1)
                        record(i - 1) = Strings.Left(temp(i), 16) & Mid(temp(i - 2), 17, 6) & Mid(temp(i), 23, 6) & Mid(temp(i - 3), 29, 6)
                        record(i) = Strings.Left(temp(i - 3), 16) & Mid(temp(i), 17, 6) & Mid(temp(i - 3), 23, 6) & Mid(record(i - 1), 17, 6) & Mid(temp(i), 35, 17)
                    Else
                        record(i - 1) = temp(i - 1)
                        record(i) = temp(i)
                    End If
            End Select
            i = i + 1
        Loop

        For l = 0 To i - 1
            swFile.WriteLine(record(l))
        Next

        GoTo Endpgm

Fig2003:

        i = 0
        Do While srFile.EndOfStream <> True
            ReDim Preserve record(i)
            record(i) = srFile.ReadLine
            Select Case Mid(record(i), 15, 2)
                Case " 1"
                    i = i + 1
                    ReDim Preserve record(i)
                    record(i) = srFile.ReadLine
                Case " 2"
                    i = i + 2
                    ReDim temp(i)
                    ReDim Preserve record(i)
                    temp(i - 2) = record(i - 2)
                    temp(i - 1) = srFile.ReadLine
                    If srFile.EndOfStream <> True Then temp(i) = srFile.ReadLine
                    If (Mid(temp(i), 15, 2) = "10" Or Mid(temp(i), 15, 2) = " 7") Then
                        i = i + 2
                        ReDim Preserve temp(i)
                        ReDim Preserve record(i)
                        record(i - 4) = Strings.Left(temp(i - 2), 16) & Mid(temp(i - 4), 17, 6) & Mid(temp(i - 2), 23, 57)
                        temp(i - 1) = srFile.ReadLine
                        temp(i) = srFile.ReadLine
                        record(i - 3) = temp(i - 1)
                        record(i - 2) = Strings.Left(temp(i), 16) & Mid(temp(i - 2), 17, 6) & Mid(temp(i), 23, 6) & Mid(temp(i - 4), 29, 6)
                        record(i - 1) = Strings.Left(temp(i - 4), 16) & Mid(temp(i), 17, 6) & Mid(temp(i - 4), 23, 6) & Mid(record(i - 2), 17, 6) & Mid(temp(i - 4), 35, 17)
                        record(i) = temp(i - 3)
                    Else
                        record(i - 1) = temp(i - 1)
                        record(i) = temp(i)
                    End If
            End Select
            i = i + 1
        Loop

        For l = 0 To i - 1
            swFile.WriteLine(record(l))
        Next

Endpgm:
        srFile.Close()
        swFile.Close()
        srFile.Dispose()
        swFile.Dispose()
        srFile = Nothing
        swFile = Nothing

        name1 = Initial.Output_files & Initial.figsFile
        'a = fs.OpenTextFile(Initial.Output_files & "\" & Initial.figsFile)
        fileo = Initial.Output_files & "\" & Initial.figsFile
        filet = Initial.Output_files & "\Basins_bk.fig"
        File.Copy(fileo, filet)
        name1 = Initial.Output_files & "\Basinsbk.fig"
        fileo = Initial.Output_files & "\Basinsbk.fig"
        filet = Initial.Output_files & "\" & Initial.figsFile
        File.Copy(fileo, filet)

        Exit Sub
goError:
        If (Err.Number = 53) Then
            MsgBox(Err.Description & " " & name1, MsgBoxStyle.AbortRetryIgnore, "Create_Files Form (FigFile Function)")
        Else
            MsgBox(Err.Description)
        End If

    End Sub

    Public Sub Yearly_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Yearly.Click
        Dim index As Short = Yearly.GetIndex(eventSender)

        Select Case index
            Case 1
                If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
                Resultsfrm_Form.Tag = "Year"
                Resultsfrm_Form.ShowDialog()
                'If Not IsAvailable(ResultsYear_Form) Then ResultsYear_Form = New ResultsYear()
                'ResultsYear_Form.ShowDialog()
                'Case 2
                '    If Not IsAvailable(ScenarioGraph_Form) Then ScenarioGraph_Form = New ScenarioGraph()
                '    ScenarioGraph_Form.ShowDialog()
        End Select

    End Sub

    Public Sub ResultsM_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ResultsM.Click
        If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
        Resultsfrm_Form.Tag = "Month"
        Resultsfrm_Form.ShowDialog()
    End Sub

    Public Sub ResultsT_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ResultsT.Click
        If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
        Resultsfrm_Form.Tag = "Total"
        Resultsfrm_Form.ShowDialog()
        'If Not IsAvailable(ResultsTotal_Form) Then ResultsTotal_Form = New ResultsTotal()
        'ResultsTotal_Form.ShowDialog()
    End Sub

    Private Sub S_Include_Click(ByRef index As Short)
        Dim gageNumber1 As String
        Dim gageNumber As String
        Dim readtemp As String
        Dim temp2 As String = String.Empty, temp1 As String = String.Empty
        Dim filename As String
        Dim FileNames As String
        Dim Subbasin As String = String.Empty
        Dim i As Short
        Dim rs As DataTable
        Dim srFile As StreamReader
        Dim srFile1 As StreamReader
        Dim subNumber As Short

        On Error GoTo goError

        name1 = Initial.Input_Files & Initial.cntrl2
        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & Initial.cntrl2))
        rs = New DataTable
        Dim totalArea As Single

        same = False
        name1 = "Subbasins"
        modifyRecords("DELETE * FROM subbasins ")
        name1 = "Sub_included"
        name1 = "Subbasins"
        rs = getDBDataTable("SELECT * FROM Subbasins")
        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
        Wait_Form.Label1(3).Text = "Looking for Subbasins used in this project"
        Wait_Form.Show()
        Wait_Form.Refresh()

        For i = 1 To Initial.limit 'Number of lines in control file before subbasin information start.
            temp1 = srFile.ReadLine
            Subbasin = Strings.Left(temp1, 8)
            subNumber = Strings.Mid(temp1, 23, 6)
        Next

        Do While srFile.EndOfStream <> True 'Read the rest of the records in the control file
            If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                If Subbasin <> "subbasin" Then Exit Do 'or until subbasin different from "subbasin" in these versions
            End If

            FileNames = srFile.ReadLine 'take subbasin file name
            filename = Mid(FileNames, Initial.col1, 13) 'take subbasin file name
            If filename = "" Then Exit Do
            temp1 = srFile.ReadLine
            name1 = Initial.Output_files & "\" & filename

            If Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Or Initial.Version = "2.1.0" Or Initial.Version = "2.3.0" Or Initial.Version = "3.1.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Then
                srFile1 = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & filename)) 'open subbasin file
                For i = 1 To 6 'Read subbasin file to take pcp and tmp gages number.
                    readtemp = srFile1.ReadLine
                    If i = 2 Then
                        totalArea = Val(Trim(CStr(Strings.Left(readtemp, 20)))) * 100
                    End If
                Next
                'read pcp gage number
                gageNumber = srFile1.ReadLine
                gageNumber = Val(Trim(gageNumber))
                gageNumber1 = VB6.Format(Str(gageNumber), "0000")
                temp2 = Space(80)
                Mid(temp2, 49, 4) = gageNumber1
                'read tmp gage number
                gageNumber = srFile1.ReadLine
                gageNumber = Val(Trim(gageNumber))
                gageNumber1 = VB6.Format(Str(gageNumber), "0000")
                Mid(temp2, 53, 4) = gageNumber1
                'read slr gage number
                gageNumber = srFile1.ReadLine
                gageNumber = Val(Trim(gageNumber))
                gageNumber1 = VB6.Format(Str(gageNumber), "0000")
                Mid(temp2, 57, 4) = gageNumber1
                'read hmd gage number
                gageNumber = srFile1.ReadLine
                gageNumber = Val(Trim(gageNumber))
                gageNumber1 = VB6.Format(Str(gageNumber), "0000")
                Mid(temp2, 61, 4) = gageNumber1
                'read wnd gage number
                gageNumber = srFile1.ReadLine
                gageNumber = Val(Trim(gageNumber))
                gageNumber1 = VB6.Format(Str(gageNumber), "0000")
                Mid(temp2, 65, 4) = gageNumber1
                srFile1.Close()
                srFile1.Dispose()
                srFile1 = Nothing
            End If

            modifyRecords("INSERT INTO subbasins (Subbasin,pcpNumber,tmpNumber,slrNumber,hmdNumber,wndNumber,rteFile,wgnFile,area,File_Number)" &
                     " VALUES('" & filename & "','" & Mid(temp2, 49, 4) &
                     "','" & Mid(temp2, 53, 4) & "','" & Mid(temp2, 57, 4) & "','" & Mid(temp2, 61, 4) & "','" &
                     Mid(temp2, 65, 4) & "','" & Strings.Left(filename, 9) & ".rte" & "','" & Mid(temp2, 7, 13) & "'," &
                     totalArea & "," & subNumber & ")")
            Subbasin = Strings.Left(temp1, 8)
            subNumber = Strings.Mid(temp1, 23, 6)
        Loop

        rs.Dispose()
        rs = Nothing
        If IsAvailable(Wait_Form) Then Wait_Form.Close()
        If Not IsAvailable(Subbasins_Included_Frm_Form) Then Subbasins_Included_Frm_Form = New Subbasins_Included_Frm()
        Subbasins_Included_Frm_Form.ShowDialog()  'Show the land use window to select them
        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        Exit Sub

goError:
        If Not srFile Is Nothing Then
            srFile.Close()
            srFile.Dispose()
            srFile = Nothing
        End If
        If Err.Number = 53 Then
            MsgBox(Err.Description & " " & name1)
        Else
            MsgBox(Err.Description)
        End If
        If IsAvailable(Wait_Form) Then Wait_Form.Close()

    End Sub

    Private Sub Subarea_Grouping_Click(ByRef index As Short)

        Select Case index
            Case 1
                Initial.subareafile = 1
            Case 2
                Initial.subareafile = 2
        End Select

    End Sub

    Public Sub Scenario_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Scenario1.Click
        Dim index As Short = Scenario1.GetIndex(eventSender)
        If Not IsAvailable(ScenariosList_Form) Then ScenariosList_Form = New ScenariosList()
        Select Case index
            Case 0
                ScenariosList_Form.ShowDialog()
                Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
                Call enable_Menu()
            Case 1
                If Not IsAvailable(NewScenario1_Form) Then NewScenario1_Form = New NewScenario1()
                NewScenario1_Form.ShowDialog()
                ScenariosList_Form.ShowDialog()
                '<06/28/2013
                Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
                '>
                Call enable_Menu()
        End Select
    End Sub

    Public Sub Scenarios_Comparation_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Scenarios_Comparation.Click
        Dim index As Short = Scenarios_Comparation.GetIndex(eventSender)
        Select Case index
            Case 1
                If Not IsAvailable(Scenarios_List_Form) Then Scenarios_List_Form = New Scenarios_List()
                Scenarios_List_Form.ShowDialog()
            Case 2
                If Not IsAvailable(Scenarios_FEM_Form) Then Scenarios_FEM_Form = New Scenarios_FEM()
                Scenarios_FEM_Form.ShowDialog()
            Case 3
                On Error Resume Next
                If Not IsAvailable(Scenarios_Summary_Form) Then Scenarios_Summary_Form = New Scenarios_Summary()
                Scenarios_Summary_Form.ShowDialog()
        End Select

    End Sub

    Public Sub Tools_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Tools.Click
        Dim index As Short = Tools.GetIndex(eventSender)
        Dim RetVal As Object

        Select Case index
            Case 0
                RetVal = Shell("Changes.EXE", 3)
            Case 1
                If Not IsAvailable(SelectFolder_Form) Then SelectFolder_Form = New SelectFolder()
                SelectFolder_Form.ShowDialog()
                Call Convert00to03()
        End Select

    End Sub

    Public Sub UpdateFiles(ByRef Fraction As Double, ByRef HRU As String)
        Dim i As Integer
        Dim basins() As String
        Dim hrus() As String
        Dim srFile, srFile1 As StreamReader
        Dim swFile, swFile1 As StreamWriter

        'fs = CreateObject("Scripting.FileSystemObject")
        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & Me.bsnFile))
        srFile1 = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & HRU))
        convertFormat = New Convertion
        i = 0
        ReDim basins(i)
        ReDim hrus(i)
        basins(i) = srFile.ReadLine
        hrus(i) = srFile1.ReadLine

        Do While srFile.EndOfStream <> True
            i = i + 1
            ReDim Preserve basins(i)
            basins(i) = srFile.ReadLine
        Loop

        i = 0
        Do While srFile1.EndOfStream <> True
            i = i + 1
            ReDim Preserve hrus(i)
            hrus(i) = srFile1.ReadLine
        Loop

        srFile.Close()
        srFile1.Close()
        srFile.Dispose()
        srFile1.Dispose()
        srFile = Nothing
        srFile1 = Nothing
        basins(1) = convertFormat.Convert(System.Math.Round(totalArea, 3), "###########0.000") & Mid(basins(1), 17, 140)
        hrus(1) = convertFormat.Convert(System.Math.Round(Fraction, 7), "#######0.0000000") & Mid(hrus(1), 17, 140)
        swFile = New StreamWriter(File.OpenWrite(Initial.Output_files & Me.bsnFile))
        swFile1 = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & HRU))

        For i = 0 To UBound(basins) - 1
            swFile.WriteLine(basins(i))
        Next

        For i = 0 To UBound(hrus) - 1
            swFile1.WriteLine(hrus(i))
        Next

        swFile.Close()
        swFile1.Close()
        swFile.Dispose()
        swFile1.Dispose()
        swFile = Nothing
        swFile1 = Nothing

    End Sub

    Public Sub CreateP2000()
        Dim i As Integer
        Dim SWATFile As String
        Dim Line As String
        Dim srFile As StreamReader = Nothing
        Dim swFile As StreamWriter = Nothing
        Try
            SWATFile = Dir(Initial.Output_files & "\*.swt")
            'fs = CreateObject("Scripting.FileSystemObject")
            srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\" & SWATFile))
            swFile = New StreamWriter(File.OpenWrite(Initial.Output_files & "\1P_2000.p"))

            For i = 1 To 3
                swFile.WriteLine(srFile.ReadLine)
            Next

            For i = 1 To 3
                srFile.ReadLine()
            Next

            For i = 1 To 3
                swFile.WriteLine(srFile.ReadLine)
            Next
            srFile.ReadLine()

            Line = "   0    0 " & "           0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+0           0 0 0 0 0 0 0 0"
            swFile.WriteLine(Line)
            Do While srFile.EndOfStream <> True
                Line = CStr(Strings.Left(srFile.ReadLine, 10))
                Line = Line & "           0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+0           0 0 0 0 0 0 0 0"
                swFile.WriteLine(Line)
            Loop

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not srFile Is Nothing Then
                srFile.Close()
                srFile.Dispose()
                srFile = Nothing
            End If
            If Not swFile Is Nothing Then
                swFile.Close()
                swFile.Dispose()
                swFile = Nothing
            End If
        End Try

    End Sub

    Public Sub CreateP2003()
        Dim i As Integer
        Dim j As Integer
        Dim SWATFile As String
        Dim Line As String
        Dim srFile As StreamReader = Nothing
        Dim swFile As StreamWriter = Nothing
        Try
            SWATFile = Dir(Initial.Output_files & "\*.swt")

            For j = 1 To 41
                srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\" & SWATFile))
                swFile = New StreamWriter(File.OpenWrite(currentDir & "\" & j & "P.dat"))

                For i = 1 To 3
                    swFile.WriteLine(srFile.ReadLine)
                Next

                For i = 1 To 3
                    srFile.ReadLine()
                Next

                For i = 1 To 3
                    swFile.WriteLine(srFile.ReadLine)
                Next

                Do While srFile.EndOfStream <> True
                    Line = CStr(Strings.Left(srFile.ReadLine, 10))
                    Line = Line & "           0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+00         0.00E+0           0 0 0 0 0 0 0 0 0 0 0 0"
                    swFile.WriteLine(Line)
                Loop
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not srFile Is Nothing Then
                srFile.Close()
                srFile.Dispose()
                srFile = Nothing
            End If
            If Not swFile Is Nothing Then
                swFile.Close()
                swFile.Dispose()
                swFile = Nothing
            End If
        End Try

    End Sub

    '*************************************************************************************
    'Name:          CopyFig                                                             **
    'Date:          04/01/06                                                            **
    'Description:   Modified the fig file adding the new P files (input source files).  **
    '               created in APEX process. These files are added after the routing    **
    '               and a save file is added to use as result instead reach file        **
    '               the routing process in SWAT trap some of the nutrients.             **
    '*************************************************************************************
    Private Sub Copyfig()
        Dim temp As String
        Dim pointPos As Integer
        Dim rtefile As String = String.Empty
        Dim strtmp1 As String
        Dim pfilecurr As String
        Dim cmd2 As Short
        Dim seq As Short
        Dim seq1 As Short
        Dim seq2 As Short
        Dim temp1 As Short
        Dim anterior As Short
        Dim seqSave As Short
        Dim rsR As DataTable
        Dim srFile, srFile1 As StreamReader
        Dim swFile As StreamWriter

        On Error GoTo goError

        rsR = New DataTable

        rsR = getDBDataTable("SELECT * FROM Menus WHERE MenuType='FigFile'")
        anterior = 0
        cmd2 = 1 'Control the sequence of commands in new fig file.
        seq = 0
        seq1 = 0 'recday sequence files.
        System.Array.Clear(PFiles, 0, PFiles.Length) 'array that contains P files to be added in the fig file
        System.Array.Clear(subs, 0, subs.Length) 'array that contains the right number of P files.
        pfilecurr = "    " 'current P file being working
        System.Array.Clear(stgnew, 0, stgnew.Length) 'this array contains the old and new sequence

        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Initial.figsFile))   'open fig file original
        srFile1 = New StreamReader(File.OpenRead(Initial.Swat_Output & "\" & "currswt.txt")) 'contains P-file names
        swFile = New StreamWriter(File.Create(Initial.Swat_Output & "\Fig_New.fig")) 'New fig file.

        Do While srFile1.EndOfStream <> True
            StrTmp = srFile1.ReadLine 'Read file that contains P-file names
            PFiles(Val(StrTmp)) = StrTmp
        Loop

        Do While srFile.EndOfStream <> True
            StrTmp = srFile.ReadLine 'Read fig file
            strtmp1 = StrTmp
            Select Case Val(Mid(StrTmp, 11, 6))
                Case 0 'Finish command
                    swFile.WriteLine(StrTmp)
                Case 1 'Subarea Command
                    Mid(StrTmp, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(StrTmp)
                    Select Case Initial.Version
                        Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                            subs(Val(Mid(StrTmp, 17, 6))) = CStr(Val(Mid(StrTmp, 69, 4)))
                        Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                            strtmp1 = srFile.ReadLine
                            subs(Val(Mid(StrTmp, 17, 6))) = CStr(Val(Mid(strtmp1, 12, 4)))
                            swFile.WriteLine(strtmp1)
                    End Select
                Case 2 'Route command
                    If rsR.Rows(0).Item("OptionName") = "After" Then
                        '/********************** This block write the Point Source File After the Route command.  *******************/
                        pfilecurr = PFiles(CInt(subs(Val(Mid(StrTmp, 23, 6)))))
                        swFile.WriteLine("route     " & "     2" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(stgnew(Val(Mid(StrTmp, 29, 6))))) 'new
                        seqSave = seqSave + 1
                        Select Case Initial.Version
                            Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                            Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                                rtefile = srFile.ReadLine
                                swFile.WriteLine(rtefile)
                                pointPos = InStr(1, rtefile, ".")
                                rtefile = Trim(Mid(rtefile, 1, pointPos) & "Dat")
                        End Select
                        If (pfilecurr <> "") Then
                            seq = seq + 1
                            seq2 = seq1 + seq
                            cmd2 = cmd2 + 1 'new
                            swFile.WriteLine("Recday    " & "    10" & Conv(cmd2) & Conv(seq2) & Space(12) & SWTNames(Val(pfilecurr)))
                            swFile.WriteLine("          " & "APEX" & Trim(pfilecurr) & "P.DAT")
                            cmd2 = cmd2 + 1

                            swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(cmd2 - 2)) 'new
                            swFile.WriteLine("save      " & "     9" & Conv(cmd2) & Conv(seqSave))
                            swFile.WriteLine("          " & rtefile)
                        End If
                        '/***************************************************************************************************************/
                    Else
                        If rsR.Rows(0).Item("OptionName") = "Canada" Then
                            '/********************** This block write the Point Source File Before the Route command Just for Canada *******************/
                            pfilecurr = PFiles(CInt(subs(Val(Mid(StrTmp, 23, 6)))))
                            If (pfilecurr <> "") Then
                                seq = seq + 1
                                seq2 = seq1 + seq
                                seqSave = seqSave + 1
                                swFile.WriteLine("Recday    " & "    10" & Conv(cmd2) & Conv(seq2) & Space(12) & SWTNames(Val(pfilecurr)))
                                swFile.WriteLine("          " & "APEX" & Trim(pfilecurr) & "P.DAT")
                                cmd2 = cmd2 + 1
                                If pfilecurr = "0017" Then
                                    swFile.WriteLine("Recmon    " & "     7" & Conv(cmd2) & Conv(1) & Space(12))
                                    swFile.WriteLine("          " & "Sb17inlet.prn")
                                    cmd2 = cmd2 + 1
                                    swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(cmd2 - 2))
                                    cmd2 = cmd2 + 1
                                End If

                                If pfilecurr = "0016" Then
                                    swFile.WriteLine("Recmon    " & "     7" & Conv(cmd2) & Conv(2) & Space(12))
                                    swFile.WriteLine("          " & "Sbs16WTP.prn")
                                    cmd2 = cmd2 + 1
                                    swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(cmd2 - 2))
                                    cmd2 = cmd2 + 1
                                End If

                                swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                                cmd2 = cmd2 + 1
                                swFile.WriteLine("Route     " & "     2" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(cmd2 - 1))
                            Else
                                swFile.WriteLine("route     " & "     2" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                                seqSave = seqSave + 1
                            End If

                            Select Case Initial.Version
                                Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                                Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                                    rtefile = srFile.ReadLine
                                    swFile.WriteLine(rtefile)
                                    pointPos = InStr(1, rtefile, ".")
                                    rtefile = Trim(Mid(rtefile, 1, pointPos) & "Dat")
                            End Select
                            swFile.WriteLine("save      " & "     9" & Conv(cmd2) & Conv(seqSave))
                            swFile.WriteLine("          " & rtefile)
                        Else
                            '/********************** This block write the Point Source File Before the Route command.*******************/
                            pfilecurr = PFiles(CInt(subs(Val(Mid(StrTmp, 23, 6)))))
                            If (pfilecurr <> "") Then
                                seq = seq + 1
                                seq2 = seq1 + seq
                                seqSave = seqSave + 1
                                swFile.WriteLine("Recday    " & "    10" & Conv(cmd2) & Conv(seq2) & Space(12) & SWTNames(Val(pfilecurr)))
                                swFile.WriteLine("          " & "APEX" & Trim(pfilecurr) & "P.DAT")
                                cmd2 = cmd2 + 1
                                'If pfilecurr = "0017" Then
                                '    swFile.WriteLine("Recmon    " & "     7" & Conv(cmd2) & Conv(1) & Space(12))
                                '    swFile.WriteLine("          " & "Sb17inlet.prn")
                                '    cmd2 = cmd2 + 1
                                '    swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(cmd2 - 2))
                                '    cmd2 = cmd2 + 1
                                'End If

                                'If pfilecurr = "0016" Then
                                '    swFile.WriteLine("Recmon    " & "     7" & Conv(cmd2) & Conv(2) & Space(12))
                                '    swFile.WriteLine("          " & "Sbs16WTP.prn")
                                '    cmd2 = cmd2 + 1
                                '    swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(cmd2 - 2))
                                '    cmd2 = cmd2 + 1
                                'End If

                                swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(cmd2 - 1) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                                cmd2 = cmd2 + 1
                                swFile.WriteLine("Route     " & "     2" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(cmd2 - 1))
                            Else
                                swFile.WriteLine("route     " & "     2" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                                seqSave = seqSave + 1
                            End If

                            Select Case Initial.Version
                                Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                                Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                                    rtefile = srFile.ReadLine
                                    swFile.WriteLine(rtefile)
                                    pointPos = InStr(1, rtefile, ".")
                                    rtefile = Trim(Mid(rtefile, 1, pointPos) & "Dat")
                            End Select
                            'swFile.WriteLine("save      " & "     9" & Conv(cmd2) & Conv(seqSave))
                            'swFile.WriteLine("          " & rtefile)
                        End If
                        '/***************************************************************************************************************/
                    End If
                Case 3 'Reserviours command
                    swFile.WriteLine("Routeres  " & "     3" & Conv(cmd2) & Mid(StrTmp, 23, 6) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                    swFile.WriteLine(srFile.ReadLine)
                Case 5 'Add command
                    temp = Mid(StrTmp, 23, 6)
                    temp1 = Val(temp)
                    temp1 = stgnew(temp1)
                    temp = Conv(temp1)
                    swFile.WriteLine("add       " & "     5" & Conv(cmd2) & Conv(stgnew(Val(Mid(StrTmp, 23, 6)))) & Conv(stgnew(Val(Mid(StrTmp, 29, 6)))))
                Case 6 'Rec Hour command
                    swFile.WriteLine("Rechour   " & "     6" & Conv(cmd2) & Mid(StrTmp, 23, 6))
                    swFile.WriteLine(srFile.ReadLine)
                Case 7 'Rec Monthly command
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(strtmp1)
                    swFile.WriteLine(srFile.ReadLine)
                Case 8 'Rec Yearly command
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(strtmp1)
                    swFile.WriteLine(srFile.ReadLine)
                Case 9 'Save command
                    seqSave = seqSave + 1
                    Mid(strtmp1, 23, 6) = Conv(seqSave)
                    swFile.WriteLine(strtmp1)
                    Select Case Initial.Version
                        Case "1.0.0", "2.0.0", "3.0.0", "4.0.0"
                        Case "1.1.0", "1.2.0", "1.3.0", "2.1.0", "2.3.0", "3.1.0", "4.1.0", "4.2.0", "4.3.0"
                            swFile.WriteLine(srFile.ReadLine)
                    End Select
                Case 10 'Rec Daily command
                    seq1 = Val(Mid(StrTmp, 23, 6))
                    seq2 = seq1 + seq
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    Mid(strtmp1, 23, 6) = Conv(seq2)
                    swFile.WriteLine(strtmp1)
                    swFile.WriteLine(srFile.ReadLine)
                Case 11 'Rec reccnst command
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(strtmp1)
                    swFile.WriteLine(srFile.ReadLine)
                Case 12 'STRUCTURE command
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    Mid(strtmp1, 23, 6) = Conv(stgnew(Val(Mid(StrTmp, 23, 6))))
                    swFile.WriteLine(strtmp1)
                Case 14 'SAVECONC command
                    'Mid(StrTmp1, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(strtmp1)
                Case 16 'autocall command
                    Mid(strtmp1, 17, 6) = Conv(cmd2)
                    swFile.WriteLine(strtmp1)
            End Select

            stgnew(Val(Mid(StrTmp, 17, 6))) = cmd2
            cmd2 = cmd2 + 1
        Loop

        rsR.Dispose()
        rsR = Nothing

        swFile.Close()
        swFile.Dispose()
        swFile = Nothing

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

        srFile1.Close()
        srFile1.Dispose()
        srFile1 = Nothing

        File.Copy(Initial.Swat_Output & "\" & Initial.figsFile, Initial.Swat_Output & "\" & "bk" & Initial.figsFile, True)
        File.Copy(Initial.Swat_Output & "\Fig_New.fig", Initial.Swat_Output & "\" & Initial.figsFile, True)

        Exit Sub

goError:
        MsgBox(Err.Description)
    End Sub

    Function Conv(ByRef cmd As Short) As String
        Dim i As Short
        Dim lenx As Short
        Dim cmdx As String

        cmdx = Trim(Str(cmd))
        lenx = Len(cmdx)
        Conv = ""
        For i = 1 To 6 - lenx
            Conv = Conv & " "
        Next
        Conv = Conv & cmdx

    End Function

    Public Sub Updates_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Updates.Click
        'frmEvalues_Form.updateNTTSoil()
    End Sub

    Public Sub UploadMeasured_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UploadMeasured.Click
        Dim ADORecordset As DataTable
        Dim sites As String
        Dim j As Integer
        Dim i As Integer
        Dim appExcel As excel.Application
        Dim wBookExcel As excel.Workbook
        Dim wSheetExcel As excel.Worksheet
        Dim sitesToUpload(0) As Short
        Dim currentSite As Short
        Dim flow, sed, orgN, orgP, NO3, minP, totalN, totalP As Single
        Dim site, year, mon As Short
        Dim myConnection As OleDb.OleDbConnection
        Dim dbConnectString As String
        Dim command As OleDb.OleDbCommand
        Dim query As String

        'define connection parms
        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        myConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = dbConnectString
        myConnection.Open()

        If Not IsAvailable(SelectFile_Form) Then SelectFile_Form = New SelectFile()
        SelectFile_Form.Combo1.Items.Clear()
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.SelectedIndex = 6
        SelectFile_Form.ShowDialog()

        'appExcel = New excel.ApplicationClass 'Changed due to the manually switch the target framework to .NET 4.0
        appExcel = New excel.Application
        If Initial.measuredfile Is Nothing Then Exit Sub
        wBookExcel = appExcel.Workbooks.Open(Initial.measuredfile)
        wSheetExcel = wBookExcel.Worksheets(1)
        Me.Cursor = Cursors.WaitCursor
        With wSheetExcel
            i = 1
            j = 0
            currentSite = 0
            If Not IsNumeric(.Cells(i, 1).value) Then
                i = i + 1
            End If
            Do While .Cells(i, 1).value <> 0
                i = i + 1
                If .Cells(i, 1).value <> currentSite And Not IsNothing(.Cells(i, 1).value) Then
                    ReDim Preserve sitesToUpload(j)
                    sitesToUpload(j) = .Cells(i, 1).value
                    j = j + 1
                    currentSite = .Cells(i, 1).value
                End If
            Loop
        End With

        sites = sitesToUpload(0)
        For i = 1 To UBound(sitesToUpload)
            sites = sites & " OR Site=" & sitesToUpload(i)
        Next

        ADORecordset = New DataTable
        modifyLocalRecords("DELETE * FROM Measured WHERE Site=" & sites, Initial.Output_files)
        With wSheetExcel
            i = 2
            '   Do While .Cells(i, 1).value <> 0 And Not IsNothing(.Cells(i, 1))
            Do While .Cells(i, 1).value <> 0
                flow = 0 : sed = 0 : NO3 = 0 : minP = 0 : orgN = 0 : orgP = 0 : totalN = 0 : totalP = 0
                site = .Cells(i, 1).value
                year = .Cells(i, 2).value
                mon = .Cells(i, 3).value
                If IsNumeric(.Cells(i, 4).value) Then flow = .Cells(i, 4).value
                If IsNumeric(.Cells(i, 5).value) Then sed = .Cells(i, 5).value
                If IsNumeric(.Cells(i, 6).value) Then NO3 = .Cells(i, 6).value
                If IsNumeric(.Cells(i, 7).value) Then minP = .Cells(i, 7).value
                If IsNumeric(.Cells(i, 8).value) Then orgN = .Cells(i, 8).value
                If IsNumeric(.Cells(i, 9).value) Then orgP = .Cells(i, 9).value
                If IsNumeric(.Cells(i, 10).value) Then totalN = .Cells(i, 10).value
                If IsNumeric(.Cells(i, 11).value) Then totalP = .Cells(i, 11).value
                query = "INSERT INTO Measured (site,[year],[mon],flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " &
                      "VALUES(" & site & "," & year & "," & mon & "," & flow & "," & sed & "," & NO3 &
                      "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP & ")"
                'modifyLocalRecords("INSERT INTO Measured (site,year,mon,flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " & _
                '                  "VALUES(" & site & "," & year & "," & mon & "," & flow & "," & sed & "," & NO3 & _
                '                  "," & NO3 & "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP, Initial.Output_files)
                command = New OleDb.OleDbCommand(query, myConnection)
                command.ExecuteNonQuery()
                i = i + 1
            Loop
        End With
        Me.Cursor = Cursors.Default
        MsgBox("Measured values were uploaded", MsgBoxStyle.OkOnly)

    End Sub

    Public Sub UploadMeasuredYear_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UploadMeasuredYear.Click
        Dim ADORecordset As DataTable
        Dim sites As String
        Dim j As Integer
        Dim i As Integer
        Dim appExcel As excel.Application
        Dim wBookExcel As excel.Workbook
        Dim wSheetExcel As excel.Worksheet
        Dim sitesToUpload() As Short
        Dim currentSite As Short
        Dim flow, sed, orgN, orgP, NO3, minP, totalN, totalP As Single
        Dim site, year As Short
        Dim myConnection As OleDb.OleDbConnection
        Dim dbConnectString As String
        Dim command As OleDb.OleDbCommand
        Dim query As String

        ReDim sitesToUpload(0)
        'define connection parms
        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        myConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = dbConnectString
        myConnection.Open()

        If Not IsAvailable(SelectFile_Form) Then SelectFile_Form = New SelectFile()
        SelectFile_Form.Combo1.Items.Clear()
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.SelectedIndex = 6
        SelectFile_Form.ShowDialog()

        appExcel = New excel.Application
        If Initial.measuredfile Is Nothing Then Exit Sub
        wBookExcel = appExcel.Workbooks.Open(Initial.measuredfile)
        wSheetExcel = wBookExcel.Worksheets(1)
        Me.Cursor = Cursors.WaitCursor
        With wSheetExcel
            i = 1
            j = 0
            currentSite = 0
            If Not IsNumeric(.Cells(i, 1).value) Then
                i = i + 1
            End If
            Do While .Cells(i, 1).value <> 0
                i = i + 1
                If .Cells(i, 1).value <> currentSite And Not IsNothing(.Cells(i, 1).value) Then
                    ReDim Preserve sitesToUpload(j)
                    sitesToUpload(j) = .Cells(i, 1).value
                    j = j + 1
                    currentSite = .Cells(i, 1).value
                End If
            Loop
        End With

        sites = sitesToUpload(0)
        For i = 1 To UBound(sitesToUpload)
            sites = sites & " OR Site=" & sitesToUpload(i)
        Next

        ADORecordset = New DataTable
        modifyLocalRecords("DELETE * FROM MeasuredYear WHERE Site=" & sites, Initial.Output_files)
        With wSheetExcel
            i = 2
            Do While .Cells(i, 1).value <> 0
                flow = 0 : sed = 0 : NO3 = 0 : minP = 0 : orgN = 0 : orgP = 0 : totalN = 0 : totalP = 0
                site = .Cells(i, 1).value
                year = .Cells(i, 2).value
                'mon = .Cells(i, 3).value
                If IsNumeric(.Cells(i, 3).value) Then flow = .Cells(i, 3).value
                If IsNumeric(.Cells(i, 4).value) Then sed = .Cells(i, 4).value
                If IsNumeric(.Cells(i, 5).value) Then NO3 = .Cells(i, 5).value
                If IsNumeric(.Cells(i, 6).value) Then minP = .Cells(i, 6).value
                If IsNumeric(.Cells(i, 7).value) Then orgN = .Cells(i, 7).value
                If IsNumeric(.Cells(i, 8).value) Then orgP = .Cells(i, 8).value
                If IsNumeric(.Cells(i, 9).value) Then totalN = .Cells(i, 9).value
                If IsNumeric(.Cells(i, 10).value) Then totalP = .Cells(i, 10).value
                query = "INSERT INTO MeasuredYear (site,[year],flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " &
                      "VALUES(" & site & "," & year & "," & flow & "," & sed & "," & NO3 &
                      "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP & ")"
                command = New OleDb.OleDbCommand(query, myConnection)
                command.ExecuteNonQuery()
                i = i + 1
            Loop
        End With
        Me.Cursor = Cursors.Default
        MsgBox("Measured values were uploaded", MsgBoxStyle.OkOnly)

    End Sub

    Public Sub UploadMeasured1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles UploadMeasured.Click
        Dim ADORecordset As DataTable
        Dim sites As String
        Dim j As Integer
        Dim i As Integer
        Dim appExcel As excel.Application
        Dim wBookExcel As excel.Workbook
        Dim wSheetExcel As excel.Worksheet
        Dim sitesToUpload(0) As Short
        Dim currentSite As Short
        Dim flow, sed, orgN, orgP, NO3, minP, totalN, totalP As Single
        Dim site, year, mon As Short
        Dim myConnection As OleDb.OleDbConnection
        Dim dbConnectString As String
        Dim command As OleDb.OleDbCommand
        Dim query As String

        'define connection parms
        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        myConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = dbConnectString
        myConnection.Open()
        If Not IsAvailable(SelectFile_Form) Then SelectFile_Form = New SelectFile()
        SelectFile_Form.Combo1.Items.Clear()
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.SelectedIndex = 6
        SelectFile_Form.ShowDialog()

        appExcel = New excel.Application
        If Initial.measuredfile Is Nothing Then Exit Sub
        wBookExcel = appExcel.Workbooks.Open(Initial.measuredfile)
        wSheetExcel = wBookExcel.Worksheets(1)
        Me.Cursor = Cursors.WaitCursor
        With wSheetExcel
            i = 1
            j = 0
            currentSite = 0
            If Not IsNumeric(.Cells(i, 1).value) Then
                i = i + 1
            End If
            Do While .Cells(i, 1).value <> 0
                i = i + 1
                If .Cells(i, 1).value <> currentSite And Not IsNothing(.Cells(i, 1).value) Then
                    ReDim Preserve sitesToUpload(j)
                    sitesToUpload(j) = .Cells(i, 1).value
                    j = j + 1
                    currentSite = .Cells(i, 1).value
                End If
            Loop
        End With

        sites = sitesToUpload(0)
        For i = 1 To UBound(sitesToUpload)
            sites = sites & " OR Site=" & sitesToUpload(i)
        Next

        ADORecordset = New DataTable
        modifyLocalRecords("DELETE * FROM Measured WHERE Site=" & sites, Initial.Output_files)
        With wSheetExcel
            i = 2
            '   Do While .Cells(i, 1).value <> 0 And Not IsNothing(.Cells(i, 1))
            Do While .Cells(i, 1).value <> 0
                flow = 0 : sed = 0 : NO3 = 0 : minP = 0 : orgN = 0 : orgP = 0 : totalN = 0 : totalP = 0
                site = .Cells(i, 1).value
                year = .Cells(i, 2).value
                mon = .Cells(i, 3).value
                If IsNumeric(.Cells(i, 4).value) Then flow = .Cells(i, 4).value
                If IsNumeric(.Cells(i, 5).value) Then sed = .Cells(i, 5).value
                If IsNumeric(.Cells(i, 6).value) Then NO3 = .Cells(i, 6).value
                If IsNumeric(.Cells(i, 7).value) Then minP = .Cells(i, 7).value
                If IsNumeric(.Cells(i, 8).value) Then orgN = .Cells(i, 8).value
                If IsNumeric(.Cells(i, 9).value) Then orgP = .Cells(i, 9).value
                If IsNumeric(.Cells(i, 10).value) Then totalN = .Cells(i, 10).value
                If IsNumeric(.Cells(i, 11).value) Then totalP = .Cells(i, 11).value
                query = "INSERT INTO Measured (site,[year],[mon],flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " &
                      "VALUES(" & site & "," & year & "," & mon & "," & flow & "," & sed & "," & NO3 &
                      "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP & ")"
                'modifyLocalRecords("INSERT INTO Measured (site,year,mon,flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " & _
                '                  "VALUES(" & site & "," & year & "," & mon & "," & flow & "," & sed & "," & NO3 & _
                '                  "," & NO3 & "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP, Initial.Output_files)
                command = New OleDb.OleDbCommand(query, myConnection)
                command.ExecuteNonQuery()
                i = i + 1
            Loop
        End With
        Me.Cursor = Cursors.Default
        MsgBox("Measured values were uploaded", MsgBoxStyle.OkOnly)

    End Sub

    Public Sub OpenFrames(ByRef OpenFrame As Short)

        Initial1_Form.Command1.Visible = True

        Select Case OpenFrame
            Case 1
                Initial1_Form.Frame1.Visible = True
                Initial1_Form.Text = "Versions Configuration"
                Initial1_Form.Frame1.Visible = True
                Initial1_Form.Width = 360 'cwg change from 5385
            Case 2
                Initial1_Form.Frame2.Visible = True
                Initial1_Form.Text = "SWAT Files Location"
                Initial1_Form.Width = 360 'cwg change from 5385
            Case 7
                Initial1_Form.Frame7.Visible = True
                Initial1_Form.Text = "Project Files Location"
                Initial1_Form.Width = 360 'cwg change from 5385
            Case 8
                Initial1_Form.Frame8.Visible = True
                Initial1_Form.Text = "Project/Scenario Selection"
            Case 9
                Initial1_Form.Frame7.Visible = True
                Initial1_Form.Frame8.Visible = True
                Initial1_Form.Text = "Project/Scenario Selection"
                Initial1_Form.Width = 550 'cwg change from 8130
                Initial1_Form.Command1.Visible = False
        End Select
    End Sub

    Private Sub CloseFrames()
        Initial1_Form.Frame1.Visible = False
        Initial1_Form.Frame2.Visible = False
        Initial1_Form.Frame7.Visible = False
        Initial1_Form.Frame8.Visible = False
    End Sub

    Public Sub SaveEnvironmentVariables()
        Dim SWATOutput_Dir, FEM_Dir, NewSWAT_Dir, APEX_Dir, DirExist As String
        Dim cons As String

        On Error GoTo Open_Error

        On Error GoTo Dir_Error
        DirExist = Dir(Initial.Dir2 & "\APEX", FileAttribute.Directory)
        If DirExist = "" Then MkDir(Initial.Dir2 & "\APEX")
        DirExist = Dir(Initial.Dir2 & "\New_SWAT", FileAttribute.Directory)
        If DirExist = "" Then MkDir(Initial.Dir2 & "\New_SWAT")
        DirExist = Dir(Initial.Dir2 & "\FEM", FileAttribute.Directory)
        If DirExist = "" Then MkDir(Initial.Dir2 & "\FEM")
        DirExist = Dir(Initial.Dir2 & "\SWAT_Output", FileAttribute.Directory)
        If DirExist = "" Then MkDir(Initial.Dir2 & "\SWAT_Output")
        DirExist = Dir(Initial.Dir2 & "\Scenarios", FileAttribute.Directory)
        If DirExist = "" Then MkDir(Initial.Dir2 & "\Scenarios")

        On Error GoTo Project_Error
        If Initial.Scenario <> "Baseline" Then
            APEX_Dir = "\Scenarios\" & Initial.Scenario & "\APEX"
            NewSWAT_Dir = "\Scenarios\" & Initial.Scenario & "\New_SWAT"
            FEM_Dir = "\Scenarios\" & Initial.Scenario & "\FEM"
            SWATOutput_Dir = "\Scenarios\" & Initial.Scenario & "\SWAT_Output"
        Else
            APEX_Dir = "\APEX"
            NewSWAT_Dir = "\New_SWAT"
            FEM_Dir = "\FEM"
            SWATOutput_Dir = "\SWAT_Output"
        End If
        Dim proj1 As DataTable

        proj1 = New DataTable

        cons = "SELECT * FROM Paths"
        proj1 = getDBDataTable(cons)
        If Not proj1.Columns.Contains("LastRun") Then
            getDBDataSet("ALTER TABLE Paths ADD COLUMN LastRun Text(10);")
            modifyRecords("UPDATE paths SET LastRun = 'APEX'")
        End If
        If Not proj1.Columns.Contains("Type1") Then
            getDBDataSet("ALTER TABLE Paths ADD COLUMN Type1 Text(255);")
        End If

        modifyRecords("INSERT INTO Paths (ProjectName,folder,Version,Version1,Scenario,SWAT_Input,APEX,New_SWAT," &
                      "FEM,SWAT_Output,CurrentOption,FEMChecked,type1) VALUES('" & Initial.Project & "','" & Initial.Dir1 & "','" &
                      Initial.Version & "','" & Initial.Version2 & "','" & Initial.Scenario & "','" & Initial.Input_Files &
                      "','" & APEX_Dir & "','" & NewSWAT_Dir & "','" & FEM_Dir & "','" & SWATOutput_Dir & "'," &
                      Initial.CurrentOption & "," & True & ",'" & Initial.Scenario_type & "," & Variables.qnSF & "," &
                      Variables.qpSF & "," & Variables.ynSF & "," & Variables.ypSF & "')")
        Exit Sub

Copy_Error:
        MsgBox("Database Project_Parameters does not Exist in CEEOT-SWAPP Folder")
        Initial.Errors = 1
        Exit Sub

Open_Error:
        MsgBox("Database Project_Parameters can not Be Opened, Please Check This Database in Your Project Folder")
        Initial.Errors = 1
        Exit Sub

Dir_Error:
        MsgBox("A Folder can not Be Created. It can Be a Security Issue")
        Initial.Errors = 1
        Exit Sub

Project_Error:
        MsgBox("The Project was already Created. Please Select another Folder or Open an Existed One")
        Initial.Errors = 1
        Exit Sub
    End Sub

    Public Sub OpenEnvironmentVariables(ByRef scenarioLocal As String)
        Dim Database As String
        Dim cons As String
        Dim proj1 As DataTable

        proj1 = New DataTable

        cons = "SELECT * FROM Subbasins"

        proj1 = getDBDataTable(cons)
        If Not proj1.Columns.Contains("File_Number") Then
            getDBDataSet("ALTER TABLE Subbasins ADD COLUMN File_Number Text(10);")
            modifyRecords("UPDATE subbasins SET File_Number = left(rteFile,5)")
        End If

        cons = "SELECT * FROM Paths WHERE ProjectName=" & "'" & Initial.Project & "'" & " AND Scenario = " & "'" & scenarioLocal & "'"
        proj1 = getDBDataTable(cons)
        If Not proj1.Columns.Contains("LastRun") Then
            getDBDataSet("ALTER TABLE Paths ADD COLUMN LastRun Text(10);")
            modifyRecords("UPDATE paths SET LastRun = 'APEX'")
        End If
        If Not proj1.Columns.Contains("Field_Type") Then
            getDBDataSet("ALTER TABLE Paths ADD COLUMN Field_Type Text(4);")
            cons = "UPDATE Paths SET Type1='Text' WHERE ProjectName=" & "'" & Initial.Project & "'"
            modifyRecords(cons)
        End If
        If Not proj1.Columns.Contains("Type1") Then
            getDBDataSet("ALTER TABLE Paths ADD COLUMN Type1 Text(50);")
        End If
        With proj1
            Initial.Project = .Rows(0).Item("ProjectName")
            Initial.Dir1 = .Rows(0).Item("Folder")
            Initial.Version = .Rows(0).Item("Version")
            Initial.CurrentOption = .Rows(0).Item("CurrentOption")
            Initial.FEMChecked = .Rows(0).Item("FEMChecked")

            Select Case Initial.Version
                Case "1.0.0"
                    Call Initial1_Form.Option1_Click(1)
                    Call Initial1_Form.Option2_Click(0)
                Case "1.1.0"
                    Call Initial1_Form.Option1_Click(1)
                    Call Initial1_Form.Option2_Click(1)
                Case "1.2.0"
                    Call Initial1_Form.Option1_Click(1)
                    Call Initial1_Form.Option2_Click(2)
                Case "1.3.0"
                    Call Initial1_Form.Option1_Click(1)
                    Call Initial1_Form.Option2_Click(3)
                Case "2.0.0"
                    Call Initial1_Form.Option1_Click(0)
                    Call Initial1_Form.Option2_Click(0)
                Case "2.1.0"
                    Call Initial1_Form.Option1_Click(0)
                    Call Initial1_Form.Option2_Click(1)
                Case "2.2.0"
                    Call Initial1_Form.Option1_Click(0)
                    Call Initial1_Form.Option2_Click(2)
                Case "2.3.0"
                    Call Initial1_Form.Option1_Click(0)
                    Call Initial1_Form.Option2_Click(3)
                Case "3.0.0"
                    Call Initial1_Form.Option1_Click(2)
                    Call Initial1_Form.Option2_Click(0)
                Case "3.1.0"
                    Call Initial1_Form.Option1_Click(2)
                    Call Initial1_Form.Option2_Click(1)
                Case "3.2.0"
                    Call Initial1_Form.Option1_Click(2)
                    Call Initial1_Form.Option2_Click(2)
                Case "4.0.0"
                    Call Initial1_Form.Option1_Click(3)
                    Call Initial1_Form.Option2_Click(0)
                Case "4.1.0"
                    Call Initial1_Form.Option1_Click(3)
                    Call Initial1_Form.Option2_Click(1)
                Case "4.2.0"
                    Call Initial1_Form.Option1_Click(3)
                    Call Initial1_Form.Option2_Click(2)
                Case "4.3.0"
                    Call Initial1_Form.Option1_Click(3)
                    Call Initial1_Form.Option2_Click(3)
            End Select

            Initial.Scenario = .Rows(0).Item("Scenario")
            Initial.Input_Files = .Rows(0).Item("SWAT_Input")
            Initial.New_Swat = Initial.Dir1 & .Rows(0).Item("New_SWAT")
            Initial.Output_files = Initial.Dir1 & .Rows(0).Item("APEX")
            Initial.FEM = Initial.Dir1 & .Rows(0).Item("FEM")
            Initial.Swat_Output = Initial.Dir1 & .Rows(0).Item("SWAT_Output")
            .Dispose()
        End With
        proj1 = Nothing

        cons = "SELECT * FROM Runs"
        proj1 = getDBDataTable(cons)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE Runs (Subbasin String)"
            proj1 = getLocalDataTable(cons, Initial.Output_files)
            cons = "SELECT subbasin FROM Sub_included"
            proj1 = getDBDataTable(cons)
            Dim i As UShort
            If proj1.Rows.Count > 0 Then
                For i = 0 To proj1.Rows.Count - 1
                    cons = "INSERT INTO Runs VALUES('" & proj1.Rows(i).Item("subbasin") & "')"
                Next
            End If
        End If

        cons = "SELECT * FROM BMPs"
        proj1 = getLocalDataTable(cons, Initial.Output_files)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE BMPs (Scenario String, BMP_Name String, BMP String, Other String, Input1 Single, Input2 Single, Input3 Single, Input4 Single, Input5 Single, Input6 Single, Input7 Single, Input8 Single)"
            proj1 = getLocalDataTable(cons, Initial.Output_files)
        Else
            If Not proj1.Columns.Contains("Other") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Other String"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Input5") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Input5 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Input6") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Input6 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Input7") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Input7 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Input8") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Input8 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Boimass") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Boimass Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Manure") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Manure Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Urine") Then
                cons = "ALTER TABLE BMPs ADD COLUMN Urine Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
        End If

        cons = "SELECT * FROM BMP_Subareas"
        proj1 = getLocalDataTable(cons, Initial.Output_files)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE BMP_Subareas (BMP String, Subarea String, HRU String, Land_Use String, Soil String)"
            proj1 = getLocalDataTable(cons, Initial.Output_files)
        Else
            If Not proj1.Columns.Contains("Input1") Then
                cons = "ALTER TABLE BMP_Subareas ADD COLUMN Input1 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Input2") Then
                cons = "ALTER TABLE BMP_Subareas ADD COLUMN Input2 Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Biomass") Then
                cons = "ALTER TABLE BMP_Subareas ADD COLUMN Biomass Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Manure") Then
                cons = "ALTER TABLE BMP_Subareas ADD COLUMN Manure Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
            If Not proj1.Columns.Contains("Urine") Then
                cons = "ALTER TABLE BMP_Subareas ADD COLUMN Urine Single"
                getLocalDataTable(cons, Initial.Output_files)
            End If
        End If

        cons = "SELECT * from ANIMALS"
        proj1 = getDBDataTable(cons)
        If Not proj1 Is Nothing Then
            If Not proj1.Columns.Contains("manure_produced") Then
                cons = "ALTER TABLE Animals ADD COLUMN manure_produced Single"
                getDBDataTable(cons)
            End If
            If Not proj1.Columns.Contains("bio_consumed") Then
                cons = "ALTER TABLE Animals ADD COLUMN bio_consumed Single"
                getDBDataTable(cons)
            End If
            If Not proj1.Columns.Contains("Urine_produced") Then
                cons = "ALTER TABLE Animals ADD COLUMN Urine_produced Single"
                getDBDataTable(cons)
            End If
        End If

        cons = "SELECT * FROM BMPs"
        proj1 = getDBDataTable(cons)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE BMPs (BMP_Name String, Status Byte)"
            proj1 = getDBDataTable(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Autoirrigation', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Filter Strip', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Land Leveling', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Manual Modification', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('No Tillage', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Pads and Pipes - Ditch Enlargement and Reservoir System', 0)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Pads and Pipes - No Ditch Improvement', 0)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Pads and Pipes - Tailwater Irrigation', 0)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Pads and Pipes - Two-stage Ditch System', 0)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Permanent Dike', 0)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Ponds', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Stream Fencing', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Waterways', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Wetland', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Rotational Grazing (Dividing Selected Fields)', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Rotational Grazing (Rotating animals in selected fields)', 1)" : modifyRecords(cons)
            cons = "INSERT INTO BMPs(BMP_Name, Status) VALUES ('Manure Application based on Annual Soil P/N', 1)" : modifyRecords(cons)
        Else
            'cons = "DELETE * FROM BMPS"
            'proj1 = getDBDataTable(cons)
        End If
        If Not proj1.Columns.Contains("RCHD") Then
            cons = "ALTER TABLE BMPs ADD COLUMN RCHD Single, RCBW Single, RCTW Single, RCHS Single, RCHN Single, RCHC Single, RCHK Single"
            getDBDataTable(cons)
            cons = "UPDATE BMPs SET RCHD=0.3, RCBW=3.0, RCTW=9.0, RCHS=0.001, RCHN=0.075, RCHC=0.001, RCHK=0.01 WHERE BMP_Name='Waterways'" : modifyRecords(cons)
        End If
        If Not proj1.Columns.Contains("CHD") Then
            cons = "ALTER TABLE BMPs ADD COLUMN CHD Single, CHS Single, CHN Single, UPN Single"
            getDBDataTable(cons)
            cons = "UPDATE BMPs SET CHD=0.0, CHS=0.0, CHN=0.3, UPN=0.3, RCHD=0.0, RCBW=0.0, RCTW=0.0, RCHS=0, RCHN=0.3, RCHC=0.001, RCHK=0.3 WHERE BMP_Name='Filter Strip'" : modifyRecords(cons)
        End If
        If Not proj1.Columns.Contains("RSEE") Then
            cons = "ALTER TABLE BMPs ADD COLUMN RSEE Single,RSVE Single,RSEP Single,RSVP Single,RSV Single,RSRR Single,RSYS Single,RSYN Single,RSHC Single,RSDP Single,RSBD Single"
            getDBDataTable(cons)
            cons = "UPDATE BMPs SET RSEE=0.31,RSVE=50,RSEP=0.3,RSVP=25,RSV=20,RSRR=20,RSYS=300,RSYN=300,RSHC=0.1,RSDP=360,RSBD=0.8 WHERE BMP_Name='Wetland'" : modifyRecords(cons)
        End If
        If Not proj1.Columns.Contains("FFPQ") Then
            cons = "ALTER TABLE BMPs ADD COLUMN FFPQ Single"
            getDBDataTable(cons)
            cons = "UPDATE BMPs SET FFPQ=0.8,CHD=0.0,CHS=0.0,CHN=0.3,UPN=0.3,RCHD=0.0,RCBW=0.0,RCTW=0.0,RCHS=0,RCHN=0.3,RCHC=0.001,RCHK=0.3 WHERE BMP_Name='Stream Fencing'" : modifyRecords(cons)
        End If
        If Not proj1.Columns.Contains("VIMX") Then
            cons = "ALTER TABLE BMPs ADD COLUMN VIMX Single"
            getDBDataTable(cons)
            cons = "UPDATE BMPs SET VIMX=5000 WHERE BMP_Name='Autoirrigation'" : modifyRecords(cons)
        End If

        cons = "SELECT * FROM APEX_Irrigation"
        proj1 = getDBDataTable(cons)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE APEX_Irrigation (IrrigationCode Integer, Description String)"
            proj1 = getDBDataTable(cons)
            'populate it
            cons = "INSERT INTO APEX_Irrigation (IrrigationCode, Description) VALUES (3,'Drip')"
            modifyRecords(cons)
            cons = "INSERT INTO APEX_Irrigation (IrrigationCode, Description) VALUES (7,'Furrow Diking')"
            modifyRecords(cons)
            cons = "INSERT INTO APEX_Irrigation (IrrigationCode, Description) VALUES (2,'Furrow/Flood')"
            modifyRecords(cons)
            cons = "INSERT INTO APEX_Irrigation (IrrigationCode, Description) VALUES (1,'Sprinkler')"
            modifyRecords(cons)
            cons = "INSERT INTO APEX_Irrigation (IrrigationCode, Description) VALUES (8,'TailWater')"
            modifyRecords(cons)
        End If

        cons = "SELECT * FROM Animals"
        proj1 = getDBDataTable(cons)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE Animals (Id Integer, Animal String, Code Integer, dry_manure Single)"
            proj1 = getDBDataTable(cons)
        End If
        cons = "SELECT * FROM Animals"
        proj1 = getDBDataTable(cons)
        If proj1.Rows.Count = 0 Then
            'populate it
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (1,'Dairy', 44, 12, 5.5, 8.1, 11.8)"
            modifyRecords(cons)
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (2,'Beef', 45, 8.5, 3.86, 9.08, 8.21)"
            modifyRecords(cons)
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (3,'Swine', 47, 11, 5.0, 8.1, 17.7)"
            modifyRecords(cons)
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (4,'Sheep', 48, 11, 5.0, 8.1, 6.8)"
            modifyRecords(cons)
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (5,'Goat', 49, 13, 5.9, 8.1, 6.8)"
            modifyRecords(cons)
            cons = "INSERT INTO Animals (Id, Animal, code, dry_manure, manure_produced, bio_consumed,urine_produced) VALUES (6,'Horse', 50, 15, 6.8, 8.1, 4.5)"
            modifyRecords(cons)
        Else
            cons = "UPDATE Animals SET code=44, dry_manure=12, manure_produced=5.5, bio_consumed=8.1,urine_produced=11.8 WHERE id=1"
            modifyRecords(cons)
            cons = "UPDATE Animals SET code=45, dry_manure=8.5, manure_produced=3.86, bio_consumed=9.08,urine_produced=8.21 WHERE id=2"
            modifyRecords(cons)
            cons = "UPDATE Animals SET code=47, dry_manure=11, manure_produced=5.0, bio_consumed=8.1,urine_produced=17.7 WHERE id=3"
            modifyRecords(cons)
            cons = "UPDATE Animals SET code=48, dry_manure=11, manure_produced=5.0, bio_consumed=8.1,urine_produced=6.8 WHERE id=4"
            modifyRecords(cons)
            cons = "UPDATE Animals SET code=49, dry_manure=13, manure_produced=5.9, bio_consumed=8.1,urine_produced=6.8 WHERE id=5"
            modifyRecords(cons)
            cons = "UPDATE Animals SET code=50, dry_manure=15, manure_produced=6.8, bio_consumed=8.1,urine_produced=4.5 WHERE id=6"
            modifyRecords(cons)
        End If

        cons = "SELECT * FROM AnimalsStream"
        proj1 = getDBDataTable(cons)
        If proj1 Is Nothing Then
            cons = "CREATE TABLE AnimalsStream (subarea String, landUse String, seq Integer, HRU String, area Single, field_id Integer, days String, hours Single, animal_type Integer, animals Single)"
            proj1 = getDBDataTable(cons)
        End If

        Database = Dir(Initial.Dir1 & "\Project_Parameters.mdb")
        If Database = "" Then
            File.Copy("Project_Parameters_Org.mdb", Initial.Dir1 & "\Project_Parameters.mdb")
        End If

        cons = "SELECT * FROM crop1310 where Numb = 658"
        proj1 = getDBDataTable(cons)
        If proj1.Rows.Count = 0 Then
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (407,'POAN',6,10.1,40.71,35,0,314,'ANNUAL BLUGRASS(POA)')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (411,'LESP',6,22.5,40.71,35,0,314,'LESPEDEZA GRASS')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (413,'LOVE',6,22.5,40.71,35,0,314,'LOVE GRASS')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (416,'SHBG',6,22.5,40.71,35,0,314,'SHERMAN BLUE GRASS')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (418,'INDI',6,22.5,50.95,35,0,314,'INDIAN GRASS')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (436,'BLUG',6,22.5,40.71,35,0,314,'KENTUCKY BLUEGRASS')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (512,'FRSG',6,22.5,40.71,35,0,314,'Forest-Grasses')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (650,'BGCA',6,10.2,50.90,35,0,314,'California BromeGrass')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (653,'RYEB',6,22.5,40.71,35,0,314,'Blue Wild Rye')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (655,'FESM',6,22.5,40.71,35,0,314,'Molate Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (656,'FESI',6,22.5,40.71,35,0,314,'Idaho Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (657,'FSCR',6,22.5,40.71,35,0,314,'Creeping Red Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO crop1310 (Numb,Name,IDC,PPLP1,PPLP2,OPV5,PlantingCode,HarvestCode,Description) VALUES (658,'FSCW',6,22.5,40.71,35,0,314,'Chewings Fescue')"
            modifyRecords(cons)
        End If

        cons = "SELECT * FROM S_A_CROP where Apex_Code = 658"
        proj1 = getDBDataTable(cons)
        If proj1.Rows.Count = 0 Then
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (407,'POAN','Annual Blue Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (411,'LESP','Lespedeza Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (413,'LOVE','Love Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (416,'SHBG','Sherman blue Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (418,'INDI','Indiana Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (436,'BLUG','Kentucky Bluegrass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (512,'FRSG','Forest - Grass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (650,'BGCA','California Bromegrass')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (653,'RYEB','Blue Wild Rye')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (655,'FESM','Molate Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (656,'FESI','Idaho Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (657,'FSCR','Crreping Red Fescue')"
            modifyRecords(cons)
            cons = "INSERT INTO S_A_CROP (Apex_Code,Name,Description) VALUES (658,'FSCW','Chewings Fescue')"
            modifyRecords(cons)
        End If

        File.Copy(Initial.OrgDir & "\CROP2110.DAT", Initial.Output_files & "\CROP2110.DAT", True)
    End Sub

    Private Sub CheckDir(ByRef Folder1 As String, ByRef DirExist As Object)
        DirExist = Dir(Folder1, FileAttribute.Directory)
    End Sub

    Public Sub ProjectList()
        Dim rsP As DataTable
        rsP = New DataTable
        Initial1_Form.List3.Items.Clear()
        rsP = Nothing
    End Sub

    Public Sub enable_Menu()
        Dim i As Integer
        Dim lastRun As String = String.Empty

        If Not Initial.Scenario Is Nothing Then lastRun = ParmDBName("SELECT LastRun FROM paths WHERE Scenario = '" & Initial.Scenario & "'")
        Me.rbTab_basinLand_Sel.Visible = False
        Me.rbTab_Scenarios.Visible = False
        Me.rbTab_APEX_Gen.Visible = False
        Me.rbTab_ExeModel.Visible = False
        Me.rbTab_ViewResults.Visible = False
        Me.rbTab_DbaseMag.Visible = False
        Me.Projects(2).Enabled = False
        Me.Tools_.Enabled = False

        For i = 1 To rbPnl_APEXFiles.Items.Count - 1
            rbPnl_APEXFiles.Items(i).Enabled = False
            'Me.Create(i).Enabled = False
        Next
        'For i = 0 To rbPnl_RunModel.Items.Count - 1  'This is not working because the options can be run any time. Except FEM that needs APEX and SWAT.
        '    rbPnl_RunModel.Items(i).Enabled = False
        'Next

        Select Case Initial.CurrentOption
            Case 10
                Me.rbTab_basinLand_Sel.Visible = True
            Case 20 To 28
                Me.rbTab_basinLand_Sel.Visible = True
                Me.rbTab_APEX_Gen.Visible = True
                For i = 1 To Initial.CurrentOption - 22
                    rbPnl_APEXFiles.Items(i).Enabled = True
                    'Me.Create(i).Enabled = True
                Next
            Case 30
                Me.rbTab_basinLand_Sel.Visible = True
                Me.rbTab_APEX_Gen.Visible = True
                Me.rbTab_ExeModel.Visible = True
                Me.rbTab_Scenarios.Visible = True
                'Me._Execute_APEX_1.Enabled = True
                'Me._Execute_APEX_2.Enabled = True
                'Me._Execute_APEX_3.Enabled = True
                'Me._Execute_APEX_4.Enabled = True
                'Me._Execute_APEX_5.Enabled = True
                If lastRun = "APEX" Or lastRun = "SWAT" Then Me._Execute_APEX_9.Enabled = True
                Me.rbTab_DbaseMag.Visible = True
                Me.rbTab_Scenarios.Visible = False
                Me.Tools_.Enabled = True
                For i = 1 To rbPnl_APEXFiles.Items.Count - 1
                    rbPnl_APEXFiles.Items(i).Enabled = True
                    'Me.Create(i).Enabled = False
                Next
            Case 31 To 40
                Me.rbTab_basinLand_Sel.Visible = True
                Me.rbTab_APEX_Gen.Visible = True
                Me.rbTab_ExeModel.Visible = True
                Me.rbTab_Scenarios.Visible = True
                'Me._Execute_APEX_1.Enabled = True
                'Me._Execute_APEX_2.Enabled = True
                'Me._Execute_APEX_3.Enabled = True
                'Me._Execute_APEX_4.Enabled = True
                'Me._Execute_APEX_5.Enabled = True
                If lastRun = "APEX" Or lastRun = "SWAT" Then Me._Execute_APEX_9.Enabled = True
                Me.rbTab_DbaseMag.Visible = True
                'Me.rbTab_Scenarios.Visible = False
                Me.Tools_.Enabled = True
                If Initial.CurrentOption <= 35 Then
                    For i = 0 To Initial.CurrentOption - 31
                        rbPnl_RunModel.Items(i).Enabled = True
                        'Me.Create(i).Enabled = True
                    Next
                End If
                Me.rbTab_ViewResults.Visible = True
            Case 40
                '    Me.rbTab_basinLand_Sel.Visible = True
                '    Me.rbTab_APEX_Gen.Visible = True
                '    Me.rbTab_ExeModel.Visible = True
                'Me.rbTab_Scenarios.Visible = False
                '    'Me.Execute_APEX(3).Enabled = True
                '    'Me.Execute_APEX(4).Enabled = True
                '    'Me.Execute_APEX(5).Enabled = True
                '    If lastRun = "APEX" Or lastRun = "SWAT" Then Me.Execute_APEX(9).Enabled = True
                '    Me.rbTab_DbaseMag.Visible = True
                '    Me.rbTab_Scenarios.Visible = True
                '    Me.Tools_.Enabled = True
                '    For i = 1 To rbPnl_APEXFiles.Items.Count - 1
                '        rbPnl_APEXFiles.Items(i).Enabled = True
                '        'Me.Create(i).Enabled = True
                '    Next
                '    Me.rbTab_ViewResults.Visible = True
                'Case 98
                '    'Me.Execute_APEX(3).Enabled = False
                '    'Me.Execute_APEX(4).Enabled = False
                '    'Me.Execute_APEX(5).Enabled = False
                '    If lastRun = "APEX" Or lastRun = "SWAT" Then Me.Execute_APEX(9).Enabled = False
                '    Me.rbTab_ExeModel.Visible = True
                '    Me.rbTab_DbaseMag.Visible = True
                '    Me.rbTab_Scenarios.Visible = True
                '    Me.Tools_.Enabled = True
                '    Me.rbTab_ViewResults.Visible = True
            Case 99
                Me.rbTab_basinLand_Sel.Visible = True
                Me.rbTab_APEX_Gen.Visible = True
                Me.rbTab_ExeModel.Visible = True
                Me.rbTab_Scenarios.Visible = True
                'Me._Execute_APEX_1.Enabled = True
                'Me._Execute_APEX_2.Enabled = True
                'Me._Execute_APEX_3.Enabled = True
                'Me._Execute_APEX_4.Enabled = True
                'Me._Execute_APEX_5.Enabled = True
                If lastRun = "APEX" Or lastRun = "SWAT" Then Me._Execute_APEX_9.Enabled = True
                Me.rbTab_DbaseMag.Visible = True
                'Me.rbTab_Scenarios.Visible = False
                Me.Tools_.Enabled = True
                If Initial.CurrentOption <= 35 Then
                    For i = 0 To Initial.CurrentOption - 31
                        rbPnl_RunModel.Items(i).Enabled = True
                        'Me.Create(i).Enabled = True
                    Next
                End If
                Me.rbTab_ViewResults.Visible = True
                '    Me.Execute_APEX(3).Enabled = True
                '    Me.Execute_APEX(4).Enabled = True
                '    Me.Execute_APEX(5).Enabled = True
                '    If lastRun = "APEX" Or lastRun = "SWAT" Then Me.Execute_APEX(9).Enabled = True
                '    Me.rbTab_ExeModel.Visible = True
                '    Me.rbTab_DbaseMag.Visible = True
                '    Me.rbTab_Scenarios.Visible = True
                '    Me.Tools_.Enabled = True
                '    Me.rbTab_ViewResults.Visible = True
        End Select

        If Initial.Scenario <> "Baseline" Then
            rbBtnAnimalStream.Enabled = False
            rbTab_APEX_Gen.Visible = False
            rbTab_basinLand_Sel.Visible = False
        Else

            rbTab_APEX_Gen.Visible = True
            rbTab_basinLand_Sel.Visible = True
            rbBtnAnimalStream.Enabled = True
        End If

        Me.Ribbon1.Refresh()
    End Sub

    Private Sub ReadLatLong()
        Dim temp1 As String
        Dim pos As Integer
        Dim lon, lat As String
        Dim i As Short
        Dim srFile As StreamReader
        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & Initial.prpfiles(1))) '
        srFile.ReadLine()
        lat = srFile.ReadLine
        lon = srFile.ReadLine
        i = 0
        pos = 8
        temp1 = "Init."

        Do While temp1 <> ""
            Initial.lat1(i) = Mid(lat, pos, 5)
            Initial.lon1(i) = Mid(lon, pos, 5)
            temp1 = Initial.lat1(i)
            i = i + 1
            pos = pos + 5
        Loop

        srFile.Close()
        srFile.Dispose()
        srFile = Nothing

    End Sub

    Private Sub takeHRUpcp()
        Dim mypos As Short
        Dim strings1 As String
        Dim rsP As DataTable
        Dim i As Short
        rsP = New DataTable


        'Yang on 6/6/2014
        strings1 = "DELETE FROM FEM WHERE [Applies To]<>'BMP'"
        modifyLocalRecords(strings1, Initial.Output_files)
        strings1 = "SELECT * FROM APEXHRUs"
        rsP = getLocalDataTable(strings1, Initial.Output_files)

        Dim TakeField As New Convertion
        Dim ADORecordset As DataTable
        Dim value As String
        Dim Rotation As Integer = 0
        ADORecordset = New DataTable
        ADORecordset = getDBDataTable("SELECT * FROM Apexfiles WHERE Apexfile = 'Operations' AND version=" & "'" & Initial.Version & "'" & " ORDER BY line, field")

        With ADORecordset
            For i = 0 To .Rows.Count - 1
                If (Not .Rows(i).Item("SwatFile") Is System.DBNull.Value AndAlso .Rows(i).Item("SwatFile") <> "") Then
                    value = ""
                    Select Case .Rows(i).Item("SwatFile")
                        Case "*.mgt"
                            'TakeField.filename = Initial.Input_Files & "\" & Trim(General.filemgt)
                            'TakeField.Leng = .Rows(i).Item("Leng")
                            'TakeField.LineNum = .Rows(i).Item("Lines")
                            'TakeField.Inicia = .Rows(i).Item("Inicia")
                            'value = TakeField.value()
                    End Select
                Else
                    value = .Rows(i).Item("Value")
                End If

                Select Case .Rows(i).Item("Line")
                    Case 0
                        Rotation = Val(value)
                End Select
            Next
        End With

        ADORecordset.Dispose()
        ADORecordset = Nothing

        Dim myConnection As New OleDb.OleDbConnection
        myConnection.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & SWAPP.Initial.Dir1 & "\Project_Parameters.mdb;"
        If myConnection.State = Data.ConnectionState.Closed Then
            myConnection.Open()
        End If

        With rsP
            For i = 0 To .Rows.Count - 1
                Wait_Form.Pbar_Scenarios.Value = (1 - (.Rows.Count - i) / .Rows.Count) * 100
                If .Rows(i).Item("Apex_Swat") = "SWAT" Then
                    Dataclass.filemgt = .Rows(i).Item("Managment")
                    Dataclass.number = .Rows(i).Item("pcpGage")
                    Dataclass.SWATSubbasins(Mid(.Rows(i).Item("Subarea"), 2, 4), Rotation, myConnection)
                Else
                    mypos = InStr(1, .Rows(i).Item("Managment"), ".", CompareMethod.Text)
                    If mypos >= 10 Then
                        Dataclass.filemgt = Mid(.Rows(i).Item("Managment"), 2, mypos - 1) & "opc"
                    Else
                        Dataclass.filemgt = Mid(.Rows(i).Item("Managment"), 1, mypos) & "opc"
                    End If
                    Dataclass.number = .Rows(i).Item("pcpGage")
                    Dataclass.APEXSubareas(Mid(.Rows(i).Item("Subarea"), 2, 4), myConnection)
                End If
                Wait_Form.Label1(8).Text = "SWAT HRU File transfered to FEM " & (i + 1).ToString & " of " & .Rows.Count.ToString
                Wait_Form.Refresh()
            Next
        End With

        myConnection.Close()

    End Sub

    Public Sub FEM_ControlFile(ByRef kind As Short, ByRef Width_Renamed As Double, ByRef ManCode As String)
        Dim swFile As StreamWriter
        Dim strFEMPath As String = Initial.Dir2 & "\FEM" 'Initial.FEM
        swFile = New StreamWriter(File.Open(strFEMPath & "\SWAPP_FEMOptions.txt", FileMode.Append))

        Select Case kind
            Case 0
                swFile.WriteLine("ManureRateCode," & "Manure")
            Case 1
                swFile.WriteLine("FilterStripWidth," & Width_Renamed)
                swFile.WriteLine("FilterStripCrop," & Initial.OpcsCode)
                swFile.WriteLine("ManureRateCode," & ManCode)
            Case 2
                swFile.WriteLine("ManureRateCode," & ManCode)
        End Select
        Close_Stream_File(swFile)

    End Sub

    Private Sub UpdateEnvironmentVariables()
        Dim cons As String
        Dim proj1 As DataSet

        On Error GoTo goError

        proj1 = getDBDataSet("SELECT CurrentOption FROM Paths")

        With proj1.Tables(0)
            If .Rows(0).Item("CurrentOption") <= Initial.CurrentOption Then
                getDBDataSet("UPDATE paths SET CurrentOption=" & Initial.CurrentOption)
                '.Rows(0).Item("CurrentOption") = Initial.CurrentOption
                Call enable_Menu()
            End If
        End With
        proj1.Dispose()

        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Public Function DetectUrban(ByRef FileToCheck As String) As Short
        Dim temp As String
        Dim srFile As StreamReader
        Dim i As Integer

        srFile = New StreamReader(File.OpenRead(Initial.Input_Files & "\" & FileToCheck))
        i = 1
        For i = 1 To 15
            srFile.ReadLine()
        Next

        temp = srFile.ReadLine
        srFile.Close()
        srFile.Dispose()
        srFile = Nothing
        DetectUrban = Val(CStr(Strings.Left(temp, 16)))
    End Function

    Public Sub ChangeFile(ByRef FileToCheck As String)
        Dim LusePos As Integer
        Dim temp As String
        Dim srFile As StreamReader
        Dim swFile As StreamWriter

        srFile = New StreamReader(Initial.Input_Files & "\" & FileToCheck)
        swFile = New StreamWriter(File.Create(Initial.Input_Files & "\Temporal"))

        temp = srFile.ReadLine
        LusePos = InStr(1, temp, "Luse:")
        LusePos = LusePos + 5
        Mid(temp, LusePos, 4) = "URBN"
        swFile.WriteLine(temp)

        Do While srFile.EndOfStream <> True
            swFile.WriteLine(srFile.ReadLine)
        Loop

        swFile.Close()
        swFile.Dispose()
        srFile.Close()
        srFile.Dispose()
        srFile = Nothing
        swFile = Nothing
        File.Copy(Initial.Input_Files & "\Temporal", Initial.Input_Files & "\" & FileToCheck, True)

    End Sub

    Private Sub Convert00to03()
        Dim myval As Integer

        'fs = CreateObject("Scripting.FileSystemObject")

        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
        Wait_Form.Label1(1).Text = "The SWAT Files are Being converted from SWAT2000 to SWAT2005"
        Wait_Form.Show()

        On Error Resume Next
        MkDir(Initial.Input_Files & "\swat2005")

        On Error GoTo goError
        File.Copy(Initial.Input_Files & "\*.*", Initial.Input_Files & "\swat2005")
        File.Copy(Initial.OrgDir & "\conv00to03.exe", Initial.Input_Files & "\swat2005\conv00to03.exe")
        myval = Me.ExecCmd(Initial.Input_Files & "\swat2005\Conv00to03.exe", Initial.Input_Files & "\swat2005")

        If myval <> 0 Then
            Exit Sub
        End If

        Initial.Input_Files = Initial.Input_Files & "\swat2005"
        Call Initial1_Form.Option2_CheckedChanged(Nothing, New System.EventArgs())
        Wait_Form.Close()
        MsgBox("The SWAT Files were Converted into the Folder " & Initial.Input_Files, MsgBoxStyle.Information)
        Exit Sub

goError:
        MsgBox(Err.Description & " --> Subroutine Convert00to03")
    End Sub

    Private Sub figAreas()
        Dim i As UShort = 0
        Dim SWTFiles As String = String.Empty
        Dim tempRecord As String = String.Empty
        Dim SWTLines, SWTCol As Short
        Dim srFile As StreamReader = Nothing
        Dim swCurrSwt As StreamWriter = New StreamWriter(File.Create(Initial.Swat_Output & "\currswt.txt"))

        Try
            convertFormat = New Convertion
            'fs = CreateObject("Scripting.FileSystemObject")
            SWTLines = 7
            SWTCol = 19
            SWTFiles = Dir(Initial.Output_files & "\*.swt")

            Do While SWTFiles <> ""
                srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\" & SWTFiles))
                If Initial.Version = "4.1.0" Or Initial.Version = "4.0.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                    SWTLines = 6
                    SWTCol = 12
                End If
                For i = 1 To SWTLines
                    tempRecord = srFile.ReadLine
                Next
                SWTArea = Val(Mid(srFile.ReadLine, 30, 8))
                SWTArea = SWTArea / 100
                SWTNames(Val(Mid(tempRecord, 19, 4))) = convertFormat.Convert(SWTArea, "##0.00")
                swCurrSwt.WriteLine(SWTFiles.Substring(0, 4))
                SWTFiles = Dir()
            Loop

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not srFile Is Nothing Then
                srFile.Close()
                srFile.Dispose()
                srFile = Nothing
            End If

            If Not swCurrSwt Is Nothing Then
                swCurrSwt.Close()
                swCurrSwt.Dispose()
                swCurrSwt = Nothing
            End If
        End Try
    End Sub

    Private Sub createAPEXBat()
        Dim apexrunall As StreamReader
        Dim apexrunx As StreamWriter
        Dim swFile As StreamWriter = Nothing
        Dim ADORecordset As DataTable
        Dim temp As String
        Dim tempo As String
        Dim rec As String
        Dim i As Integer

        Try

            swFile = New StreamWriter(File.Create(Initial.Output_files & "\APEXBat.txt"))
            swFile.Write("del *.swt")

            ADORecordset = getDBDataTable("SELECT SubBasin FROM Sub_Included ")

            With ADORecordset
                For i = 0 To .Rows.Count - 1
                    'cwg DBNull check added
                    If Not IsDBNull(.Rows(i).Item("SubBasin")) AndAlso .Rows(i).Item("SubBasin") <> "" Then
                        tempo = .Rows(i).Item("Subbasin")
                        temp = Mid(tempo, 2, 8)
                        apexrunx = New StreamWriter(File.Create(Initial.Output_files & "\" & temp & ".dat"))
                        apexrunall = New StreamReader(File.OpenRead(Initial.Output_files & "\APEXRUN.dat"))
                        Do While Not apexrunall.EndOfStream
                            rec = apexrunall.ReadLine
                            If Strings.Left(rec, 8) = CDbl(temp) Then
                                apexrunx.WriteLine(rec)
                                swFile.WriteLine("")
                                swFile.WriteLine("del apexrun.dat")
                                swFile.WriteLine("copy " & temp & ".dat apexrun.dat")
                                Select Case Initial.Version
                                    Case "4.1.0", "4.0.0", "4.2.0", "4.3.0"
                                        swFile.WriteLine("apex0604.exe")
                                    Case "1.1.0", "1.2.0", "1.3.0"
                                        swFile.WriteLine("apex0806.exe")
                                    Case Else
                                        swFile.WriteLine("apex2110.exe")
                                End Select
                                Exit Do
                            End If
                        Loop
                        apexrunx.Close()
                        apexrunx.Dispose()
                        apexrunx = Nothing
                        apexrunall.Close()
                        apexrunall.Dispose()
                        apexrunall = Nothing
                    End If
                Next
            End With
            swFile.Close()
            swFile.Dispose()
            swFile = Nothing

            File.Copy(Initial.Output_files & "\APEXRUN.dat", Initial.Output_files & "\APEXRunAll.dat", True)
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            If Not swFile Is Nothing Then
                swFile.Close()
                swFile.Dispose()
                swFile = Nothing
            End If
        End Try

    End Sub

    Public Sub HideColumns(ByVal grid As AxMSDataGridLib.AxDataGrid)
        'For col As Integer = 0 To grid.Columns.Count - 1
        '    grid.Columns(col).Visible = False
        'Next col
    End Sub

    'Function doAPEXProcess(ByVal sRunBat As String) As Integer
    '    Dim myProcess As Process = New Process
    '    Dim i As Integer
    '    Dim sReturn As Integer = 0

    '    Try
    '        ' set the file name and the command line args
    '        myProcess.StartInfo.FileName = "cmd.exe"
    '        myProcess.StartInfo.Arguments = "/C " & sRunBat & " " & Chr(34) & " && exit"
    '        ' start the process in a hidden window
    '        'myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Hidden
    '        myProcess.StartInfo.CreateNoWindow = True
    '        ' allow the process to raise events
    '        myProcess.EnableRaisingEvents = True
    '        ' add an Exited event handler
    '        'AddHandler myProcess.Exited, AddressOf processNLEAPExited
    '        myProcess.Start()
    '        For i = 0 To 10000000
    '            If myProcess.HasExited Then
    '                Exit For
    '            End If
    '        Next i

    '        If myProcess.ExitCode = 0 Then
    '            sReturn = 0
    '        Else
    '            sReturn = myProcess.ExitCode
    '        End If

    '        myProcess.Close()
    '        myProcess.Dispose()
    '        Return sReturn

    '    Catch ex As Exception
    '        Return ex.Message
    '    End Try
    'End Function

    Private Sub _stdFiles_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _stdFiles_1.Click
        If Not IsAvailable(txtFiles_Form) Then txtFiles_Form = New txtFiles()
        txtFiles_Form.Tag = Initial.Swat_Output & "\input.std"
        txtFiles_Form.Show()
    End Sub

    Private Sub _stdFiles_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _stdFiles_2.Click
        If Not IsAvailable(txtFiles_Form) Then txtFiles_Form = New txtFiles()
        txtFiles_Form.Tag = Initial.Swat_Output & "\output.std"
        txtFiles_Form.Show()
    End Sub

    Private Sub Help1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Help1.Click
        Dim index As Short = 1
        Dim helpV As Object
        'cwg TODO
        With CommonDialog1Open
            Select Case index
                Case 1
                    helpV = Initial.OrgDir & "\" & "SWAPP.hlp"
                    '.HelpFile = helpV
                    '.HelpCommand = MSComDlg.HelpConstants.cdlHelpKey
                    '.HelpCommand = cdlHelpSetContents
                    '.ShowHelp()
                Case 2
                    helpV = Initial.OrgDir & "\" & "APEX2110 User Manual.hlp"
                    '.HelpFile = helpV
                    '.HelpCommand = MSComDlg.HelpConstants.cdlHelpKey
                    '.ShowHelp()
                Case 4
                Case 6
                    helpV = Initial.OrgDir & "\" & "Tutorial.hlp"
                    '.HelpFile = helpV
                    '.HelpCommand = MSComDlg.HelpConstants.cdlHelpKey
                    '.ShowHelp()
            End Select

        End With

    End Sub

    Private Sub _help_1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _help_1.Click
        Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\SWAPP.hlp")
    End Sub

    Private Sub _help_2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _help_2.Click
        Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\APEX2110 USER MANUAL.HLP")
    End Sub

    Private Sub _help_4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _help_4.Click
        'frmAbout.ShowDialog()
    End Sub

    Private Sub _help_6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles _help_6.Click
        Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\tutorial.HLP")
    End Sub

    Private Sub GenerateRepresentaiveFarmsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GenerateRepresentaiveFarmsToolStripMenuItem.Click
        Dim RepFarm_bat As String
        Dim tempVal As Object
        'direxe = Initial.Output_files
        RepFarm_bat = "RunCensus.bat"
        'fileo = Initial.FEM & "\SWAPP_FEMOptions.txt"
        'filet = Initial.OrgDir & "\FEM\SWAPP_FEMOptions.txt"
        'tempVal = ExecCmd(Initial.Output_files & "\" & RepFarm_bat, direxe)
        tempVal = Me.ExecCmd(Initial.OrgDir & "\fem\" & RepFarm_bat, Initial.OrgDir & "\fem")
        If tempVal = 0 Then
            MsgBox("MS-DOS " & Initial.OrgDir & "\fem\" & RepFarm_bat & " Process Was Successfully Executed", , "Confirmation")
        Else
            Exit Sub
        End If
    End Sub

    Private Sub rbBtn_New_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_New.Click
        'copied code from Old style menu
        'On Error GoTo goError
        If Not IsAvailable(Initial1_Form) Then Initial1_Form = New Initial1()
        Dim Check_Local As Object
        Dim Database As String
        Dim myfile As String
        'Dim cn1 As ADODB.Connection
        Dim proj1 As DataTable
        Initial1_Form.CreateFolder.Visible = True
        'cn1 = New ADODB.Connection
        proj1 = New DataTable
        'cwg Initial1_Form.HelpContextID = 21
        Call CloseFrames()
        Call OpenFrames(7)
        Initial1_Form.ShowDialog()
        If Initial.cancel1 = 1 Then Exit Sub
        Call CloseFrames()
        Call OpenFrames(1)
        Initial1_Form.ShowDialog()
        If Initial.cancel1 = 1 Then Exit Sub
        Call CloseFrames()
        Call OpenFrames(2)
        Initial1_Form.ShowDialog()
        If Initial.cancel1 = 1 Then Exit Sub
        Call CloseFrames()
        myfile = Dir(Initial.Input_Files & "\file.cio")
        If myfile <> "file.cio" Then
            MsgBox("The Path Entered for SWAT Original Files is not Correct", , "Initial - Subdirectories Selection")
            Exit Sub
        End If
        If Initial.MsgBox_Answer = 6 Then Call Convert00to03()
        Initial.Scenario = "Baseline"
        Initial.Dir2 = Initial.Dir1
        Initial.Scenario = "Baseline"
        Initial.CurrentOption = 10
        Initial.Output_files = Initial.Dir2 & "\APEX"
        Initial.New_Swat = Initial.Dir2 & "\New_SWAT"
        Initial.FEM = Initial.Dir2 & "\FEM"
        Initial.Swat_Output = Initial.Dir2 & "\SWAT_Output"
        If Initial.Errors = 1 Then
            Initial.Errors = 0
            Exit Sub
        End If
        myfile = Dir(Initial.OrgDir & "\FEM", FileAttribute.Directory)
        If myfile <> "FEM" Then
            MkDir((Initial.OrgDir & "\FEM"))
        End If
        Database = Dir(Initial.Dir1 & "\Project_Parameters.mdb")

        If Database = "" Then File.Copy(Initial.OrgDir & "\Project_Parameters_Org.mdb", Initial.Dir1 & "\Project_Parameters.mdb")
        Call enable_Menu()

        Call Initial1_Form.Assign_Version()
        Call Form_Initials()
        Call Initial1_Form.Assign_Version()
        Call SaveEnvironmentVariables()
        Call FEM_ControlFile(0, 49.99, "Manure")
        Check_Local = Dir(Initial.Output_files & "\Local.mdb")
        If Check_Local = "" Then File.Copy(Initial.OrgDir & "\Local.mdb", Initial.Output_files & "\Local.mdb")
        OpenEnvironmentVariables("Baseline")
        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        Exit Sub
    End Sub

    Private Sub rbBtn_Open_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_Open.Click
        'On Error GoTo goError
        If Not IsAvailable(Initial1_Form) Then Initial1_Form = New Initial1()
        Initial1_Form.CreateFolder.Visible = False
        'cwg Initial1_Form.HelpContextID = 22
        Call OpenFrames(9)
        Initial1_Form.ShowDialog()
        'cwg xx Initial1_Form.Width = 5385
        If Initial.cancel1 = 1 Then Exit Sub
        'If Initial.NumProj = 0 Then
        '    MsgBox("There are not Projects Created. Select Project --> New to Create a New One")
        '    Exit Sub
        'End If
        Call CloseFrames()
        Call OpenEnvironmentVariables(Initial.Scenario)
        Call enable_Menu()
        If Not IsAvailable(Create_Files_Form) Then Create_Files_Form = New Create_Files
        Call Form_Initials()
        Call Initial1_Form.Assign_Version()
        'Call FEM_ControlFile(0, 49.99, "P")
        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        Exit Sub
        'goError:
        '        MsgBox(Err.Description & " --> Subroutine Projects")
        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        Exit Sub
    End Sub

    Private Sub rbBtn_Close_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_Close.Click
        'Dim index As Short = Close_Renamed.GetIndex(eventSender)
        Me.Close()
    End Sub

    Private Sub rbBtn_SelDel_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_SelDel.Click
        'Scenarios_Select/Delete
        If Not IsAvailable(ScenariosList_Form) Then ScenariosList_Form = New ScenariosList()
        ScenariosList_Form.ShowDialog()
        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        Call enable_Menu()
    End Sub

    Private Sub rbBtn_Create_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_Create.Click
        If Not IsAvailable(ScenariosList_Form) Then ScenariosList_Form = New ScenariosList()
        If Not IsAvailable(NewScenario1_Form) Then NewScenario1_Form = New NewScenario1()
        NewScenario1_Form.ShowDialog()
        ScenariosList_Form.ShowDialog()
        '<06/28/2013
        Me.Text = " CEEOT-SWAPP. Project --> " & Initial.Project & "   Scenario --> " & Initial.Scenario
        '>
        Call enable_Menu()
    End Sub

    Private Sub rbBtn_Select_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_Select.Click
        Call S_Include_Click(1)
        Call UpdateEnvironmentVariables()
    End Sub

    Private Sub rbBnt_SelSubs_Click(sender As System.Object, e As System.EventArgs) Handles rbBnt_SelSubs.Click
        'Dim ans As Integer
        'Dim subbasinNumber As Object
        'Dim temp As Object
        'Dim Mensaje As String
        'Dim SWATFile As String
        'Dim i As Integer
        'Dim Swat_bat As String
        'Dim Sw_bat As String
        'Dim myval As Object
        'Dim direxe As String
        'Dim Check_Local As String
        Dim name_Renamed As String = "rbBnt_SelSubs_Click"
        'Dim adoRec As DataTable
        'Dim swFile As StreamWriter = Nothing
        'Dim srFile As StreamReader
        'Dim bat_file, files As String
        'Dim tempDT As DataTable

        Try
            'myval = 0

            SelectVersion()
            'direxe = Initial.Output_files

            If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()

            With Wait_Form
                If Not IsAvailable(frmSelectSubBasins_Form) Then frmSelectSubBasins_Form = New frmSelectSubBasins()
                frmSelectSubBasins_Form.ShowDialog()
            End With

            Initial.CurrentOption = 33
            enable_Menu()
            Call UpdateEnvironmentVariables()
            Wait_Form.Close()

        Catch ex As Exception
            MsgBox(Err.Description, , "Execute_Apex - " & name_Renamed)
        Finally
            'swFile.Close()
            'swFile.Dispose()
            'swFile = Nothing
        End Try
    End Sub

    '    Private Sub rbBnt_ExecuteAll_Click(index As UShort)
    '        'Dim ans As Integer
    '        Dim subbasinNumber As Single = 0
    '        'Dim temp As Object
    '        Dim Mensaje As String
    '        Dim SWATFile As String
    '        Dim i As Integer
    '        'Dim Swat_bat As String = String.Empty
    '        'Dim Sw_bat As String = String.Empty
    '        Dim myval As Object
    '        Dim direxe As String
    '        'Dim Check_Local As String
    '        Dim name_Renamed As String
    '        'Dim adoRec As DataTable
    '        Dim swFile As StreamWriter = Nothing
    '        Dim srFile As StreamReader
    '        Dim bat_file, files As String
    '        'Dim tempDT As DataTable

    '        On Error GoTo goError

    '        ValidateSubasinsSelected()
    '        SelectVersion()

    '        myval = 0

    '        direxe = Initial.Output_files

    '        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()

    '        With Wait_Form
    '            .Label1(1).Text = "The APEX Program is Being Executed"
    '            .Show()
    '            srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\APEXBat.txt"))
    '            Do While srFile.EndOfStream <> True
    '                swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & APEX_bat))
    '                For i = 1 To 5
    '                    swFile.WriteLine(srFile.ReadLine)
    '                Next
    '                swFile.Close()
    '                myval = ExecCmd(Initial.Output_files & "\" & APEX_bat, direxe)

    '                If myval <> 0 Then
    '                    MsgBox("Subbasins " & subbasinNumber & " Has Problems - Check it out and try again", MsgBoxStyle.Information, "APEX")
    '                    Exit Sub
    '                End If
    '            Loop

    '            srFile.Close()
    '            srFile.Dispose()
    '            srFile = Nothing
    '            name_Renamed = "*.SWT"
    '            Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "*.SWT")
    '            name_Renamed = "SW*"
    '            On Error Resume Next
    '            Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "*.SW")
    '            On Error GoTo 0
    '            SWATFile = Dir(Initial.OrgDir & "\SWAT200*.exe")
    '            If SWATFile <> "" Then Dataclass.copyAllFiles(Initial.OrgDir, Initial.Swat_Output, "SWAT200*.exe")
    '            .Label1(3).Text = "The SWAT to APEX Interface is Being Executed"
    '            .Check1(1).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(1).Visible = True
    '            .Refresh()
    '            name_Renamed = "SW*.bat"
    '            Dataclass.copyAllFiles(Initial.OrgDir, Initial.Output_files, "sw*.bat")
    '            myval = ExecCmd(Initial.Swat_Output & "\" & Sw_bat, Initial.Swat_Output)

    '            If myval <> 0 Then
    '                Exit Sub
    '            End If
    '            'save all of the areas to add in the fig file in SWAT
    '            Call figAreas()
    '            'add input source files from APEX to fig file in SWAT
    '            Call Copyfig()
    '            name_Renamed = "SWat.bat"
    '            On Error Resume Next
    '            Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "\swat*.bat")
    '            On Error GoTo 0
    '            name_Renamed = "*.*"
    '            Dataclass.copyAllFiles(Initial.New_Swat, Initial.Swat_Output, "*.*")
    '            .Label1(5).Text = "The SWAT Program is Being Executed"
    '            .Check1(3).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(3).Visible = True
    '            .Refresh()

    '            myval = ExecCmd(Initial.Swat_Output & "\" & Swat_bat, Initial.Swat_Output)

    '            'Reach.ReadFile   changed to take saved files instead reach file
    '            Reach.ReadSaveFile()
    '            .Check1(5).CheckState = System.Windows.Forms.CheckState.Checked
    '            .Check1(5).Visible = True
    '            .Refresh()
    '            Mensaje = "APEX and SWAT Programs within SWAPP Were Successfully Executed"
    '            modifyRecords("UPDATE paths SET LastRun='APEX' WHERE Scenario='" & Initial.Scenario & "'")

    '            If index = 0 Then
    '                .Label1(7).Text = "The FEM Program is Being Executed"
    '                Execute_FEM_Click(_Execute_FEM_1, New System.EventArgs())
    '                .Check1(7).CheckState = System.Windows.Forms.CheckState.Checked
    '                .Check1(7).Visible = True
    '                .Refresh()
    '                Mensaje = "APEX, SWAT, and FEM Programs within SWAPP Were Successfully Executed"
    '            End If

    '            If myval = 0 Then
    '                MsgBox(Mensaje, , "Confirmation")
    '            End If
    '        End With

    '        Initial.CurrentOption = 40
    '        Call UpdateEnvironmentVariables()
    '        Wait_Form.Close()
    '        On Error Resume Next
    '        swFile.Close()
    '        swFile.Dispose()
    '        swFile = Nothing
    '        Exit Sub

    'goError:
    '        MsgBox(Err.Description, , "Execute_Apex - " & name_Renamed)
    '    End Sub

    Private Sub Execute_APEX_SWAT_FEM()
    End Sub

    Private Sub rbBnt_Environment_Click(sender As System.Object, e As System.EventArgs) Handles rbBnt_Environment.Click
        If Not IsAvailable(Scenarios_List_Form) Then Scenarios_List_Form = New Scenarios_List()
        Scenarios_List_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_BMPsParameters_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_BMPsParameters.Click
        If Not IsAvailable(BMPsParameters_Form) Then BMPsParameters_Form = New BMPsParameters()
        BMPsParameters_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_APEXFile_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_APEXFile.Click
        If Not IsAvailable(Edit_APEX_inputs1_Form) Then Edit_APEX_inputs1_Form = New Edit_APEX_Inputs1()
        Edit_APEX_inputs1_Form.Show()
    End Sub

    Private Sub rbBtn_SWATFile_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_SWATFile.Click
        If Not IsAvailable(Edit_Subbasins_Inputs_Form) Then Edit_Subbasins_Inputs_Form = New Edit_Subbasins_Inputs()
        Edit_Subbasins_Inputs_Form.Tag = "Swat"
        Edit_Subbasins_Inputs_Form.Show()
    End Sub

    Private Sub rbBtn_SWAPPHelp_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_SWAPPHelp.Click
        Try
            Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\SWAPP.hlp")
        Catch ex As Exception
            ' Display the message.
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub rbBtn_SWAPTut_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_SWAPTut.Click
        Try
            Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\tutorial.HLP")
        Catch ex As Exception
            ' Display the message.
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub rbBtn_APEXHelp_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_APEXHelp.Click
        Try
            Help.ShowHelp(Create_Files_Form, "C:\CEEOTSWAPP_2005\SWAPP.NET1\APEX2110 USER MANUAL.HLP")
        Catch ex As Exception
            ' Display the message.
            Console.WriteLine(ex.Message)
        End Try
    End Sub

    Private Sub rbBtn_TotalR_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_TotalR.Click
        If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
        Resultsfrm_Form.Tag = "Total"
        Resultsfrm_Form.ShowDialog()
        'If Not IsAvailable(ResultsTotal_Form) Then ResultsTotal_Form = New ResultsTotal()
        'ResultsTotal_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_YearR_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_YearR.Click
        If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
        Resultsfrm_Form.Tag = "Year"
        Resultsfrm_Form.ShowDialog()
        'If Not IsAvailable(ResultsYear_Form) Then ResultsYear_Form = New ResultsYear()
        'ResultsYear_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_MonthR_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_MonthR.Click
        If Not IsAvailable(Resultsfrm_Form) Then Resultsfrm_Form = New Resultsfrm()
        Resultsfrm_Form.Tag = "Month"
        Resultsfrm_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_InputSWAT_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_InputSWAT.Click
        If Not IsAvailable(txtFiles_Form) Then txtFiles_Form = New txtFiles()
        txtFiles_Form.Tag = Initial.Swat_Output & "\input.std"
        txtFiles_Form.Show()
    End Sub

    Private Sub rbBtn_OutputSWAT_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_OutputSWAT.Click
        If Not IsAvailable(txtFiles_Form) Then txtFiles_Form = New txtFiles()
        txtFiles_Form.Tag = Initial.Swat_Output & "\output.std"
        txtFiles_Form.Show()
    End Sub

    Private Sub rbBtn_APEXR_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_APEXR.Click
        If Not IsAvailable(Out_Files_Form) Then Out_Files_Form = New Out_Files()
        Out_Files_Form.Show()
    End Sub

    Private Sub rbBtn_AnnalFEM_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_AnnalFEM.Click
        Initial.FEMRes = 0
        If Not IsAvailable(ResultsFEMfrm_Form) Then ResultsFEMfrm_Form = New ResultsFEMfrm()
        ResultsFEMfrm_Form.Show()
    End Sub

    Private Sub rbBtn_AverFEM_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_AverFEM.Click
        Initial.FEMRes = 1
        If Not IsAvailable(ResultsFEMfrm_Form) Then ResultsFEMfrm_Form = New ResultsFEMfrm()
        ResultsFEMfrm_Form.Show()
    End Sub

    Private Sub rbBtn_EconomicR_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_EconomicR.Click
        If Not IsAvailable(Scenarios_FEM_Form) Then Scenarios_FEM_Form = New Scenarios_FEM()
        Scenarios_FEM_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_SumRes_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_SumRes.Click
        On Error Resume Next
        If Not IsAvailable(Scenarios_Summary_Form) Then Scenarios_Summary_Form = New Scenarios_Summary()
        Scenarios_Summary_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_LoadM_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_LoadM.Click
        Dim ADORecordset As DataTable
        Dim sites As String
        Dim j As Integer
        Dim i As Integer
        Dim appExcel As excel.Application
        Dim wBookExcel As excel.Workbook
        Dim wSheetExcel As excel.Worksheet
        Dim sitesToUpload(0) As Short
        Dim currentSite As Short
        Dim flow, sed, orgN, orgP, NO3, minP, totalN, totalP As Single
        Dim site, year, mon As Short
        Dim myConnection As OleDb.OleDbConnection
        Dim dbConnectString As String
        Dim command As OleDb.OleDbCommand
        Dim query As String

        'define connection parms
        dbConnectString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Output_files & "\Local.mdb;"
        myConnection = New OleDb.OleDbConnection
        myConnection.ConnectionString = dbConnectString
        myConnection.Open()

        If Not IsAvailable(SelectFile_Form) Then SelectFile_Form = New SelectFile()
        SelectFile_Form.GroupBox1.Visible = True
        SelectFile_Form.Combo1.Items.Clear()
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.Items.Add("Excel (*.xls*)")
        SelectFile_Form.Combo1.SelectedIndex = 6
        SelectFile_Form.ShowDialog()

        appExcel = New excel.Application
        If Initial.measuredfile Is Nothing Then Exit Sub
        wBookExcel = appExcel.Workbooks.Open(Initial.measuredfile)
        wSheetExcel = wBookExcel.Worksheets(1)
        Me.Cursor = Cursors.WaitCursor
        With wSheetExcel
            i = 1
            j = 0
            currentSite = 0
            If Not IsNumeric(.Cells(i, 1).value) Then
                i = i + 1
            End If
            Do While .Cells(i, 1).value <> 0
                i = i + 1
                If .Cells(i, 1).value <> currentSite And Not IsNothing(.Cells(i, 1).value) Then
                    ReDim Preserve sitesToUpload(j)
                    sitesToUpload(j) = .Cells(i, 1).value
                    j = j + 1
                    currentSite = .Cells(i, 1).value
                End If
            Loop
        End With

        sites = sitesToUpload(0)
        For i = 1 To UBound(sitesToUpload)
            sites = sites & " OR Site=" & sitesToUpload(i)
        Next

        ADORecordset = New DataTable
        modifyLocalRecords("DELETE * FROM Measured WHERE Site=" & sites, Initial.Output_files)
        With wSheetExcel
            i = 2
            Do While .Cells(i, 1).value <> 0
                flow = 0 : sed = 0 : NO3 = 0 : minP = 0 : orgN = 0 : orgP = 0 : totalN = 0 : totalP = 0
                site = .Cells(i, 1).value
                year = .Cells(i, 2).value
                mon = .Cells(i, 3).value
                If IsNumeric(.Cells(i, 4).value) Then flow = .Cells(i, 4).value
                If IsNumeric(.Cells(i, 5).value) Then sed = .Cells(i, 5).value
                If IsNumeric(.Cells(i, 6).value) Then NO3 = .Cells(i, 6).value
                If IsNumeric(.Cells(i, 7).value) Then minP = .Cells(i, 7).value
                If IsNumeric(.Cells(i, 8).value) Then orgN = .Cells(i, 8).value
                If IsNumeric(.Cells(i, 9).value) Then orgP = .Cells(i, 9).value
                If IsNumeric(.Cells(i, 10).value) Then totalN = .Cells(i, 10).value
                If IsNumeric(.Cells(i, 11).value) Then totalP = .Cells(i, 11).value
                query = "INSERT INTO Measured (site,[year],[mon],flow,sed,NO3,minP,OrgN,orgP,totalN,totalP) " &
                      "VALUES(" & site & "," & year & "," & mon & "," & flow & "," & sed & "," & NO3 &
                      "," & minP & "," & orgN & "," & orgP & "," & totalN & "," & totalP & ")"
                command = New OleDb.OleDbCommand(query, myConnection)
                command.ExecuteNonQuery()
                i = i + 1
            Loop
        End With
        Me.Cursor = Cursors.Default
        MsgBox("Measured values were uploaded", MsgBoxStyle.OkOnly)
    End Sub

    Private Sub rbBtn_DispM_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_DispM.Click
        If Not IsAvailable(MeasuredValues_Form) Then MeasuredValues_Form = New MeasuredValues()
        MeasuredValues_Form.Tag = "Month"
        MeasuredValues_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_CalER_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_CalER.Click
        Dim adoSites As DataTable

        adoSites = New DataTable
        adoSites = getLocalDataTable("SELECT DISTINCT site FROM Measured ORDER BY Site", Initial.Output_files)

        If Not adoSites Is Nothing Then
            If adoSites.Rows.Count = 0 Then
                MsgBox("There is not monthly measured values - Upload monthly muesured values and try again")
            Else
                If Not IsAvailable(frmEvalues_Form) Then frmEvalues_Form = New frmEvalues()
                frmEvalues_Form.ShowDialog()
            End If
        End If
    End Sub

    Private Sub rbBtn_DispER_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_DispER.Click
        If Not IsAvailable(Measured_Predicted_EValues_Form) Then Measured_Predicted_EValues_Form = New Measured_Predicted_EValues()
        Measured_Predicted_EValues_Form.Tag = "Month"
        Measured_Predicted_EValues_Form.ShowDialog()
    End Sub

    Private Sub rbBtn_MvsP_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_MvsP.Click
        If Not IsAvailable(MeasuredData_Form) Then MeasuredData_Form = New MeasuredData()
        MeasuredData_Form.Tag = "Month"
        MeasuredData_Form.Show()
    End Sub

    Private Sub RibbonButton1_Click(sender As System.Object, e As System.EventArgs)
        MessageBox.Show("welcome")
    End Sub

    Private Sub rbBtn_AllAPEX_Click(sender As System.Object, e As System.EventArgs) Handles rbBtn_AllAPEX.Click
        'Dim Pest_File As String
        'Dim h, d, g, p As StreamWriter
        'Dim subarea As String

        On Error GoTo goError

        If ValidateLandUses() Then
            Exit Sub
        End If
        'Create all of the APEX Files
        'Call Create_all_Click()
        'create control file
        rbBtn_ControlFile_Click(sender, e)
        'create operation files
        rbBtn_OperationFiles_Click(sender, e)
        'create subarea files
        rbBtn_SubareaFiles_Click(sender, e)
        'create soil files
        rbBtn_SoilFiles_Click(sender, e)
        'create site files
        rbBtn_SiteFile_Click(sender, e)
        'create weather files
        rbBtn_WeatherFiles_Click(sender, e)
        'create weather stationfiles
        rbBtn_WPMFiles_Click(sender, e)
        'Initial.CurrentOption = 30
        'enable_Menu()
        'Call createAPEXBat()
        '*****************************
        'Call UpdateEnvironmentVariables()
        'Wait_Form.Close()
        'cwg Me.MousePointer = 0 'or Default
        MsgBox("APEX Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Exit Sub
goError:
        MsgBox(Err.Description)
    End Sub

    Private Sub rbBtn_ControlFile_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnControl.Click
        Dim Pest_File As String

        On Error GoTo goError
        If ValidateLandUses() Then
            Exit Sub
        End If

        'Create Control files
        Wait_Form.Label1(3).Text = "The APEX Control File is Being Created"
        Wait_Form.Show()
        Wait_Form.Pbar_Scenarios.Value = 0
        Wait_Form.Refresh()
        Control.Apexcont((0))
        Pest_File = Dir(Initial.Input_Files & "\Pest.dat")
        If Pest_File <> "" And Pest_File <> " " Then Control.Pesticide()
        Control.Fertilizer()
        Wait_Form.Pbar_Scenarios.Value = 50
        Wait_Form.Refresh()
        If Not sender.text = "Create" Then MsgBox("APEX Control File Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 23
        If Initial.Version = "" Then Initial.CurrentOption = 23
        Me._Create_3.Enabled = True
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        Wait_Form.Pbar_Scenarios.Value = 100
        Wait_Form.Refresh()
        Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)
    End Sub

    Private Sub rbBtn_OperationFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnOperation.Click
        Dim g As StreamWriter

        Try
            If ValidateLandUses() Then
                Exit Sub
            End If

            'Create Control files
            Wait_Form.Label1(3).Text = "The APEX General Files are Being Copied"
            Wait_Form.Show()
            Call cpyApex()
            Wait_Form.Label1(3).Text = "The APEX Operation Files are Being Created"
            Wait_Form.Show()
            Wait_Form.Pbar_Scenarios.Value = 0
            Wait_Form.Refresh()
            g = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Opcs))
            g.Close()
            g.Dispose()
            g = Nothing
            Create_Files_Form.Tag = "4"
            Sitefiles.FEMFiles(3.ToString)
            Wait_Form.Pbar_Scenarios.Value = 25
            Wait_Form.Refresh()
            Sitefiles.SiteFiles(3.ToString)
            Wait_Form.Pbar_Scenarios.Value = 100
            Wait_Form.Refresh()
            If Not sender.text = "Create" Then MsgBox("APEX Operation Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
            Initial.CurrentOption = 24
            Me._Create_4.Enabled = True
            '*****************************
            Call UpdateEnvironmentVariables()
            enable_Menu()
            Wait_Form.Close()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub rbBtn_SubareaFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnSubarea.Click
        Dim d As StreamWriter
        On Error GoTo goError
        If ValidateLandUses() Then
            Exit Sub
        End If

        'Create Control files
        Initial.subareafile = 2

        Wait_Form.Label1(3).Text = "The APEX Subarea Files are Being Created"
        Wait_Form.Show()
        Wait_Form.Pbar_Scenarios.Value = 25
        Wait_Form.Refresh()
        d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.suba))
        d.Close()
        d.Dispose()
        d = Nothing
        Create_Files_Form.Tag = "2"
        Sitefiles.SiteFiles(4)
        Wait_Form.Pbar_Scenarios.Value = 75
        If Not sender.text = "Create" Then MsgBox("APEX Subarea Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 25
        Call createAPEXBat()
        Wait_Form.Pbar_Scenarios.Value = 100
        Me._Create_5.Enabled = True
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Private Sub rbBtn_SoilFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnSoil.Click
        Dim d As StreamWriter

        On Error GoTo goError
        If ValidateLandUses() Then
            Exit Sub
        End If

        'Create Control files
        Wait_Form.Label1(3).Text = "The APEX Soil Files are Being Created"
        Wait_Form.Show()
        Wait_Form.Pbar_Scenarios.Value = 25
        Wait_Form.Refresh()
        d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Soil))
        d.Close()
        d.Dispose()
        d = Nothing
        Create_Files_Form.Tag = "3"
        Sitefiles.SiteFiles(5)
        Wait_Form.Pbar_Scenarios.Value = 75
        If Not sender.text = "Create" Then MsgBox("APEX Soil Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 26
        Me._Create_6.Enabled = True
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        Wait_Form.Pbar_Scenarios.Value = 100
        Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Private Sub rbBtn_SiteFile_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnSite.Click
        Dim d As StreamWriter

        On Error GoTo goError
        If ValidateLandUses() Then
            Exit Sub
        End If

        'Create Control files
        Wait_Form.Label1(3).Text = "The APEX Site Files are Being Created"
        Wait_Form.Show()
        Wait_Form.Pbar_Scenarios.Value = 25
        Wait_Form.Refresh()
        d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.Site))
        d.Close()
        d.Dispose()
        d = Nothing
        Create_Files_Form.Tag = "5"
        Sitefiles.SiteFiles(6)
        Wait_Form.Pbar_Scenarios.Value = 75
        If Not sender.text = "Create" Then MsgBox("APEX Site Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 27
        Me._Create_7.Enabled = True
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        Wait_Form.Pbar_Scenarios.Value = 100
        Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Private Sub rbBtn_WeatherFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnWeather.Click
        On Error GoTo goError
        If ValidateLandUses() Then
            Exit Sub
        End If

        'Create Control files
        'Wait_Form.Label1(3).Text = "The APEX Weather Files are Being Created"
        'Wait_Form.Show()
        'Wait_Form.Refresh()
        'Create_Files_Form.Tag = "6"
        Call Control.Apexcont(1)
        'Wait_Form.Pbar_Scenarios.Value = 0
        If CDbl(Create_Files_Form.pcpgages) <> 0 Then
            If Initial.Version = "4.0.0" Or Initial.Version = "4.1.0" Or Initial.Version = "4.2.0" Or Initial.Version = "4.3.0" _
                Or Initial.Version = "1.1.0" Or Initial.Version = "1.2.0" Or Initial.Version = "1.3.0" Then
                Dataclass.Weather1()
            Else
                Sitefiles.SiteFiles(7)
            End If
            'Dataclass.Weather1
        End If
        If Not sender.text = "Create" Then MsgBox("APEX Weather Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 28
        'Me._Create_8.Enabled = True
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        '   Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Private Sub rbBtn_WPMFiles_Click(sender As System.Object, e As System.EventArgs) Handles rbBtnWeatherStation.Click
        Dim d As StreamWriter

        On Error GoTo goError

        If ValidateLandUses() Then
            Exit Sub
        End If
        'Create Control files
        'Wait_Form.Label1(2).Text = "The APEX .wpm Files are Being Created"
        'Wait_Form.Show()
        'Wait_Form.Refresh()

        'fs = CreateObject("Scripting.FileSystemObject")
        d = New StreamWriter(File.OpenWrite(Initial.Output_files & "\" & Initial.wpm1))
        d.Close()
        d.Dispose()
        d = Nothing
        'Me.Refresh()
        'Create_Files_Form.Tag = "7"
        Sitefiles.SiteFiles(8)
        'Wait_Form.Check1(2).CheckState = System.Windows.Forms.CheckState.Checked
        'Wait_Form.Check1(2).Visible = True
        'Wait_Form.Show()
        'Wait_Form.Label1(4).Text = "The SWAPP General Files are Being Copied"
        'Wait_Form.Check1(3).CheckState = System.Windows.Forms.CheckState.Checked
        'Wait_Form.Check1(3).Visible = True
        'Wait_Form.Refresh()
        Call cpyApexSwat()
        'Wait_Form.Label1(5).Text = "The SWAT General Files are Being Copied"
        'Wait_Form.Check1(4).CheckState = System.Windows.Forms.CheckState.Checked
        'Wait_Form.Check1(4).Visible = True
        'Wait_Form.Refresh()
        Call cpySwat()
        If Not sender.text = "Create" Then MsgBox("APEX .wpm Files Generated", MsgBoxStyle.OkOnly, "APEX Files Generation")
        Initial.CurrentOption = 30
        '*****************************
        Call UpdateEnvironmentVariables()
        enable_Menu()
        'Wait_Form.Close()
        Exit Sub
goError:
        MsgBox(Err.Description)

    End Sub

    Public Function ValidateSubasinsSelected() As UShort
        Dim records As UInteger = 0
        Dim dt As New DataTable

        'records = GetTableRecords("SELECT Subbasin FROM Runs", Initial.Output_files)
        dt = getLocalDataTable("SELECT Subbasin FROM Runs", Initial.Output_files)
        'Dim Sw_Bat As String = String.Empty
        'Dim bat_file As String = Dir(Initial.Output_files & "\APEXBat.txt")

        'If bat_file = "" Then
        '    MsgBox("No subbasins selected to run - Select subbasins and try again", vbOKOnly, "Select Subbasins")
        '    Return 1
        'End If

        Return dt.Rows.Count
    End Function

    Public Function ValidateLandUses() As UShort
        Dim Sw_Bat As String = String.Empty
        Dim Exclude As DataTable

        Exclude = New DataTable
        Exclude = getDBDataTable("SELECT * FROM exclude WHERE folder=" & "'" & Initial.Dir1 & "'" & " AND project = " & " '" & Initial.Project & "'")
        If Exclude.Rows.Count <= 0 = True Then
            MsgBox("No Land-Uses Selected, Please Select Land-Uses to Simulate in APEX and Try Again")
            Return 1
        End If

        Exclude.Dispose()
        Exclude = Nothing

        SelectVersion()
        Return 0
    End Function

    Public Sub SelectVersion()
        'Dim Swat_bat As String
        'Dim Sw_bat As String

        Select Case Initial.Version
            Case "1.0.0"    'versions 1.x.x were changed from APEX0806 to APEX0806 (Last one). 1/14/2015
                APEX_bat = "Apex0806.bat"
                Sw_bat = "Sw0604_2000.bat"
                Swat_bat = "Swat2000.bat"
            Case "1.1.0"
                APEX_bat = "Apex0806.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2003.bat"
            Case "1.2.0"
                APEX_bat = "Apex0806.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2009.bat"
                '<New version of SWAT_2012 is added
            Case "1.3.0"
                APEX_bat = "Apex0806.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2012.bat"
                '4/16/2013>
            Case "2.0.0"
                APEX_bat = "Apex2110.bat"
                Sw_bat = "Sw2110_2000.bat"
                Swat_bat = "Swat2000.bat"
            Case "2.1.0"
                APEX_bat = "Apex2110.bat"
                Sw_bat = "Sw2110_2003.bat"
                Swat_bat = "Swat2003.bat"
            Case "3.0.0"
                APEX_bat = "Epic3060.bat"
            Case "3.1.0"
                APEX_bat = "Epic3060.bat"
            Case "4.0.0"
                APEX_bat = "Apex0604.bat"
                Sw_bat = "Sw0604_2000.bat"
                Swat_bat = "Swat2000.bat"
            Case "4.1.0"
                APEX_bat = "Apex0604.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2003.bat"
            Case "4.2.0"
                APEX_bat = "Apex0604.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2009.bat"
                '<New version of SWAT_2012 is added
            Case "1.3.0"
                APEX_bat = "Apex2110.bat"
                Sw_bat = "Sw2110_2003.bat"
                Swat_bat = "Swat2012.bat"
            Case "2.3.0"
                APEX_bat = "Apex2110.bat"
                Sw_bat = "Sw2110_2003.bat"
                Swat_bat = "Swat2012.bat"
            Case "4.3.0"
                APEX_bat = "Apex0604.bat"
                Sw_bat = "Sw0604_2003.bat"
                Swat_bat = "Swat2012.bat"
                '4/16/2013>
        End Select

        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()
    End Sub

    Private Function run_apex() As String
        Dim srFile As StreamReader = Nothing
        Dim swFile As StreamWriter = Nothing
        Dim myval As Integer = 0
        Dim direxe As String = Initial.Output_files
        Dim i, j As UShort
        Dim subbasinNumber As String = String.Empty

        With Wait_Form
            .Label1(1).Text = "The APEX Program is Being Executed"
            .Show()
            srFile = New StreamReader(File.OpenRead(Initial.Output_files & "\APEXBat.txt"))
            j = 1
            Do While srFile.EndOfStream <> True
                swFile = New StreamWriter(File.Create(Initial.Output_files & "\" & APEX_bat))
                For i = 1 To 5
                    swFile.WriteLine(srFile.ReadLine)
                Next
                swFile.Close()
                swFile.Dispose()
                swFile = Nothing
                'update apexcont and parms files if there is manure application based on N or P.
                'Update_Params(j)
                j += 1
                'run_process(Initial.Output_files & "\" & APEX_bat)
                myval = ExecCmd(Initial.Output_files & "\" & APEX_bat, direxe)

                If myval <> 0 Then
                    Return "Subbasins " & subbasinNumber & " Has Problems - Check it out and try again"
                Else

                End If
            Loop
        End With

        If Not swFile Is Nothing Then
            swFile.Close()
            swFile.Dispose()
            swFile = Nothing
        End If
        If Not srFile Is Nothing Then
            srFile.Close()
            srFile.Dispose()
            srFile = Nothing
        End If

        Return "OK"

    End Function

    'Private Sub MA_Parm(input8 As UShort)
    '    Dim temp As String = String.Empty
    '    Dim i As UShort = 1
    '    Dim sr_file As StreamReader = New StreamReader(Initial.Output_files + "\" + Initial.parm)
    '    Dim sw_file As StreamWriter = New StreamWriter(Initial.Output_files + "\bk" + Initial.parm)
    '    Try
    '        'modify parms.dat changing parm(43)
    '        'read the first 34 lines. Parm 43 is in line 35
    '        Do While i < 35
    '            sw_file.WriteLine(sr_file.ReadLine())
    '            i += 1
    '        Loop
    '        temp = sr_file.ReadLine()
    '        Mid(temp, 17, 8) = Format(input8 * in_to_m, "##0.0000").PadLeft(8)  '(1) Line 2 field 1 (MNUL)
    '        sw_file.WriteLine(temp)
    '        Do While sr_file.EndOfStream <> True
    '            sw_file.WriteLine(sr_file.ReadLine)
    '        Loop

    '        Close_Stream_File(sr_file)
    '        Close_Stream_File(sw_file)

    '        System.IO.File.Copy(Initial.Output_files + "\bk" + Initial.parm, Initial.Output_files + "\\" + Initial.parm, True)
    '        System.IO.File.Delete(Initial.Output_files + "\bk" + Initial.parm)
    '    Catch ex As Exception
    '    Finally
    '        Close_Stream_File(sr_file)
    '        Close_Stream_File(sw_file)
    '    End Try

    'End Sub

    'Private Sub Update_Params(j As UShort)
    '    Dim runs As DataTable = Nothing
    '    Dim bmps As DataTable = Nothing

    '    Try
    '        'find the subbasin being simulated
    '        runs = getLocalDataTable("SELECT TOP " & j & " * FROM runs", Initial.Output_files)
    '        'find the paramas for manure application if any
    '        bmps = getLocalDataTable("SELECT * FROM bmps INNER JOIN bmp_subareas ON bmps.bmp_name = bmp_subareas.bmp WHERE bmp_subareas.subarea = '" & runs.Rows(runs.Rows.Count - 1).Item(0) & "' AND bmps.BMP LIKE 'Manure Application%'", Initial.Output_files)
    '        If bmps.Rows.Count > 0 Then
    '            MA_Control(bmps.Rows(0).Item("Input5"), bmps.Rows(0).Item("Input4"))
    '            MA_Parm(bmps.Rows(0).Item("Input8"))
    '        Else
    '            If Not Initial.Scenario.Contains("Baseline") Then
    '                File.Copy(Initial.Dir1 & "\APEX\Apexcont.dat", Initial.Output_files & "\Apexcont.dat", True)
    '                File.Copy(Initial.Dir1 & "\APEX\" & Initial.parm, Initial.Output_files & "\" & Initial.parm, True)
    '            End If
    '        End If
    '    Catch ex As Exception

    '    End Try
    'End Sub

    'Private Sub MA_Control(input5 As UShort, input4 As Single)
    '    'Dim file As String = "apexcont.dat"
    '    Dim temp As String = String.Empty
    '    Dim sr_file As StreamReader = New StreamReader(Initial.Output_files + "\" + Initial.cont)
    '    Dim sw_file As StreamWriter = New StreamWriter(Initial.Output_files + "\bk" + Initial.cont)
    '    Try
    '        'modify apexcont adding the upper limit and changing the MNUL code
    '        sw_file.WriteLine(sr_file.ReadLine())
    '        temp = sr_file.ReadLine()
    '        Mid(temp, 1, 4) = Format(input5, "###0").PadLeft(4)  '(1) Line 2 field 1 (MNUL)
    '        sw_file.WriteLine(temp)
    '        sw_file.WriteLine(sr_file.ReadLine)
    '        sw_file.WriteLine(sr_file.ReadLine)
    '        sw_file.WriteLine(sr_file.ReadLine)
    '        'modify apexcont adding the upper the N and P application rate. Add the same because just one is able to be running at a time.
    '        temp = sr_file.ReadLine
    '        Mid(temp, 25, 32) = Format(input4, "####0.00").PadLeft(8) & "   62.00   120.0" & Format(input4, "####0.00").PadLeft(8)
    '        sw_file.WriteLine(temp)
    '        Do While sr_file.EndOfStream <> True
    '            sw_file.WriteLine(sr_file.ReadLine)
    '        Loop
    '        Close_Stream_File(sr_file)
    '        Close_Stream_File(sw_file)

    '        System.IO.File.Copy(Initial.Output_files + "\bk" + Initial.cont, Initial.Dir2 + "\" + Initial.cont, True)
    '        System.IO.File.Delete(Initial.Output_files + "\bk" + Initial.cont)
    '    Catch ex As Exception
    '    Finally
    '        Close_Stream_File(sr_file)
    '        Close_Stream_File(sw_file)
    '    End Try

    'End Sub

    Private Sub Add_NUtrients(subarea As String)
        Dim srFile As StreamReader = Nothing
        Dim swFile As StreamWriter = Nothing
        Dim point_source As String = "APEX" & subarea.Substring(1, 4) & "P.dat"
        Dim temp() As String
        Dim i As UShort = 0
        Dim value As Single = 0

        Try
            srFile = New StreamReader(File.OpenRead(Initial.Swat_Output & "\" & point_source))
            swFile = New StreamWriter(File.Create(Initial.Swat_Output & "\bk_" & point_source))
            For i = 0 To 5
                swFile.WriteLine(srFile.ReadLine)
            Next
            Do While srFile.EndOfStream <> True
                temp = srFile.ReadLine.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                If temp(0) = "AVEAN" Or temp(0) = "" Then Exit Do
                temp(4) += Variables.ynSF
                temp(5) += Variables.ypSF
                temp(6) += Variables.qnSF
                temp(9) += Variables.qpSF
                'Write the new values in the point source file
                For i = 0 To temp.Length - 1
                    Select Case i
                        Case 0, 1
                            swFile.Write(CShort(temp(i)).ToString("####0").PadLeft(5))
                        Case 2
                            swFile.Write(Format(CSng(temp(i)), "0.00000E+00").PadLeft(13))
                        Case Is > 2
                            swFile.Write(Format(CSng(temp(i)), "0.00000E+00").PadLeft(12))
                    End Select
                Next
                swFile.WriteLine()
            Loop
            srFile.Close()
            srFile.Dispose()
            srFile = Nothing
            swFile.Close()
            swFile.Dispose()
            swFile = Nothing
            'replace source point file with the new one.
            File.Copy(Initial.Swat_Output & "\" & point_source, Initial.Swat_Output & "\org_" & point_source, True)
            File.Copy(Initial.Swat_Output & "\bk_" & point_source, Initial.Swat_Output & "\" & point_source, True)

        Catch ex As Exception

        End Try
    End Sub

    Private Sub SFNutrients(subarea As String)
        Dim totalManure As Single = 0
        Dim conConnection As New ADODB.Connection
        Dim cmdCommand As New ADODB.Command
        Dim rstRecordSet As New ADODB.Recordset
        Dim myConnection As New OleDb.OleDbConnection
        Dim animal_code As UShort = 0
        Dim animals_stream As DataTable
        Dim animals_row As DataRow
        Dim dry_manure As Single = 0
        Dim bmp As DataTable

        Try
            Variables.qnSF = 0  'no3
            Variables.qpSF = 0  'po4
            Variables.ynSF = 0  'org n
            Variables.ypSF = 0  'org p

            animals_stream = getDBDataTable("SELECT * FROM AnimalsStream WHERE Subarea = '" & subarea & "'")
            If Not animals_stream Is Nothing Then
                myConnection.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & SWAPP.Initial.Dir1 & "\Project_Parameters.mdb;"
                If myConnection.State = Data.ConnectionState.Closed Then
                    myConnection.Open()
                End If
                For Each animals_row In animals_stream.Rows
                    'get all of the fields with animals in stream for the subbasin
                    bmp = getLocalDataTable("SELECT TOP 1 *, BMPs.Other AS BMP_Other FROM BMPs, BMP_Subareas WHERE BMPs.BMP_NAME=BMP_Subareas.BMP AND Subarea='" & subarea & "' AND HRU='" & animals_row.Item("field_id") & "' AND Scenario='" & Initial.Scenario & "' AND BMPs.BMP='Stream Fencing'", Initial.Output_files)
                    If bmp.Rows.Count = 0 Then
                        'if stream fencing bp found calculate nutrient loading to add in daily.
                        'totalManure = bmp.Item("Input1") * bmp.Item("Input3") / 24 * bmp.Item("Input4")
                        dry_manure = getData("SELECT dry_manure FROM Animals WHERE Code=" & animals_row.Item("animal_type").ToString & "", myConnection)
                        totalManure = animals_row.Item("animals") * animals_row.Item("hours") / 24 * dry_manure
                        'Step1: read “Fertilizer” table in “Project_Parameters.mdb”
                        conConnection.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source= " & Initial.Dir1 & "\Project_Parameters.mdb;"
                        conConnection.Open()
                        With cmdCommand
                            .ActiveConnection = conConnection
                            .CommandText = "SELECT Min_N,Min_P, Org_N,Org_P FROM Fertilizer WHERE Code=" & animals_row.Item("animal_type") & ";"
                        End With
                        rstRecordSet.Open(cmdCommand)
                        Variables.qnSF += rstRecordSet(0).Value * totalManure * animals_row.Item("days") / 365  'no3
                        Variables.qpSF += rstRecordSet(1).Value * totalManure * animals_row.Item("days") / 365  'po4
                        Variables.ynSF += rstRecordSet(2).Value * totalManure * animals_row.Item("days") / 365  'org n
                        Variables.ypSF += rstRecordSet(3).Value * totalManure * animals_row.Item("days") / 365  'org p
                        conConnection.Close()
                    End If
                Next
                myConnection.Close()
                myConnection.Dispose()
                myConnection = Nothing
            End If

        Catch ex As Exception
            MsgBox(ex.Message, , "Subroutine = SFNutrients")
        End Try
    End Sub

    Private Function run_swat() As String
        Dim name_Renamed As String
        Dim SWATFile As String
        Dim myval As UShort = 0
        Dim bmp_subareas As DataTable
        Dim bmp As DataRow

        With Wait
            .Label1(1).Text = "The SWAT Simulation has started"
            .Show()
            'name_Renamed = "*.SWT"
            ' Copy SWAT files from apex to swat
            'Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "*.SWT", True)  ' no need because the program will take the .swt file from apex folder
            'name_Renamed = "SW*"
            'On Error Resume Next
            'Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "*.SW", True)
            'On Error GoTo 0
            SWATFile = Dir(Initial.OrgDir & "\SWAT20*.exe")
            'copy the executable SWAT files, but do not replace the original in the SWAT folder.
            'If SWATFile <> "" Then Dataclass.copyAllFiles(Initial.OrgDir, Initial.Swat_Output, "SWAT20*.exe", False)   'SWAT folder should havethe original SWAT program
            .Label1(3).Text = "The SWAT to APEX Interface is Being Executed"
            .Check1(1).CheckState = System.Windows.Forms.CheckState.Checked
            .Check1(1).Visible = True
            .Refresh()

            'name_Renamed = "SW*.bat"
            'Dataclass.copyAllFiles(Initial.OrgDir, Initial.Swat_Output, "sw*.bat", True)
            'Dataclass.copyAllFiles(Initial.OrgDir, Initial.Swat_Output, "sw*.exe", False)
            'If Initial.Version = "1.3.0" Or Initial.Version = "4.3.0" Then
            '    transfer_swat_files()
            'Else
            '    myval = ExecCmd(Initial.Swat_Output & "\" & Sw_bat, Initial.Swat_Output)
            'End If

            'If myval <> 0 Then
            '    Return "Error transfer SWAT files from APEX to SWAT"
            'End If

            transfer_swat_files()

            .Label1(5).Text = "The Fig file is being recreated"
            .Check1(3).CheckState = System.Windows.Forms.CheckState.Checked
            .Check1(3).Visible = True
            .Refresh()
            'save all of the areas to add in the fig file in SWAT
            Call figAreas()
            'add input source files from APEX to fig file in SWAT
            Call Copyfig()
            name_Renamed = "SWat.bat"
            .Label1(7).Text = "Preparing the files in the SWAT folder"
            .Check1(5).CheckState = System.Windows.Forms.CheckState.Checked
            .Check1(5).Visible = True
            .Refresh()
            On Error Resume Next
            Dataclass.copyAllFiles(Initial.Output_files, Initial.Swat_Output, "\swat*.bat", True)
            On Error GoTo 0
            name_Renamed = "*.*"
            Dataclass.copyAllFiles(Initial.New_Swat, Initial.Swat_Output, "*.*", True)
            .Label1(1).Text = "The SWAT Program is Being Executed"
            .Check1(7).CheckState = System.Windows.Forms.CheckState.Checked
            .Check1(7).Visible = True
            .Refresh()

            'check if file.cio is printing the rch file by month, year, or day and change it to day if it is not.
            changeFileCio()
            'check if there are Stream Fencing BMPs and add the values for nutrients into the point source file for the subarea with this BMP.
            'bmp_subareas = getLocalDataTable("SELECT DISTINCT BMP_Subareas.Subarea FROM BMPs, BMP_Subareas WHERE BMPs.BMP_NAME=BMP_Subareas.BMP AND Scenario='" & Initial.Scenario & "' AND BMPs.BMP='Stream Fencing'", Initial.Output_files)
            'if there are naimal in stream the nutrient deposited in the stream are calculated here.
            bmp_subareas = getDBDataTable("SELECT DISTINCT Subarea FROM AnimalsStream")
            If Not bmp_subareas Is Nothing Then
                For Each bmp In bmp_subareas.Rows
                    'calculate nutrients 
                    SFNutrients(bmp.Item("Subarea"))
                    'Read the point source file and add the new nutrients calculated
                    Add_NUtrients(bmp.Item("Subarea"))
                Next
            End If
            myval = ExecCmd(Initial.Swat_Output & "\" & Swat_bat, Initial.Swat_Output)
            If myval <> 0 Then
                Return "Error running SWAT"
            End If

            '.Label1(9).Text = "Saving SWAT results"
            '.Check1(1).CheckState = System.Windows.Forms.CheckState.Checked
            '.Check1(1).Visible = True
            '.Refresh()
            Reach.ReadSaveFile()
            .Close()
        End With
        'Mensaje = "APEX and SWAT Programs within SWAPP Were Successfully Executed"
        modifyRecords("UPDATE paths SET LastRun='APEX' WHERE Scenario='" & Initial.Scenario & "'")
        Return "OK"
    End Function

    Private Function run_FEM() As String
        Dim msg As String = String.Empty
        'With Wait
        '.Label1(7).Text = "The FEM Program is Being Executed"
        'msg = Execute_FEM_Click()
        '.Check1(7).CheckState = System.Windows.Forms.CheckState.Checked
        '.Check1(7).Visible = True
        '.Refresh()
        'End With

        Return msg
    End Function

    Private Sub rbBtnSubarea_DoubleClick(sender As Object, e As EventArgs) Handles rbBtnSubarea.DoubleClick

    End Sub

    Private Sub rbBnt_ExecuteAll_Click1(sender As System.Object, e As System.EventArgs) Handles rbBnt_ExecuteAll.Click, rbBnt_APEX_SWAT.Click, rbBnt_SWAT.Click, rbBnt_FEM.Click, rbBnt_APEX.Click
        Dim index As Short = sender.Value
        Dim ans As Integer
        Dim temp As Object
        Dim i As Integer
        Dim Check_Local As String
        Dim adoRec As DataTable
        Dim bat_file, files As String
        Dim tempDT As DataTable
        Dim currdir As String
        Dim a As Short
        Dim msg As String = String.Empty
        'On Error GoTo goError

        'direxe = Initial.Output_files

        SelectVersion()
        If ValidateSubasinsSelected() <= 0 Then
            MsgBox("No subbasins selected to run - Select subbasins and try to run", vbOKOnly, "Select Subbasins")
            Exit Sub
        End If

        If Not IsAvailable(Wait_Form) Then Wait_Form = New Wait()

        With Wait_Form
            .Pbar_Scenarios.Visible = False
            Select Case index
                Case 3 '3: Execute APEX / SWAT / FEM  5:Execute APEX / SWAT
                    'first Execute APEX
                    msg = run_apex()
                    If msg = "OK" Then 'run SWAT
                        msg = run_swat()
                    End If

                    If msg = "OK" Then 'run SWAT
                        msg = run_FEM()
                    End If

                    If msg = "OK" Then
                        MsgBox("APEX, SWAT, and FEM ran successfully", , "Confirmation")
                    End If
                    Initial.CurrentOption = 34
                Case 5 '3: Execute APEX / SWAT / FEM  5:Execute APEX / SWAT
                    'first Execute APEX
                    msg = run_apex()
                    If msg = "OK" Then 'run SWAT
                        msg = run_swat()
                    End If

                    If msg = "OK" Then
                        MsgBox("APEX, and SWAT ran successfully", , "Confirmation")
                    End If
                    Initial.CurrentOption = 34
                Case 4 'Execute &APEX
                    'first Execute APEX
                    msg = run_apex()
                    If msg = "OK" Then
                        MsgBox("APEX ran successfully", , "Confirmation")
                    End If
                    Initial.CurrentOption = 34
                Case 9 'Execute &SWAT
                    msg = run_swat()
                    If msg = "OK" Then
                        MsgBox("SWAT ran successfully", , "Confirmation")
                    End If
                    Initial.CurrentOption = 35
                Case 10
                    run_FEM()
                    If msg = "OK" Then
                        MsgBox("FEM ran successfully", , "Confirmation")
                    End If
                    Initial.CurrentOption = 40
            End Select
            .Pbar_Scenarios.Visible = True
        End With

        Initial.CurrentOption = 40
        Call UpdateEnvironmentVariables()
        On Error Resume Next
        Exit Sub

        'goError:
        'MsgBox(Err.Description, , "Execute_Apex - " & name_Renamed)
    End Sub

    Private Sub transfer_swat_files()
        Dim dt As DataTable
        Dim dr As DataRow
        Dim swat_file As String = String.Empty
        Dim point_source As String = String.Empty
        Dim sr As StreamReader = Nothing
        Dim sw As StreamWriter = Nothing
        Dim swat_columns As String()
        Dim i As UShort = 0
        Dim j As UShort = 0
        Dim swat_string As String = String.Empty

        Try
            'read runs table fro local db to take the subbasins simuLated in APEX
            dt = getLocalDataTable("SELECT * FROM Runs", Initial.Output_files)
            For Each dr In dt.Rows
                i = 0
                swat_file = dr.Item("Subbasin").ToString.Substring(1, 8) & ".swt"
                sr = New StreamReader(Path.Combine(Initial.Output_files, swat_file))
                point_source = "APEX" & swat_file.Substring(0, 4) & "P.dat"
                sw = New StreamWriter(Path.Combine(Initial.Swat_Output, point_source))
                'sr.ReadLine()
                sw.WriteLine()
                'sr.ReadLine()
                sw.WriteLine(point_source)
                sw.WriteLine()
                sw.WriteLine()
                sw.WriteLine()
                sw.WriteLine(point_source)
                'sr.ReadLine()
                Do While sr.EndOfStream <> True
                    swat_columns = sr.ReadLine().Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                    If swat_columns.Length = 0 Then
                        Continue Do
                    Else
                        If swat_columns(0).Trim <> "1" And i = 0 Then Continue Do
                        If swat_columns(0).Contains("AVEAN") Then Exit Do
                    End If
                    If i = 0 Then
                        i += 1
                        swat_string = Format(0, "##0").PadLeft(5) & Format(0, "##0").PadLeft(5)   'write day and year as zeros
                    Else
                        swat_string = Format(CSng(swat_columns(0)), "##0").PadLeft(5) & Format(CSng(swat_columns(1)), "##0").PadLeft(5)  'write day and year 
                    End If
                    swat_string &= " "
                    For j = 2 To 6
                        If swat_columns(j) < 0 Then
                            swat_columns(j) = 0
                        End If
                        swat_string &= Format(CSng(swat_columns(j)), "0.00000E+00").PadLeft(12)  'print Flow, Sed, Org N, Org P, and NO3
                    Next
                    swat_string &= Format(0, "0.00000E+00").PadLeft(12)   'print NH3 in zero
                    swat_string &= Format(0, "0.00000E+00").PadLeft(12)     'print NO2 in zero
                    swat_string &= Format(CSng(swat_columns(j)), "0.00000E+00").PadLeft(12)     'print PO4
                    For j = 1 To 10
                        swat_string &= Format(0, "0.00000E+00").PadLeft(12)  'print 10 more values en zero to have the whole row.
                    Next
                    'swat_string &= Format(0, "0.00000E+00").PadLeft(12)
                    sw.WriteLine(swat_string)
                Loop
                'Mid(swat_string, 1, 5) = Format(0, "##0").PadLeft(5)
                'Mid(swat_string, 6, 5) = Format(0, "##0").PadLeft(5)
                'sw.WriteLine(swat_string)
                sw.Flush()
                If Not sr Is Nothing Then
                    sr.Close()
                    sr.Dispose()
                    sr = Nothing
                End If
                If Not sw Is Nothing Then
                    sw.Close()
                    sw.Dispose()
                    sw = Nothing
                End If
            Next

        Catch ex As Exception
        Finally
            If Not sr Is Nothing Then
                sr.Close()
                sr.Dispose()
                sr = Nothing
            End If
            If Not sw Is Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If

        End Try
    End Sub

    Private Sub rbBtnAnimalStream_Click(sender As Object, e As System.EventArgs) Handles rbBtnAnimalStream.Click
        If Not IsAvailable(Animals_in_Stream_Form) Then Animals_in_Stream_Form = New Animals_in_Stream
        Animals_in_Stream_Form.ShowDialog()  'Show the land use window to select them
    End Sub
End Class


