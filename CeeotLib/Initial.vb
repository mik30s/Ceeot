Option Strict Off
Option Explicit On

Module Initial
    Private mvarcont, mvarparm, mvaropcs, mvarsoil, mvarsuba, mvarsite, mvarwpm1, mvarsublines, mvarfertilizer As Object
    Private mvarHerd As String
    Private mvarPesticide, mvarPest, mvarOp As Object
    Private mvartill As String
    Private mvarVersion2, mvarVersion1, mvarVersion, mvarInput_Files, MvarDir_Bas As Object
    Private MvarDir_Sce As String
    Private mvardir2, mvarDir1, mvardeleted As Object
    Private mvarOrgDir, mvarProject, mvarEspace, mvarEspace1, mvarScenario, mvarBas As Object
    Private mvarSce As String
    Private mvarNew_Swat, mvarSWAT_Output, mvarOutput_files, mvarFEM, MvarDir_BasF As Object
    Private MvarDir_SceF As String
    Private mvarcntrl3, mvarcntrl1, mvarcntrl, mvarcntrl2, mvarcntrl4 As Object
    Private mvarRunFile As String
    Private mvarcancel1, mvarcol4, mvarcol2, mvarcol1, mvarlimit, mvarcol3, mvarcol5, MvarFEMRes As Object
    Private MvarNumProj As Short
    Private mvarErrors, mvarcol7, mvarcol8, mvarngn, mvarYearSim, mvarcol6, MvarSubaLines, mvarFile_Number As Object
    Private mvarLayers As Short
    Private mvarlat1(50) As Object
    Private mvarlon1(50) As Object
    Private mvarChemical As Object
    Private mvarOperation1 As Object
    Private mvarStrSelect As String
    Private mvarSubChanged, mvarFEMChecked As Object
    Private mvarPestChanged(3) As Object
    Private mvarOpChanged(11) As Boolean
    Private mvarCurrentOption As Byte
    Private mvarMaxLen, mvarMsgBox_Answer, mvarOpcsCode As Object
    Private mvarCodeSelected As Short
    Private mvarMeasuredFile As String
    Private mvarHerds As Object
    Private mvarOwners As Short
    Private mvarFig As String
    Private mvarSubareaFile As Integer
    Private mvarprpfiles(18) As String
    Private mvartemfiles(18) As String
    Private mvarslrfiles As String
    Private mvarhmdfiles As String
    Private mvarwndfiles As String
    Private mvarSoilName, mvarLandUse As String
    Private mvaryear_col As UShort
    Private mvarScenario_type As String
    Private mvarBMP As String

    Public Structure ValueModified ' Create user-defined to identified values changed
        Dim modified As Boolean ' Define if value was modified.
        Dim new_Renamed As String ' Define new value
        Dim format_Renamed As String ' Use to determine the lenght of the field in the file
        Dim position As Short ' Position of the field within the file.
        Dim Line As Short ' Line in wich the field is located within the file.
    End Structure

    Public FEMdbChoice As Object

    Public Property prpfiles(ByVal i As Short) As String
        Get
            prpfiles = mvarprpfiles(i)
        End Get
        Set(ByVal Value As String)
            mvarprpfiles(i) = Value
        End Set
    End Property
    Public Property temfiles(ByVal i As Short) As String
        Get
            temfiles = mvartemfiles(i)
        End Get
        Set(ByVal Value As String)
            mvartemfiles(i) = Value
        End Set
    End Property
    Public Property slrfiles() As String
        Get
            slrfiles = mvarslrfiles
        End Get
        Set(ByVal Value As String)
            mvarslrfiles = Value
        End Set
    End Property
    Public Property hmdfiles() As String
        Get
            hmdfiles = mvarhmdfiles
        End Get
        Set(ByVal Value As String)
            mvarhmdfiles = Value
        End Set
    End Property
    Public Property wndfiles() As String
        Get
            wndfiles = mvarwndfiles
        End Get
        Set(ByVal Value As String)
            mvarwndfiles = Value
        End Set
    End Property

    Public Property subareafile() As Integer
        Get
            subareafile = Trim(mvarSubareaFile)
        End Get
        Set(ByVal Value As Integer)
            mvarSubareaFile = Value
        End Set
    End Property

    Public Property figsFile() As String
        Get
            figsFile = Trim(mvarFig)
        End Get
        Set(ByVal Value As String)
            mvarFig = Value
        End Set
    End Property
    Public Property owners() As Short
        Get
            owners = mvarOwners
        End Get
        Set(ByVal Value As Short)
            mvarOwners = Value
        End Set
    End Property
    Public Property herds() As Short
        Get
            herds = mvarHerds
        End Get
        Set(ByVal Value As Short)
            mvarHerds = Value
        End Set
    End Property
    Public Property measuredfile() As String
        Get
            measuredfile = mvarMeasuredFile
        End Get
        Set(ByVal Value As String)
            mvarMeasuredFile = Value
        End Set
    End Property
    Public Property OpChanged(ByVal i As Short) As Boolean
        Get
            OpChanged = mvarOpChanged(i)
        End Get
        Set(ByVal Value As Boolean)
            mvarOpChanged(i) = Value
        End Set
    End Property
    Public Property PestChanged(ByVal i As Short) As Boolean
        Get
            PestChanged = mvarPestChanged(i)
        End Get
        Set(ByVal Value As Boolean)
            mvarPestChanged(i) = Value
        End Set
    End Property
    Public Property FEMChecked() As Boolean
        Get
            FEMChecked = mvarFEMChecked
        End Get
        Set(ByVal Value As Boolean)
            mvarFEMChecked = Value
        End Set
    End Property
    Public Property strSelect() As String
        Get
            strSelect = mvarStrSelect
        End Get
        Set(ByVal Value As String)
            mvarStrSelect = Value
        End Set
    End Property
    Public Property Operation1() As String
        Get
            Operation1 = mvarOperation1
        End Get
        Set(ByVal Value As String)
            mvarOperation1 = Value
        End Set
    End Property
    Public Property Chemical() As String
        Get
            Chemical = mvarChemical
        End Get
        Set(ByVal Value As String)
            mvarChemical = Value
        End Set
    End Property
    Public Property Op() As String
        Get
            Op = mvarOp
        End Get
        Set(ByVal Value As String)
            mvarOp = Value
        End Set
    End Property
    Public Property Pest() As String
        Get
            Pest = mvarPest
        End Get
        Set(ByVal Value As String)
            mvarPest = Value
        End Set
    End Property
    Public Property Pesticide() As String
        Get
            Pesticide = mvarPesticide
        End Get
        Set(ByVal Value As String)
            mvarPesticide = Value
        End Set
    End Property
    Public Property SubChanged() As Boolean
        Get
            SubChanged = mvarSubChanged
        End Get
        Set(ByVal Value As Boolean)
            mvarSubChanged = Value
        End Set
    End Property
    Public Property CurrentOption() As Byte
        Get
            CurrentOption = mvarCurrentOption
        End Get
        Set(ByVal Value As Byte)
            mvarCurrentOption = Value
        End Set
    End Property
    Public Property lat1(ByVal i As Short) As String
        Get
            lat1 = mvarlat1(i)
        End Get
        Set(ByVal Value As String)
            mvarlat1(i) = Value
        End Set
    End Property
    Public Property lon1(ByVal i As Short) As String
        Get
            lon1 = mvarlon1(i)
        End Get
        Set(ByVal Value As String)
            mvarlon1(i) = Value
        End Set
    End Property

    Public Property codeSelected() As Short
        Get
            codeSelected = mvarCodeSelected
        End Get
        Set(ByVal Value As Short)
            mvarCodeSelected = Value
        End Set
    End Property
    Public Property OpcsCode() As Short
        Get
            OpcsCode = mvarOpcsCode
        End Get
        Set(ByVal Value As Short)
            mvarOpcsCode = Value
        End Set
    End Property
    Public Property MaxLen() As Short
        Get
            MaxLen = mvarMaxLen
        End Get
        Set(ByVal Value As Short)
            mvarMaxLen = Value
        End Set
    End Property
    Public Property MsgBox_Answer() As Short
        Get
            MsgBox_Answer = mvarMsgBox_Answer
        End Get
        Set(ByVal Value As Short)
            mvarMsgBox_Answer = Value
        End Set
    End Property
    Public Property Layers() As Short
        Get
            Layers = mvarLayers
        End Get
        Set(ByVal Value As Short)
            mvarLayers = Value
        End Set
    End Property
    Public Property File_Number() As Short
        Get
            File_Number = mvarFile_Number
        End Get
        Set(ByVal Value As Short)
            mvarFile_Number = Value
        End Set
    End Property
    Public Property Errors() As Short
        Get
            Errors = mvarErrors
        End Get
        Set(ByVal Value As Short)
            mvarErrors = Value
        End Set
    End Property
    Public Property SubaLines() As Short
        Get
            SubaLines = MvarSubaLines
        End Get
        Set(ByVal Value As Short)
            MvarSubaLines = Value
        End Set
    End Property
    Public Property NumProj() As Short
        Get
            NumProj = MvarNumProj
        End Get
        Set(ByVal Value As Short)
            MvarNumProj = Value
        End Set
    End Property
    Public Property ngn() As Short
        Get
            ngn = mvarngn
        End Get
        Set(ByVal Value As Short)
            mvarngn = Value
        End Set
    End Property
    Public Property YearSim() As Short
        Get
            YearSim = mvarYearSim
        End Get
        Set(ByVal Value As Short)
            mvarYearSim = Value
        End Set
    End Property
    Public Property FEMRes() As Short
        Get
            FEMRes = MvarFEMRes
        End Get
        Set(ByVal Value As Short)
            MvarFEMRes = Value
        End Set
    End Property
    Public Property cancel1() As Short
        Get
            cancel1 = mvarcancel1
        End Get
        Set(ByVal Value As Short)
            mvarcancel1 = Value
        End Set
    End Property
    Public Property col1() As Short
        Get
            col1 = mvarcol1
        End Get
        Set(ByVal Value As Short)
            mvarcol1 = Value
        End Set
    End Property
    Public Property col2() As Short
        Get
            col2 = mvarcol2
        End Get
        Set(ByVal Value As Short)
            mvarcol2 = Value
        End Set
    End Property
    Public Property col3() As Short
        Get
            col3 = mvarcol3
        End Get
        Set(ByVal Value As Short)
            mvarcol3 = Value
        End Set
    End Property
    Public Property col4() As Short
        Get
            col4 = mvarcol4
        End Get
        Set(ByVal Value As Short)
            mvarcol4 = Value
        End Set
    End Property
    Public Property col5() As Short
        Get
            col5 = mvarcol5
        End Get
        Set(ByVal Value As Short)
            mvarcol5 = Value
        End Set
    End Property
    Public Property col6() As Short
        Get
            col6 = mvarcol6
        End Get
        Set(ByVal Value As Short)
            mvarcol6 = Value
        End Set
    End Property
    Public Property col7() As Short
        Get
            col7 = mvarcol7
        End Get
        Set(ByVal Value As Short)
            mvarcol7 = Value
        End Set
    End Property
    Public Property col8() As Short
        Get
            col8 = mvarcol8
        End Get
        Set(ByVal Value As Short)
            mvarcol8 = Value
        End Set
    End Property
    Public Property limit() As Short
        Get
            limit = mvarlimit
        End Get
        Set(ByVal Value As Short)
            mvarlimit = Value
        End Set
    End Property


    Public Property Bas() As String
        Get
            Bas = mvarBas
        End Get
        Set(ByVal Value As String)
            mvarBas = Value
        End Set
    End Property
    Public Property Sce() As String
        Get
            Sce = mvarSce
        End Get
        Set(ByVal Value As String)
            mvarSce = Value
        End Set
    End Property
    Public Property Dir_Bas() As String
        Get
            Dir_Bas = MvarDir_Bas
        End Get
        Set(ByVal Value As String)
            MvarDir_Bas = Value
        End Set
    End Property
    Public Property Dir_Sce() As String
        Get
            Dir_Sce = MvarDir_Sce
        End Get
        Set(ByVal Value As String)
            MvarDir_Sce = Value
        End Set
    End Property
    Public Property Dir_BasF() As String
        Get
            Dir_BasF = MvarDir_BasF
        End Get
        Set(ByVal Value As String)
            MvarDir_BasF = Value
        End Set
    End Property
    Public Property Dir_SceF() As String
        Get
            Dir_SceF = MvarDir_SceF
        End Get
        Set(ByVal Value As String)
            MvarDir_SceF = Value
        End Set
    End Property
    Public Property deleted() As String
        Get
            deleted = mvardeleted
        End Get
        Set(ByVal Value As String)
            mvardeleted = Value
        End Set
    End Property
    Public Property cont() As String
        Get
            cont = mvarcont
        End Get
        Set(ByVal Value As String)
            mvarcont = Value
        End Set
    End Property
    Public Property sublines() As String
        Get
            sublines = mvarsublines
        End Get
        Set(ByVal Value As String)
            mvarsublines = Value
        End Set
    End Property
    Public Property herd() As String
        Get
            herd = mvarHerd
        End Get
        Set(ByVal Value As String)
            mvarHerd = Value
        End Set
    End Property
    Public Property Fertilizer() As String
        Get
            Fertilizer = mvarfertilizer
        End Get
        Set(ByVal Value As String)
            mvarfertilizer = Value
        End Set
    End Property
    Public Property parm() As String
        Get
            parm = mvarparm
        End Get
        Set(ByVal Value As String)
            mvarparm = Value
        End Set
    End Property
    Public Property Till() As String
        Get
            Till = mvartill
        End Get
        Set(ByVal Value As String)
            mvartill = Value
        End Set
    End Property
    Public Property wpm1() As String
        Get
            wpm1 = mvarwpm1
        End Get
        Set(ByVal Value As String)
            mvarwpm1 = Value
        End Set
    End Property
    Public Property Opcs() As String
        Get
            Opcs = mvaropcs
        End Get
        Set(ByVal Value As String)
            mvaropcs = Value
        End Set
    End Property
    Public Property Site() As String
        Get
            Site = mvarsite
        End Get
        Set(ByVal Value As String)
            mvarsite = Value
        End Set
    End Property
    Public Property Soil() As String
        Get
            Soil = mvarsoil
        End Get
        Set(ByVal Value As String)
            mvarsoil = Value
        End Set
    End Property
    Public Property suba() As String
        Get
            suba = mvarsuba
        End Get
        Set(ByVal Value As String)
            mvarsuba = Value
        End Set
    End Property
    Public Property RunFile() As String
        Get
            RunFile = mvarRunFile
        End Get
        Set(ByVal Value As String)
            mvarRunFile = Value
        End Set
    End Property
    Public Property cntrl() As String
        Get
            cntrl = mvarcntrl
        End Get
        Set(ByVal Value As String)
            mvarcntrl = Value
        End Set
    End Property
    Public Property cntrl1() As String
        Get
            cntrl1 = mvarcntrl1
        End Get
        Set(ByVal Value As String)
            mvarcntrl1 = Value
        End Set
    End Property
    Public Property cntrl2() As String
        Get
            cntrl2 = mvarcntrl2
        End Get
        Set(ByVal Value As String)
            mvarcntrl2 = Value
        End Set
    End Property
    Public Property cntrl3() As String
        Get
            cntrl3 = mvarcntrl3
        End Get
        Set(ByVal Value As String)
            mvarcntrl3 = Value
        End Set
    End Property
    Public Property cntrl4() As String
        Get
            cntrl4 = mvarcntrl4
        End Get
        Set(ByVal Value As String)
            mvarcntrl4 = Value
        End Set
    End Property
    Public Property Project() As String
        Get
            Project = mvarProject
        End Get
        Set(ByVal Value As String)
            mvarProject = Value
        End Set
    End Property
    Public Property Scenario() As String
        Get
            Scenario = mvarScenario
        End Get
        Set(ByVal Value As String)
            mvarScenario = Value
        End Set
    End Property
    Public Property Espace() As String
        Get
            Espace = mvarEspace
        End Get
        Set(ByVal Value As String)
            mvarEspace = Value
        End Set
    End Property
    Public Property Espace1() As String
        Get
            Espace1 = mvarEspace1
        End Get
        Set(ByVal Value As String)
            mvarEspace1 = Value
        End Set
    End Property
    Public Property Version2() As String
        Get
            Version2 = mvarVersion2
        End Get
        Set(ByVal Value As String)
            mvarVersion2 = Value
        End Set
    End Property
    Public Property Version() As String
        Get
            Version = mvarVersion
        End Get
        Set(ByVal Value As String)
            mvarVersion = Value
        End Set
    End Property
    Public Property Version1() As String
        Get
            Version1 = mvarVersion1
        End Get
        Set(ByVal Value As String)
            mvarVersion1 = Value
        End Set
    End Property
    Public Property Dir1() As String
        Get
            Dir1 = mvarDir1
        End Get
        Set(ByVal Value As String)
            mvarDir1 = Value
        End Set
    End Property
    Public Property Dir2() As String
        Get
            Dir2 = mvardir2
        End Get
        Set(ByVal Value As String)
            mvardir2 = Value
        End Set
    End Property
    Public Property OrgDir() As String
        Get
            OrgDir = mvarOrgDir
        End Get
        Set(ByVal Value As String)
            mvarOrgDir = Value
        End Set
    End Property
    Public Property Input_Files() As String
        Get
            Input_Files = mvarInput_Files
        End Get
        Set(ByVal Value As String)
            mvarInput_Files = Value
        End Set
    End Property
    Public Property Swat_Output() As String
        Get
            Swat_Output = mvarSWAT_Output
        End Get
        Set(ByVal Value As String)
            mvarSWAT_Output = Value
        End Set
    End Property
    Public Property New_Swat() As String
        Get
            New_Swat = mvarNew_Swat
        End Get
        Set(ByVal Value As String)
            mvarNew_Swat = Value
        End Set
    End Property
    Public Property Output_files() As String
        Get
            Output_files = mvarOutput_files
        End Get
        Set(ByVal Value As String)
            mvarOutput_files = Value
        End Set
    End Property
    Public Property FEM() As String
        Get
            FEM = mvarFEM
        End Get
        Set(ByVal Value As String)
            mvarFEM = Value
        End Set
    End Property
    Public Property soilName() As String
        Get
            soilName = mvarSoilName
        End Get
        Set(ByVal Value As String)
            mvarSoilName = Value
        End Set
    End Property

    Public Property landUse() As String
        Get
            landUse = mvarLandUse
        End Get
        Set(ByVal Value As String)
            mvarLandUse = Value
        End Set
    End Property

    Public Property Year_Col As UShort
        Get
            Year_Col = mvaryear_col
        End Get
        Set(ByVal Value As UShort)
            mvaryear_col = Value
        End Set
    End Property

    Public Property Scenario_type As String
        Get
            Scenario_type = mvarScenario_type
        End Get
        Set(ByVal Value As String)
            mvarScenario_type = Value
        End Set
    End Property

    Public Property bmp As String
        Get
            bmp = mvarBMP
        End Get
        Set(ByVal Value As String)
            mvarBMP = Value
        End Set
    End Property
End Module