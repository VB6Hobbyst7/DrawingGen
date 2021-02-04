'Requires AutoCAD 2015 64-bit installed
'Requires Microsoft Office 2010 64-bit installed
'Requires Microsoft Visual Basic 2017 installed
'Requires Cad2Win License install ... 1/25/2021
'Requires References
'   COM  Microsoft ActiveX Data Objects 2.8 Library
'   File C:\Program Files\Common Files\Autodesk Shared\acax18enu.tlb
'   File C:\Program Files\Common Files\Autodesk Shared\axdb18enu.tlb
'
'
'   Added Custom Connection Code 6/24/2015 Ed Jackson

' Adding SharePoint and Config Versioning.

Option Explicit On

'Imports Autodesk.AutoCAD.Interop
Imports AutoCAD
'Imports Autodesk.AutoCAD.Interop.Common
Imports System.IO
Imports System.Data
Imports System.Data.OleDb
Imports System.Threading
Imports ADODB
'test comment



Module DrawGen

    Dim sDataPath As String                 'Arg Path to Product Data folder
    Dim sInFile As String                   'Arg Path and file name of input config file
    Dim bDebug As Boolean                   'Arg Turn debug on
    Dim sOutPath As String                  'Path to write drawing files - same folder as sInFile
    Dim sPgmsPath As String                 'Path to ancillary programs

    Dim sDrawingFileName As String          'List of drawings created written to sOutPath
    Dim sDebugFileName As String            'Debug trace file written to sOutPath
    Dim sErrorFileName As String            'Error messages written to sOutPath
    Dim ErrorCount As Integer               'Number of errors found
    Dim WarningCount As Integer             'Number of warnings found
    Dim sTrace As String                    'Current execution location
    Dim sAssembly As String                 'Major.Minor.Build.Revision
    Public Const LOG_WARNING As String = "Warning"          'Log error as warning and continue processing
    Public Const LOG_ERROR As String = "Error"              'Log error and continue processing
    Public Const LOG_FATAL_ERROR As String = "Fatal Error"  'Log error and exit program

    Dim DatabaseDB As String                'ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=X:\Path\Access.mdb; Persist Security Info=False"
    Dim DB As ADODB.Connection              'Model Database object definition
    Dim DBN As ADODB.Connection             'Standard Notes Database object definition
    Dim Acces As ADODB.Recordset            'Recordset from the DB definition for the Accessories
    Dim InConns As ADODB.Recordset          'Recordset from the DB definition for the Connections
    Dim OutConns As ADODB.Recordset         'Recordset from the DB definition for the Connections
    Dim Drive As ADODB.Recordset            'Recordset from the DB definition for the Drives
    Dim Data As ADODB.Recordset             'Recordset from the DB definition for the unit for a specific model
    Dim Anchor As ADODB.Recordset           'Recordset from the DB definition for the Anchorage
    Dim FlowDB As ADODB.Recordset           'Recordset from the DB definition for the Flow table
    Dim NotesRecSet As ADODB.Recordset      'Recordset from the DB definition for the Notes table
    Dim tblVersion As ADODB.Recordset       'Recordset from the DB definition for the Version table
    Dim TransTable As ADODB.Recordset       'Recordset from the DB definition for the Translation table

    '**Added 6/2/2015
    Dim Custom As ADODB.Recordset           'Recordset from the DB definition for the Custom Items
    '**End Added

    Dim ModelNum As String                  'string variable of the model number extracted from the configuration file
    Dim AryConfigFile(0) As String          'String Array of the entire config file, each count equals a line from the config file
    Dim intLineCount As Integer             'Integer of the count of the line in the config file
    Dim TranslatedConfig(0) As String       'Array of the translated config file
    Dim TranslateCount As Integer           'Line count for the translated file
    Dim Accessories(0, 0) As String         'String array of all the accessories that have been chosen, there are 31 variables for each acces
    Dim Accesscnt As Integer                'total count of the accessories

    '**Added 6/2/2015
    Dim CustomItems(0, 0) As String         'String array of all the Custom items.  May not be needed
    Dim Customcnt As Integer                'total count of the accessories    May not be needed
    '**End Added

    Dim DwgData(0, 0) As String             'contains the file, layers, notes and dims for each drawing
    Dim DwgCount As Integer                 'total number of drawings to be generated
    Dim AryGrp() As String                  'Array of the group names in each drawing    '3.0.2.0
    Dim lngWieght(12) As Long               'long array of the all the weight data including distribution points
    Dim Flow As Long                        'string containing the flow
    Dim InletSize As String                 'Size override for the inlet
    Dim OutletSize As String                'size override for the outlet
    Dim EqSize As String                    'size override for the equalizer
    Dim BySize As String                    'size override for the bypass
    Dim EmailAddress As String              'Contains the email address for the submittal
    Dim strCustData As String               'Suplementary Customomer data gathered from the website
    Dim Hand As String                      'The value for the hand of the unit
    Dim strFileSuffix As String             'Suffix that is attatched to the end of file name
    Dim AccGroup(5) As String               'Accessory groups current job falls into
    Dim FileType As String                  'Filetype requested from peopleSoft
    Dim ModelCallout As String              'Variable to accomodate the extended model number
    Dim OEMBlock As String                  'Path to title block drawing
    Dim OEMCopy As String                   'Path to copyright drawing
    Dim ConfigVersion As String             'Configuration run iterarion
    Dim SharePointSite As String            'Site name for the SharePOint intgration
    


    Sub Main()

        sAssembly = My.Application.Info.AssemblyName & " " & My.Application.Info.Version.ToString
        Console.WriteLine(sAssembly)
        sTrace = Now() & "  Main: " & sAssembly & Environment.NewLine

        Call Initialize()

        Call LoadJobFile()

        If AryConfigFile(0) <> "AUTO-SUB Config File" Then
            'file is generated from Peoplesoft
            Call FileTranslate()
        End If

        Call SubProcess()

        Call EndProgram()

    End Sub 'Main



    Sub Initialize()

        'Args(0) Full path and file name of this executable program
        'Args(1) Full path to drawing data folder
        'Args(2) Full path and file name of input configuration file

        Dim asArgs() As String = System.Environment.GetCommandLineArgs()
        Dim iCount As Integer
        Dim iPos As Integer
        Dim j As Integer
        Dim iTrace As Integer

        On Error GoTo Err_Initialize

        sTrace = sTrace & Now() & "  Initialize: Enter" & Environment.NewLine

        iTrace = 201
        sTrace = sTrace & Now() & "  Initialize: Get Args() Count" & Environment.NewLine
        iCount = System.Environment.GetCommandLineArgs.Count
        If iCount < 3 Then
            GoTo Err_Initialize
        End If
        sTrace = sTrace & Now() & "  Initialize: Args() Count = " & iCount & Environment.NewLine

        iTrace = 202
        sTrace = sTrace & Now() & "  Initialize: Get Args(0) sPgmsPath" & Environment.NewLine
        sPgmsPath = asArgs(0)
        iPos = InStrRev(sPgmsPath, "\")
        sPgmsPath = Left(sPgmsPath, iPos)
        If Right(sPgmsPath, 1) = "\" Then sPgmsPath = Left(sPgmsPath, Len(sPgmsPath) - 1)
        sTrace = sTrace & Now() & "  Initialize: sPgmsPath = " & sPgmsPath & Environment.NewLine

        iTrace = 203
        sTrace = sTrace & Now() & "  Initialize: Get Args(1) sDataPath" & Environment.NewLine
        sDataPath = asArgs(1)
        If Right(sDataPath, 1) = "\" Then sDataPath = Left(sDataPath, Len(sDataPath) - 1)
        If Dir(sDataPath, vbDirectory) = "" Then GoTo Err_Initialize
        sTrace = sTrace & Now() & "  Initialize: sDataPath = " & sDataPath & Environment.NewLine

        iTrace = 204
        sTrace = sTrace & Now() & "  Initialize: Get Args(2) sInFile" & Environment.NewLine
        sInFile = asArgs(2)
        If Dir(sInFile) = "" Then GoTo Err_Initialize
        sTrace = sTrace & Now() & "  Initialize: sInFile = " & sInFile & Environment.NewLine

        iTrace = 205
        sTrace = sTrace & Now() & "  Initialize: Get sOutPath" & Environment.NewLine
        iPos = InStrRev(sInFile, "\")
        sOutPath = Left(sInFile, iPos)
        sTrace = sTrace & Now() & "  Initialize: sOutPath = " & sOutPath & Environment.NewLine

        iTrace = 206
        sTrace = sTrace & Now() & "  Initialize: Get bDebug Flag" & Environment.NewLine
        If Dir(sPgmsPath & "\DrawGen.debug") = "" Then
            bDebug = False
        Else
            bDebug = True
        End If
        sTrace = sTrace & Now() & "  Initialize: bDebug = " & bDebug & Environment.NewLine

        iTrace = 207
        sTrace = sTrace & Now() & "  Initialize: Set log file names" & Environment.NewLine
        sDrawingFileName = "DrawGen.log"
        sDebugFileName = "Debug.log"
        sErrorFileName = "Error.log"
        sTrace = sTrace & Now() & "  Initialize: sDrawingFileName = " & sDrawingFileName & Environment.NewLine
        sTrace = sTrace & Now() & "  Initialize: sDebugFileName = " & sDebugFileName & Environment.NewLine
        sTrace = sTrace & Now() & "  Initialize: sErrorFileName = " & sErrorFileName & Environment.NewLine

        If bDebug Then
            File.AppendAllText(sOutPath & sDebugFileName, sTrace)
        End If

        sTrace = "Initialize: Create database objects" : LogDebug(sTrace)
        DatabaseDB = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & sDataPath & "\Sub Program - DBs.mdb;Persist Security Info=False"
        DB = New ADODB.Connection                       'Model Databases database object
        DBN = New ADODB.Connection                      'Standard Notes database object

        'sets the record sets to the tables used in each model
        sTrace = "Initialize: Set table objects" : LogDebug(sTrace)
        Acces = New ADODB.Recordset
        InConns = New ADODB.Recordset
        OutConns = New ADODB.Recordset
        Drive = New ADODB.Recordset
        Data = New ADODB.Recordset
        Anchor = New ADODB.Recordset
        FlowDB = New ADODB.Recordset
        NotesRecSet = New ADODB.Recordset
        tblVersion = New ADODB.Recordset
        TransTable = New ADODB.Recordset

        '**Added 6/2/2015
        Custom = New ADODB.Recordset
        '**End Added

        sTrace = "Initialize: Init arrays" : LogDebug(sTrace)
        DwgCount = 2
        ReDim DwgData(5, DwgCount - 1)
        For j = 0 To 5
            DwgData(j, 0) = ""
            DwgData(j, 1) = ""
        Next j

        ReDim Accessories(33, 0)
        For j = 0 To 32
            Accessories(j, 0) = ""
        Next j

        '**Added 6/2/2015
        ReDim CustomItems(5, 0)
        '**End Added

        '***Added 2/4/2021
        SharePointSite = "https://bacglobal.sharepoint.com/sites/AutocadTEST/Autocad%20TEST/"

        '***End Added


        sTrace = "Initialize: Exit" : LogDebug(sTrace)

        Exit Sub

Err_Initialize:

        If sOutPath <> "" And sErrorFileName <> "" Then
            Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)
        Else
        End If

        Environment.Exit(CInt(iTrace))

    End Sub 'Initialize



    Sub LoadJobFile()

        Dim inFile As FileStream
        Dim iInFileNum As Integer
        Dim textline As String

        sTrace = "LoadJobFile: Enter" : LogDebug(sTrace)

        On Error GoTo Err_LoadJobFile

        sTrace = "LoadJobFile: Open InFile" : LogDebug(sTrace)
        intLineCount = 0
        inFile = New FileStream(sInFile, FileMode.Open, FileAccess.Read)
        'iterates through each line of the opened config file
        sTrace = "LoadJobFile: Read InFile" : LogDebug(sTrace)
        Dim fileReader As New StreamReader(inFile)
        While fileReader.Peek > -1
            'adds a additional iteration to the config array
            ReDim Preserve AryConfigFile(intLineCount)
            textline = fileReader.ReadLine
            'takes the current line and adds it to the the next array iteration
            AryConfigFile(intLineCount) = RTrim(textline)
            intLineCount = intLineCount + 1
        End While
        'closes the current config file
        sTrace = "LoadJobFile: Close InFile" : LogDebug(sTrace)
        fileReader.Close()
        inFile.Close()

        If intLineCount < 2 Then
            sTrace = "LoadJobFile: Found empty job file" : LogDebug(sTrace)
            GoTo Err_LoadJobFile
        End If

        sTrace = "LoadJobFile: Exit" : LogDebug(sTrace)

        Exit Sub

Err_LoadJobFile:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'LoadJobFile



    Sub FileTranslate()

        'translator for peoplesoft output
        Dim i As Integer
        Dim h As Integer
        Dim ModelDB As String
        Dim NotesDB As String

        Dim PSModelTypeIn As String
        Dim Lang As String
        Dim Units As String
        Dim ConfigSetLen As Integer
        Dim ConfigFileLen As Integer
        Dim LineLen As Integer
        Dim TempModel As String
        Dim Replace As String
        Dim ReplaceLen As Integer
        Dim Replaced As String

        sTrace = "FileTranslate: Enter" : LogDebug(sTrace)

        On Error GoTo Err_FileTranslate
        ConfigVersion = "Nothing"
        TranslateCount = 4
        ReDim TranslatedConfig(TranslateCount)

        'For i = 0 To intLineCount - 1
        '    LineLen = Len(AryConfigFile(i))
        '    For h = 0 To 59
        '       If Mid(AryConfigFile(i), LineLen - h, 1) <> " " Then Exit For
        '    Next h
        'AryConfigFile(i) = Left(AryConfigFile(i), LineLen - h)
        'Next i

        PSModelTypeIn = AryConfigFile(3)
        FileType = "dwg"
        FileType = Left(LCase(AryConfigFile(2)), 3)

        '3.0.1.0 - INC019506 - 06/16/2014 - Begin
        Lang = "ENGLISH"
        Units = "IMPERIAL-1ST"
        If Len(Trim(AryConfigFile(4))) = 0 Or Len(Trim(AryConfigFile(4))) > 20 Then
        Else
            Lang = Trim(AryConfigFile(4))
            Units = Trim(AryConfigFile(5))
        End If
        '3.0.1.0 - INC019506 - 06/16/2014 - End
        ModelDB = ""

        sTrace = "FileTranslate: Open Database database" : LogDebug(sTrace)
        DB.ConnectionString = DatabaseDB
        DB.Open()
        sTrace = "FileTranslate: Open Databases table" : LogDebug(sTrace)
        Dim ModelTable As ADODB.Recordset
        ModelTable = New ADODB.Recordset
        ModelTable.LockType = LockTypeEnum.adLockOptimistic
        ModelTable.Open("DataBases", DB, CursorTypeEnum.adOpenDynamic)
        ModelTable.MoveFirst()
        Do While ModelTable.EOF = False
            '3.0.1.0 - INC019506 - 06/16/2014 - Begin
            'If ModelTable.Fields.Item("PeopleRef").Value.ToString = PSModelTypeIn Then ModelDB = ModelTable.Fields.Item("DataBase").Value.ToString
            'ModelTable.MoveNext()
            If ModelTable.Fields.Item("PeopleRef").Value.ToString = PSModelTypeIn And Lang = ModelTable.Fields.Item("Lang").Value.ToString And ModelTable.Fields.Item("Units").Value.ToString = Units Then
                ModelDB = ModelTable.Fields.Item("DataBase").Value.ToString
                sTrace = "FileTranslate: ModelDB = " & ModelDB : LogDebug(sTrace)
                Exit Do
            End If
            ModelTable.MoveNext()
            '3.0.1.0 - INC019506 - 06/16/2014 - End
        Loop
        sTrace = "FileTranslate: Close Databases Table" : LogDebug(sTrace)
        ModelTable.Close()

        '3.0.1.0 - INC019506 - 06/16/2014 - Begin
        sTrace = "FileTranslate: Open NotesDB table" : LogDebug(sTrace)
        ModelTable.Open("NotesDB", DB, CursorTypeEnum.adOpenDynamic)
        ModelTable.MoveFirst()
        Do While ModelTable.EOF = False
            If Lang = ModelTable.Fields.Item("Lang").Value.ToString Then
                NotesDB = ModelTable.Fields.Item("DataBase").Value.ToString
                sTrace = "FileTranslate: NotesDB = " & NotesDB : LogDebug(sTrace)
                Exit Do
            End If
            ModelTable.MoveNext()
        Loop
        sTrace = "FileTranslate: Close NotesDB Table" : LogDebug(sTrace)
        ModelTable.Close()
        '3.0.1.0 - INC019506 - 06/16/2014 - End

        sTrace = "FileTranslate: Close Database database" : LogDebug(sTrace)
        DB.Close()

        If ModelDB = "" Then
            sTrace = "FileTranslate: ModelDB not found" : LogDebug(sTrace)
            GoTo Err_FileTranslate
        End If

        TranslatedConfig(1) = AryConfigFile(0)
        TranslatedConfig(2) = "Database=" & sDataPath & ModelDB
        TranslatedConfig(3) = "FileType=" & AryConfigFile(2)

        sTrace = "FileTranslate: Open Model database" : LogDebug(sTrace)
        DB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & sDataPath & ModelDB & ";Persist Security Info=False"
        DB.Open()

        sTrace = "FileTranslate: Open Trans table" : LogDebug(sTrace)
        TransTable.LockType = LockTypeEnum.adLockOptimistic
        TransTable.Open("Trans", DB, CursorTypeEnum.adOpenDynamic)  'opens the table for the translator

        '3.0.1.0 - INC019506 - 06/16/2014 - Begin
        If bDebug Then
            sTrace = "FileTranslate: Input AryConfigFile" : LogDebug(sTrace)
            For i = 0 To intLineCount - 1
                sTrace = "FileTranslate: AryConfigFile(" & i & ") = " & AryConfigFile(i) : LogDebug(sTrace)
            Next
        End If
        '3.0.1.0 - INC019506 - 06/16/2014 - End

        sTrace = "FileTranslate: Translate AryConfigFile" : LogDebug(sTrace)
        For i = 5 To intLineCount - 1   'itereate through the data base entries and config file to match up values and rewrite a new config file
            ConfigFileLen = Len(AryConfigFile(i))
            TransTable.MoveFirst()
            Do While TransTable.EOF = False
                ConfigSetLen = Len(TransTable.Fields("PeopleSoftCat").Value) + 1
                If Left(AryConfigFile(i), ConfigSetLen) = TransTable.Fields("PeopleSoftCat").Value.ToString & " " Then
                    If Mid(AryConfigFile(i), 31, 8) <> "010_NONE" Then
                        TranslateCount = TranslateCount + 1
                        ReDim Preserve TranslatedConfig(TranslateCount)
                        If TransTable.Fields("ConfigType").Value.ToString = "Inlet" Then
                            TranslatedConfig(TranslateCount) = "Inlet=" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Outlet" Then
                            TranslatedConfig(TranslateCount) = "Outlet=" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Drive" Then
                            TranslatedConfig(TranslateCount) = "Drive=" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Anchorage" Then
                            TranslatedConfig(TranslateCount) = "Anchorage=" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString.ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Accessory" Then
                            TranslatedConfig(TranslateCount) = "*A*-" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If

                        '**Added 6/2/2015
                        If TransTable.Fields("ConfigType").Value.ToString = "Custom" Then
                            TranslatedConfig(TranslateCount) = "*C*-" & TransTable.Fields("Prefix").Value.ToString & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        '**End Added

                        '**Added 11/5/2020cOM
                        If TransTable.Fields("ConfigType").Value.ToString = "ConfigRev" Then
                            TranslatedConfig(TranslateCount) = "ConfigRev=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)

                            ConfigVersion = Mid(TranslatedConfig(TranslateCount), Len(TranslatedConfig(TranslateCount)) - 10)
                        End If
                        '**End Added

                        If TransTable.Fields("ConfigType").Value.ToString = "InletSize" Then
                            TranslatedConfig(TranslateCount) = "InletSize=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 35)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "OutletSize" Then
                            TranslatedConfig(TranslateCount) = "OutletSize=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 35)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "EqualizerSize" Then
                            TranslatedConfig(TranslateCount) = "EqualizerSize=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 35)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "BypassSize" Then
                            TranslatedConfig(TranslateCount) = "BypassSize=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 35)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        '3.0.1.0 - INC019506 - 05/23/2014 - Begin
                        'If TransTable.Fields("ConfigType").Value.ToString = "Model" Then
                        '    TranslatedConfig(4) = "Model=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                        '    If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        '    If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        '    If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        'End If
                        If TransTable.Fields("ConfigType").Value = "Model" Then
                            TranslatedConfig(4) = "Model=" & Mid(AryConfigFile(i), 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value <> "none" Then Call ApplyReplace(TranslatedConfig(4), TranslatedConfig(4), TransTable)
                            If TransTable.Fields("Append").Value <> "none" Then Call ApplyAppend(TranslatedConfig(4), TranslatedConfig(4), TransTable)
                            If TransTable.Fields("CheckFor").Value <> "none" Then Call CheckForAppend(TranslatedConfig(4), TranslatedConfig(4), TransTable)
                        End If
                        '3.0.1.0 - INC019506 - 05/23/2014 - End
                        If TransTable.Fields("ConfigType").Value.ToString = "Model-Callout1" Then
                            TranslatedConfig(TranslateCount) = "ModelCallout1=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Model-Callout2" Then
                            TranslatedConfig(TranslateCount) = "ModelCallout2=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Model-Callout3" Then
                            TranslatedConfig(TranslateCount) = "ModelCallout3=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "Hand" Then
                            TranslatedConfig(TranslateCount) = "Hand=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                        If TransTable.Fields("ConfigType").Value.ToString = "OEM" Then
                            TranslatedConfig(TranslateCount) = "OEM=" & Mid(AryConfigFile(i).ToString, 31, ConfigFileLen - 30)
                            If TransTable.Fields("Replace").Value.ToString <> "none" Then Call ApplyReplace(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("Append").Value.ToString <> "none" Then Call ApplyAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                            If TransTable.Fields("CheckFor").Value.ToString <> "none" Then Call CheckForAppend(TranslatedConfig(TranslateCount), TranslatedConfig(TranslateCount), TransTable)
                        End If
                    End If
                End If
                TransTable.MoveNext()
            Loop
        Next i

        TranslateCount = TranslateCount + 2
        ReDim Preserve TranslatedConfig(TranslateCount)
        TranslatedConfig(TranslateCount - 1) = "Email=" & AryConfigFile(1)
        TranslatedConfig(TranslateCount) = "EOF"

        sTrace = "FileTranslate: Close Trans Table" : LogDebug(sTrace)
        TransTable.Close()
        sTrace = "FileTranslate: Close Model database" : LogDebug(sTrace)
        DB.Close()

        'removed for OEM support 8/12/09
        'For i = 1 To intLineCount - 1
        '    If Left(AryConfigFile(i), 5) = "MODEL" Then
        '        ConfigFileLen = Len(AryConfigFile(i))
        '        TempModel = Mid(AryConfigFile(i), 31, ConfigFileLen - 30)
        '    End If
        'Next i
        'TranslatedConfig(4) = "Model=" & TempModel

        intLineCount = TranslateCount
        ReDim AryConfigFile(TranslateCount)
        For i = 0 To TranslateCount
            AryConfigFile(i) = TranslatedConfig(i)
        Next i

        '3.0.1.0 - INC019506 - 06/16/2014 - Begin
        If bDebug Then
            sTrace = "FileTranslate: Translated AryConfigFile" : LogDebug(sTrace)
            For i = 0 To TranslateCount
                sTrace = "FileTranslate: AryConfigFile(" & i & ") = " & AryConfigFile(i) : LogDebug(sTrace)
            Next
        End If
        '3.0.1.0 - INC019506 - 06/16/2014 - End

        sTrace = "FileTranslate: Exit" : LogDebug(sTrace)

        Exit Sub

Err_FileTranslate:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'FileTranslate



    Sub ApplyReplace(Textin As String, ByRef TextOut As String, DBtable As ADODB.Recordset)

        Dim Replaced As String
        Dim ReplaceLen As Integer
        Dim Replace As String
        Dim j As Integer
        Dim bmark, emark As Integer

        sTrace = "ApplyReplace: Enter" : LogDebug(sTrace)

        TextOut = Textin

        sTrace = "ApplyReplace: Textin = " & Textin : LogDebug(sTrace)

        On Error GoTo Err_ApplyReplace

        If DBtable.Fields("Replace").Value.ToString <> "none" Then
            ReplaceLen = Len(TransTable.Fields("Replace").Value)
            For j = 1 To ReplaceLen
                If Mid(DBtable.Fields("Replace").Value.ToString, j, 1) = "[" Then bmark = j
                If Mid(DBtable.Fields("Replace").Value.ToString, j, 1) = "]" Then
                    emark = j
                    Replace = Mid(DBtable.Fields("Replace").Value.ToString, bmark + 1, emark - bmark - 1)
                    Replaced = Right(Textin, emark - bmark - 1)
                    If Replaced = Replace Then
                        Replaced = DBtable.Fields("ReplaceWith").Value.ToString
                        ReplaceLen = Len(Replace)
                        TextOut = Left(Textin, Len(Textin) - Len(Replace)) & Replaced
                        Exit For
                    End If
                End If
            Next j
        End If

        sTrace = "ApplyReplace: TextOut = " & TextOut : LogDebug(sTrace)
        sTrace = "ApplyReplace: Exit" : LogDebug(sTrace)

        Exit Sub

Err_ApplyReplace:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'ApplyReplace



    Sub ApplyAppend(Textin As String, ByRef TextOut As String, DBtable As ADODB.Recordset)

        'combines append fields to designated peoplesoft settings

        Dim h As Integer
        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim appendtext As String
        Dim Append As String
        Dim LineLen As Integer
        Dim Appendlen As Integer

        sTrace = "ApplyAppend: Enter" : LogDebug(sTrace)

        sTrace = "ApplyAppend: Textin = " & Textin : LogDebug(sTrace)

        On Error GoTo Err_ApplyAppend

        If DBtable.Fields("Append").Value.ToString <> "none" Then
            Append = DBtable.Fields("Append").Value.ToString
            Appendlen = Len(Append)
            For h = 1 To Appendlen
                If Mid(Append, h, 1) = "[" Then bmark = h
                If Mid(Append, h, 1) = "]" Then
                    emark = h
                    appendtext = Mid(Append, bmark + 1, emark - bmark - 1)
                    For i = 0 To intLineCount - 1
                        LineLen = Len(AryConfigFile(i))
                        If Left(AryConfigFile(i), 30) = appendtext & Space(30 - (emark - bmark - 1)) Then Textin = Textin & "_" & Right(AryConfigFile(i), LineLen - 30)
                    Next i
                End If
            Next h
        End If
        TextOut = Textin

        sTrace = "ApplyAppend: TextOut = " & TextOut : LogDebug(sTrace)
        sTrace = "ApplyAppend: Exit" : LogDebug(sTrace)

        Exit Sub

Err_ApplyAppend:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'ApplyAppend



    Sub CheckForAppend(Textin As String, ByRef TextOut As String, DBtable As ADODB.Recordset)

        Dim h As Integer
        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim Append As String
        Dim LineLen As Integer
        Dim Appendlen As Integer
        Dim Settingtext As String
        Dim CheckText As String
        Dim CheckLen As Integer

        sTrace = "CheckForAppend: Enter" : LogDebug(sTrace)

        sTrace = "CheckForAppend: Textin = " & Textin : LogDebug(sTrace)

        On Error GoTo Err_CheckForAppend

        If DBtable.Fields("Check").Value.ToString <> "none" Then
            Append = DBtable.Fields("CheckFor").Value.ToString
            Appendlen = Len(Append)
            CheckLen = Len(DBtable.Fields("Check").Value.ToString)
            For i = 0 To intLineCount - 1
                LineLen = Len(AryConfigFile(i).ToString)
                If Left(AryConfigFile(i).ToString, 30) = DBtable.Fields("Check").Value.ToString & Space(30 - (CheckLen)) Then
                    Settingtext = Right(AryConfigFile(i), LineLen - 30)
                    For h = 1 To Appendlen
                        If Mid(Append, h, 1) = "[" Then bmark = h
                        If Mid(Append, h, 1) = "]" Then
                            emark = h
                            CheckText = Mid(Append, bmark + 1, emark - bmark - 1)
                            If CheckText = Settingtext Then TextOut = Textin & DBtable.Fields("CheckAppend").Value.ToString
                        End If
                    Next h
                End If
            Next i
        End If

        On Error GoTo 0

        sTrace = "CheckForAppend: TextOut = " & TextOut : LogDebug(sTrace)
        sTrace = "CheckForAppend: Exit" : LogDebug(sTrace)

        Exit Sub

Err_CheckForAppend:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'CheckForAppend



    Sub SubProcess()

        Dim OEM As String
        Dim i As Integer
        Dim linecnt As Integer
        Dim ModelDB As String

        sTrace = "SubProcess: Enter" : LogDebug(sTrace)

        On Error GoTo Err_SubProcess

        OEM = "BAC"
        OEMBlock = ""
        OEMCopy = ""

        sTrace = "SubProcess: Check for OEM" : LogDebug(sTrace)
        For i = 0 To intLineCount - 1
            linecnt = Len(AryConfigFile(i))
            If Left(AryConfigFile(i), 4) = "OEM=" Then
                OEM = Right(AryConfigFile(i), (linecnt - 4))
                Exit For
            End If
        Next i

        sTrace = "SubProcess: Open Database database" : LogDebug(sTrace)
        DB.ConnectionString = DatabaseDB
        DB.Open()
        sTrace = "SubProcess: Open Title table" : LogDebug(sTrace)
        Dim TitleTable As ADODB.Recordset
        TitleTable = New ADODB.Recordset
        TitleTable.LockType = LockTypeEnum.adLockBatchOptimistic
        TitleTable.Open("TitleBlock", DB, CursorTypeEnum.adOpenDynamic)
        TitleTable.MoveFirst()
        Do While TitleTable.EOF = False
            If TitleTable.Fields.Item("OEM_Name").Value.ToString = OEM Then
                OEMBlock = TitleTable.Fields.Item("TitleBlockPath").Value.ToString
                OEMCopy = TitleTable.Fields.Item("CopyBlockPath").Value.ToString
            End If
            TitleTable.MoveNext()
        Loop
        sTrace = "SubProcess: Close Title table" : LogDebug(sTrace)
        TitleTable.Close()
        sTrace = "SubProcess: Close Database database" : LogDebug(sTrace)
        DB.Close()

        Dim ModelDBLen As Integer
        strFileSuffix = ""

        sTrace = "SubProcess: Open Model database" : LogDebug(sTrace)
        ModelDBLen = Len(AryConfigFile(2))
        ModelDB = Right(AryConfigFile(2), ModelDBLen - 9)   'defines the model database
        DB.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & ModelDB & ";Persist Security Info=False"
        DB.Open()

        'set the tables to the current data base. See the database guild for an explanation of the table
        'Acces.Open "Accessories", DB, adOpenDynamic
        sTrace = "SubProcess: Open InConnections table" : LogDebug(sTrace)
        InConns.Open("InConnections", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Open OutConnections table" : LogDebug(sTrace)
        OutConns.Open("OutConnections", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Open Drives table" : LogDebug(sTrace)
        Drive.Open("Drives", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Open Model_Data table" : LogDebug(sTrace)
        Data.Open("Model_Data", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Open Anchorage table" : LogDebug(sTrace)
        Anchor.Open("Anchorage", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Open Flows tables" : LogDebug(sTrace)
        FlowDB.Open("Flows", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "Open Version table" : LogDebug(sTrace)
        tblVersion.Open("Version", DB, CursorTypeEnum.adOpenDynamic)

        sTrace = "SubProcess: Call ModelSet" : LogDebug(sTrace)
        Call ModelSet()
        sTrace = "SubProcess: Call ConnGet" : LogDebug(sTrace)
        Call ConnGet()
        sTrace = "SubProcess: Call DriveSet" : LogDebug(sTrace)
        Call DriveSet()
        sTrace = "SubProcess: Call AccesGet" : LogDebug(sTrace)
        Call AccesGet()

        '**Added 6/2/2015
        sTrace = "SubProcess: Call CustomGet" : LogDebug(sTrace)
        Call CustomGet()
        '**end Added

        sTrace = "SubProcess: Call WieghtGet" : LogDebug(sTrace)
        Call WieghtGet()
        sTrace = "SubProcess: Call FlowSet" : LogDebug(sTrace)
        Call FlowSet()

        sTrace = "SubProcess: Call DataCompile" : LogDebug(sTrace)
        Call DataCompile()
        sTrace = "SubProcess: Call ConfigureDwg" : LogDebug(sTrace)
        Call ConfigureDwg()

        sTrace = "SubProcess: Close Model tables" : LogDebug(sTrace)
        'Acces.Close
        InConns.Close()
        OutConns.Close()
        Drive.Close()
        Data.Close()
        Anchor.Close()
        FlowDB.Close()
        tblVersion.Close()

        sTrace = "SubProcess: Close Model database" : LogDebug(sTrace)
        DB.Close()
        On Error GoTo 0

        strCustData = ""
        EmailAddress = ""

        sTrace = "SubProcess: Exit" : LogDebug(sTrace)

        Exit Sub

Err_SubProcess:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'SubProcess



    Sub ModelSet()

        'this sub iterates through config file to find the model number,
        'then iterates through the Model tables to set row per the model

        Dim i As Integer
        Dim linecnt As Integer
        Dim Mod1, Mod2, Mod3 As String

        sTrace = "ModelSet: Enter" : LogDebug(sTrace)

        InletSize = ""          'resets all the overide sizes
        OutletSize = ""
        BySize = ""
        EqSize = ""
        Mod1 = ""
        Mod2 = ""
        Mod3 = ""

        On Error GoTo Err_ModelSet
        sTrace = "ModelSet: Get Model" : LogDebug(sTrace)

        For i = 0 To intLineCount - 1
            linecnt = Len(AryConfigFile(i))
            If Left(AryConfigFile(i), 6) = "Model=" Then
                ModelNum = Right(AryConfigFile(i), (linecnt - 6))
            End If
            If Left(AryConfigFile(i), 10) = "InletSize=" And linecnt > 10 Then
                InletSize = Right(AryConfigFile(i), (linecnt - 10))
            End If
            If Left(AryConfigFile(i), 11) = "OutletSize=" And linecnt > 11 Then
                OutletSize = Right(AryConfigFile(i), (linecnt - 11))
            End If
            If Left(AryConfigFile(i), 14) = "EqualizerSize=" And linecnt > 14 Then
                EqSize = Right(AryConfigFile(i), (linecnt - 14))
            End If
            If Left(AryConfigFile(i), 11) = "BypassSize=" And linecnt > 11 Then
                BySize = Right(AryConfigFile(i), (linecnt - 11))
            End If
            If Left(AryConfigFile(i), 6) = "Email=" And linecnt > 6 Then
                EmailAddress = Right(AryConfigFile(i), (linecnt - 6))
            End If
            If Left(AryConfigFile(i), 9) = "CustData=" And linecnt > 9 Then
                strCustData = Right(AryConfigFile(i), (linecnt - 9))
            End If
            If Left(AryConfigFile(i), 14) = "ModelCallout1=" And linecnt > 14 Then
                Mod1 = Right(AryConfigFile(i), (linecnt - 14))
            End If
            If Left(AryConfigFile(i), 14) = "ModelCallout2=" And linecnt > 14 Then
                Mod2 = Right(AryConfigFile(i), (linecnt - 14))
            End If
            If Left(AryConfigFile(i), 14) = "ModelCallout3=" And linecnt > 14 Then
                Mod3 = Right(AryConfigFile(i), (linecnt - 14))
            End If
        Next i

        If strCustData = "" Then strCustData = " "

        If Mod1 <> "" Then
            ModelCallout = Mod1
        Else
            If Mod2 <> "" Then
                ModelCallout = Mod2
            Else
                ModelCallout = Mod3
            End If
        End If

        If AryConfigFile(0) = "AUTO-SUB Config File" Then ModelCallout = ModelNum

        If ModelNum = "" Then
            sTrace = "ModelSet: ModelNum not found in AryConfigFile" : LogDebug(sTrace)
            GoTo Err_ModelSet
        Else
            sTrace = "ModelSet: ModelNum " & ModelNum & " found in AryConfigFile" : LogDebug(sTrace)
        End If

        sTrace = "ModelSet: Search Model" : LogDebug(sTrace)

        Data.MoveFirst()
        Do While Data.EOF = False
            If Data.Fields.Item("Model Number").Value.ToString = ModelNum Then Exit Do
            Data.MoveNext()
        Loop

        If Data.EOF = True Then
            sTrace = "ModelSet: ModelNum " & ModelNum & " not found in database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9) : LogDebug(sTrace)
            GoTo Err_ModelSet
        Else
            sTrace = "ModelSet: ModelNum " & ModelNum & " found in database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9) : LogDebug(sTrace)
        End If

        AccGroup(0) = Data.Fields("Box_Group").Value.ToString
        AccGroup(1) = Data.Fields("Cell_Group").Value.ToString
        AccGroup(2) = Data.Fields("Group_1").Value.ToString
        AccGroup(3) = Data.Fields("Group_2").Value.ToString
        AccGroup(4) = Data.Fields("Group_3").Value.ToString
        AccGroup(5) = Data.Fields("Group_4").Value.ToString

        sTrace = "ModelSet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_ModelSet:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'ModelSet



    Sub ConnGet()

        'this sub iterates through config file to find the inlet and outlet,
        'then iterates through the conns tables to set row per the model, inlet and outlet

        Dim i As Integer
        Dim strConnI As String = ""
        Dim strConnO As String = ""
        Dim linecnt As Integer
        Dim ErrorText As String

        sTrace = "ConnGet: Enter" : LogDebug(sTrace)

        On Error GoTo Err_ConnGet
        ErrorText = ""

        sTrace = "ConnGet: Get Connection Config" : LogDebug(sTrace)
        For i = 0 To intLineCount - 1
            If Left(AryConfigFile(i), 6) = "Inlet=" Then
                linecnt = Len(AryConfigFile(i))
                strConnI = Right(AryConfigFile(i), (linecnt - 6))
            End If
            If Left(AryConfigFile(i), 7) = "Outlet=" Then
                linecnt = Len(AryConfigFile(i))
                strConnO = Right(AryConfigFile(i), (linecnt - 7))
            End If
            If Left(AryConfigFile(i), 5) = "Hand=" Then
                linecnt = Len(AryConfigFile(i))
                Hand = Right(AryConfigFile(i), (linecnt - 5))
            End If
        Next i

        sTrace = "ConnGet: Set & Check for Connection Config Error" : LogDebug(sTrace)
        If strConnI = "" Then
            ErrorText = "Locate inlet "
        End If
        If strConnO = "" Then
            ErrorText = ErrorText & "Locate outlet "
        End If
        If ErrorText <> "" Then GoTo Err_ConnGet

        sTrace = "ConnGet: Get In Connection" : LogDebug(sTrace)
        InConns.MoveFirst()
        Do While InConns.EOF = False
            Select Case InConns.Fields.Item("Model Number").Value.ToString
                Case ModelNum, AccGroup(0), AccGroup(1), AccGroup(2), AccGroup(3), AccGroup(4), AccGroup(5), "ALL", "all", "All"
                    If InConns.Fields.Item("Hand").Value.ToString = Hand Or InConns.Fields.Item("Hand").Value.ToString = "none" Then
                        If InConns.Fields.Item("Inlet").Value.ToString = strConnI Then Exit Do
                    End If
            End Select
            InConns.MoveNext()
        Loop

        sTrace = "ConnGet: Set In Connection Error" : LogDebug(sTrace)
        If InConns.EOF = True Then
            ErrorText = ErrorText & "Locate In Connection " & strConnI & " for model number " & ModelNum & " in database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9) & " "
        End If

        sTrace = "ConnGet: Get Out Connection" : LogDebug(sTrace)
        OutConns.MoveFirst()
        Do While OutConns.EOF = False
            Select Case OutConns.Fields.Item("Model Number").Value.ToString
                Case ModelNum, AccGroup(0), AccGroup(1), AccGroup(2), AccGroup(3), AccGroup(4), AccGroup(5), "ALL", "all", "All"
                    If OutConns.Fields.Item("Hand").Value.ToString = Hand Or OutConns.Fields.Item("Hand").Value.ToString = "none" Then
                        If OutConns.Fields.Item("Outlet").Value.ToString = strConnO Then Exit Do
                    End If
            End Select
            OutConns.MoveNext()
        Loop

        sTrace = "ConnGet: Set Out Connection Error" : LogDebug(sTrace)
        If OutConns.EOF = True Then
            ErrorText = ErrorText & "Locate Out Connection " & strConnO & " for model number " & ModelNum & " in database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9) & " "
        End If

        sTrace = "ConnGet: Check for Connection Errors" : LogDebug(sTrace)
        If ErrorText <> "" Then GoTo Err_ConnGet

        sTrace = "ConnGet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_ConnGet:

        Call LogUnstructuredError("Error at " & sTrace & " " & ErrorText, LOG_FATAL_ERROR)

    End Sub 'ConnGet



    Sub DriveSet()

        'this sub iterates through config file to find the drive,
        'then iterates through the drive tables to set row per the model and drive type

        Dim i As Integer
        Dim strDrive As String = ""
        Dim linecnt As Integer
        Dim ErrorText As String

        sTrace = "DriveSet: Enter" : LogDebug(sTrace)

        On Error GoTo Err_DriveSet
        ErrorText = ""
        sTrace = "DriveSet: Get Drive Config" : LogDebug(sTrace)

        For i = 0 To intLineCount - 1
            If Left(AryConfigFile(i), 5) = "Drive" Then
                linecnt = Len(AryConfigFile(i))
                strDrive = Right(AryConfigFile(i), (linecnt - 6))
            End If
        Next i

        If strDrive = "" Then
            ErrorText = "Locate the drive in the config file"
            LogDebug(ErrorText)
            GoTo Err_DriveSet
        End If

        sTrace = "DriveSet: Get Drive" : LogDebug(sTrace)
        Drive.MoveFirst()
        Do While Drive.EOF = False
            Select Case Drive.Fields.Item("Model Number").Value.ToString
                Case ModelNum, AccGroup(0), AccGroup(1), AccGroup(2), AccGroup(3), AccGroup(4), AccGroup(5), "ALL", "all", "All"
                    If Drive.Fields.Item("Drive Type").Value.ToString = strDrive Then
                        If Drive.Fields.Item("Hand").Value.ToString = Hand Then Exit Do 'added 01/14/09
                        If LCase(Drive.Fields.Item("Hand").Value.ToString) = "none" Then Exit Do 'added 01/14/09
                    End If
            End Select
            Drive.MoveNext()
        Loop

        If Drive.EOF = True Then
            ErrorText = "Locate drive " & strDrive & "for model number " & ModelNum & " in database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9)
            LogDebug(ErrorText)
            GoTo Err_DriveSet
        End If

        sTrace = "DriveSet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_DriveSet:

        Call LogUnstructuredError("Error at " & sTrace & " " & ErrorText, LOG_FATAL_ERROR)

    End Sub 'DriveSet



    Sub AccesGet()

        'runs through the config file array to pull out the accessories and gather the accessory data
        'char length of each line for parsing the text

        Dim i As Integer
        Dim j As Integer
        Dim linecnt As Integer

        sTrace = "AccesGet: Enter" : LogDebug(sTrace)

        Accesscnt = 0

        On Error GoTo Err_AccesGet

        If AccGroup(0) = "none" Then AccGroup(0) = ModelNum
        If AccGroup(1) = "none" Then AccGroup(1) = ModelNum
        If AccGroup(2) = "none" Then AccGroup(2) = ModelNum
        If AccGroup(3) = "none" Then AccGroup(3) = ModelNum
        If AccGroup(4) = "none" Then AccGroup(4) = ModelNum
        If AccGroup(5) = "none" Then AccGroup(5) = ModelNum

        sTrace = "AccesGet: Select from acces table" : LogDebug(sTrace)

        Acces.Open("SELECT * FROM Accessories WHERE (Model_Number IN ('" & ModelNum & "', '" & AccGroup(0) & "','" & AccGroup(1) & "', '" & AccGroup(2) _
                                                     & "','" & AccGroup(3) & "','" & AccGroup(4) & "','" & AccGroup(5) & "','All','ALL','all'))", DB, CursorTypeEnum.adOpenDynamic)

        'iterate through each line in the config array
        For i = 0 To intLineCount - 1
            'checks for an accessory tag
            If Left(AryConfigFile(i), 3) = "*A*" Or Left(AryConfigFile(i), 3) = "*Z" Then
                linecnt = Len(AryConfigFile(i))
                sTrace = "AccesGet: Locate accessory " & i & " " & AryConfigFile(i) : LogDebug(sTrace)
                'redimension the accessory array to add additionnal accessories
                ReDim Preserve Accessories(33, Accesscnt)
                For j = 0 To 32
                    Accessories(j, Accesscnt) = ""
                Next j
                'moves to the front of the accessory table in access
                Acces.MoveFirst()
                Do While Acces.EOF = False
                    'check the current row for matching model number
                    Select Case Acces.Fields.Item("Model_Number").Value.ToString
                        Case ModelNum, AccGroup(0), AccGroup(1), AccGroup(2), AccGroup(3), AccGroup(4), AccGroup(5), "All", "ALL", "all"
                            If LCase(Acces.Fields.Item("Hand").Value.ToString) = LCase(Hand) _
                            Or LCase(Acces.Fields.Item("Hand").Value.ToString) = "none" Then
                                'check the current row matching accessory
                                If LCase(Acces.Fields.Item("Accessory").Value.ToString) = LCase(Right(AryConfigFile(i), linecnt - 4)) _
                                Or LCase(Acces.Fields.Item("WebPageDesc").Value.ToString) = LCase(Right(AryConfigFile(i), linecnt - 4)) Then
                                    'If Acces.BOF = False Or Acces.eof = False Then
                                    'assigns the fields from the row to the acces array
                                    Accessories(0, Accesscnt) = Acces.Fields.Item("Model_Number").Value.ToString
                                    Accessories(1, Accesscnt) = Acces.Fields.Item("Accessory").Value.ToString
                                    Accessories(3, Accesscnt) = Acces.Fields.Item("Accessory Key").Value.ToString
                                    Accessories(4, Accesscnt) = Acces.Fields.Item("Incompatable").Value.ToString
                                    Accessories(5, Accesscnt) = Acces.Fields.Item("File").Value.ToString
                                    Accessories(6, Accesscnt) = Acces.Fields.Item("Layer").Value.ToString
                                    Accessories(2, Accesscnt) = Acces.Fields.Item("LayerOff").Value.ToString
                                    Accessories(7, Accesscnt) = Acces.Fields.Item("Notes").Value.ToString
                                    Accessories(8, Accesscnt) = Acces.Fields.Item("SizedforFlow").Value.ToString
                                    Accessories(9, Accesscnt) = Acces.Fields.Item("Shipping Weight").Value.ToString
                                    Accessories(10, Accesscnt) = Acces.Fields.Item("Operating Weight").Value.ToString
                                    Accessories(33, Accesscnt) = Acces.Fields.Item("Heaviest Section Add").Value.ToString 'added 01/14/09
                                    Accessories(11, Accesscnt) = Acces.Fields.Item("Plan-A Point 1").Value.ToString
                                    Accessories(12, Accesscnt) = Acces.Fields.Item("Plan-A Point 2").Value.ToString
                                    Accessories(13, Accesscnt) = Acces.Fields.Item("Plan-A Point 3").Value.ToString
                                    Accessories(14, Accesscnt) = Acces.Fields.Item("Plan-A Point 4").Value.ToString
                                    Accessories(15, Accesscnt) = Acces.Fields.Item("Plan-A Point 5").Value.ToString
                                    Accessories(16, Accesscnt) = Acces.Fields.Item("Plan-A Point 6").Value.ToString
                                    Accessories(17, Accesscnt) = Acces.Fields.Item("Plan-A Point 7").Value.ToString
                                    Accessories(18, Accesscnt) = Acces.Fields.Item("Plan-A Point 8").Value.ToString
                                    Accessories(19, Accesscnt) = Acces.Fields.Item("Plan-A Point 9").Value.ToString
                                    Accessories(20, Accesscnt) = Acces.Fields.Item("Plan-A Point 10").Value.ToString
                                    Accessories(21, Accesscnt) = Acces.Fields.Item("Plan-B Point 1").Value.ToString
                                    Accessories(22, Accesscnt) = Acces.Fields.Item("Plan-B Point 2").Value.ToString
                                    Accessories(23, Accesscnt) = Acces.Fields.Item("Plan-B Point 3").Value.ToString
                                    Accessories(24, Accesscnt) = Acces.Fields.Item("Plan-B Point 4").Value.ToString
                                    Accessories(25, Accesscnt) = Acces.Fields.Item("Plan-B Point 5").Value.ToString
                                    Accessories(26, Accesscnt) = Acces.Fields.Item("Plan-B Point 6").Value.ToString
                                    Accessories(27, Accesscnt) = Acces.Fields.Item("Plan-B Point 7").Value.ToString
                                    Accessories(28, Accesscnt) = Acces.Fields.Item("Plan-B Point 8").Value.ToString
                                    Accessories(29, Accesscnt) = Acces.Fields.Item("Plan-B Point 9").Value.ToString
                                    Accessories(30, Accesscnt) = Acces.Fields.Item("Plan-B Point 10").Value.ToString
                                    Accessories(31, Accesscnt) = Acces.Fields.Item("Dims").Value.ToString
                                    Accessories(32, Accesscnt) = Acces.Fields.Item("Prefix").Value.ToString
                                    Exit Do
                                End If
                            End If
                    End Select
                    Acces.MoveNext()       'move to the next row in the table
                    If Acces.EOF = True Then
                        Accessories(0, Accesscnt) = "none"
                        Exit Do
                    End If
                Loop
                '        Acces.Close
                'adds another to the accessory count
                Accesscnt = Accesscnt + 1
            End If
        Next i
        sTrace = "AccesGet: Close acces table" : LogDebug(sTrace)
        Acces.Close()

        Dim x, y As Integer
        If bDebug Then
            For x = 0 To UBound(Accessories, 1)
                For y = 0 To UBound(Accessories, 2)
                    sTrace = "AccesGet: " & "Accessories(" & x & "," & y & ") = " & Accessories(x, y) : LogDebug(sTrace)
                Next
            Next
        End If

        sTrace = "AccesGet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_AccesGet:

        Call LogUnstructuredError("Error at " & sTrace & " in Database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9), LOG_ERROR)

    End Sub 'AccesGet


    Sub CustomGet() '**Added 6/2/2015
        Dim i As Integer
        Dim j As Integer
        Dim linecnt As Integer
        Dim CustomPre As String
        Dim StartMark As Integer
        Dim ValidLoc As Boolean
        Dim DimstringA As String
        Dim DimstringB As String

        sTrace = "CustomGet: Enter" : LogDebug(sTrace)

        Customcnt = 0

        On Error GoTo Err_CustomGet

        sTrace = "CustomGet: Select from Custom table" : LogDebug(sTrace)

        Custom.Open("SELECT * FROM Custom WHERE (Model_Number IN ('" & ModelNum & "', '" & AccGroup(0) & "','" & AccGroup(1) & "', '" & AccGroup(2) _
                                                 & "','" & AccGroup(3) & "','" & AccGroup(4) & "','" & AccGroup(5) & "','All','ALL','all'))", DB, CursorTypeEnum.adOpenDynamic)

        'iterate through each line in the config array
        For i = 0 To intLineCount - 1
            'checks for a custom tag
            If Left(AryConfigFile(i), 3) = "*C*" Then
                linecnt = Len(AryConfigFile(i))
                sTrace = "CustomGet: Locate accessory " & i & " " & AryConfigFile(i) : LogDebug(sTrace)
                'redimension the accessory array to add additionnal accessories

                ReDim Preserve Accessories(33, Accesscnt)
                For j = 0 To 32
                    Accessories(j, Accesscnt) = ""
                Next j

                'moves to the front of the Custom table in access
                Custom.MoveFirst()
                Do While Custom.EOF = False
                    'check the current row for matching model number
                    Select Case Custom.Fields.Item("Model_Number").Value.ToString
                        Case ModelNum, AccGroup(0), AccGroup(1), AccGroup(2), AccGroup(3), AccGroup(4), AccGroup(5), "All", "ALL", "all"
                            If LCase(Custom.Fields.Item("Hand").Value.ToString) = LCase(Hand) _
                            Or LCase(Custom.Fields.Item("Hand").Value.ToString) = "none" Then
                                'check the current row matching custom
                                CustomPre = LCase(Mid(AryConfigFile(i), 5, Len(Custom.Fields.Item("Custom_Prefix").Value.ToString)))
                                If LCase(Custom.Fields.Item("Custom_Prefix").Value.ToString) = CustomPre Then
                                    DimstringA = "error"
                                    DimstringB = "none"
                                    If Custom.Fields.Item("GreaterLessB").Value.ToString = "none" Then
                                        DimstringA = Right(AryConfigFile(i), Len(AryConfigFile(i)) - Len(Custom.Fields.Item("Custom_Prefix").Value.ToString) - 5)
                                        DimstringB = "none"
                                    Else
                                        For j = Len(CustomPre) + 6 To Len(AryConfigFile(i))
                                            If Mid(AryConfigFile(i), j, 1) = "_" Then
                                                DimstringA = Mid(AryConfigFile(i), Len(CustomPre) + 6, j - Len(CustomPre) - 6)
                                                DimstringB = Right(AryConfigFile(i), Len(AryConfigFile(i)) - j)
                                            End If
                                        Next
                                    End If

                                    If Custom.Fields.Item("Dim_Type").Value.ToString = "arch" Then
                                        ValidLoc = False
                                        If Custom.Fields.Item("GreaterLessB").Value.ToString = "<" Then
                                            If Custom.Fields.Item("GreaterLess_DimB").Value.ToString > DimstringB Then
                                                If Custom.Fields.Item("GreaterLess").Value.ToString = ">" Then
                                                    If DimstringA > Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "<" Then
                                                    If DimstringA < Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "=" Then
                                                    If DimstringA = Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                Else
                                                    ValidLoc = False
                                                End If
                                            End If
                                        ElseIf Custom.Fields.Item("GreaterLessB").Value.ToString = ">" Then
                                            If Custom.Fields.Item("GreaterLess_DimB").Value.ToString < DimstringB Then
                                                If Custom.Fields.Item("GreaterLess").Value.ToString = ">" Then
                                                    If DimstringA > Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "<" Then
                                                    If DimstringA < Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "=" Then
                                                    If DimstringA = Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                Else
                                                    ValidLoc = False
                                                End If
                                            End If
                                        ElseIf Custom.Fields.Item("GreaterLessB").Value.ToString = "=" Then
                                            If Custom.Fields.Item("GreaterLess_DimB").Value.ToString = DimstringB Then
                                                If Custom.Fields.Item("GreaterLess").Value.ToString = ">" Then
                                                    If DimstringA > Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "<" Then
                                                    If DimstringA < Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "=" Then
                                                    If DimstringA = Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                                Else
                                                    ValidLoc = False
                                                End If
                                            End If
                                        ElseIf LCase(Custom.Fields.Item("GreaterLessB").Value.ToString) = "none" Then
                                            If Custom.Fields.Item("GreaterLess").Value.ToString = ">" Then
                                                If DimstringA > Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                            ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "<" Then
                                                If DimstringA < Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                            ElseIf Custom.Fields.Item("GreaterLess").Value.ToString = "=" Then
                                                If DimstringA = Custom.Fields.Item("GreaterLess_Dim").Value.ToString Then ValidLoc = True
                                            ElseIf LCase(Custom.Fields.Item("GreaterLess").Value.ToString) = "none" Then
                                                ValidLoc = True
                                            Else
                                                ValidLoc = False
                                            End If

                                        End If
                                    Else
                                        ValidLoc = True
                                    End If


                                    If ValidLoc = True Then
                                        If LCase(Custom.Fields.Item("Dims").Value.ToString) <> "none" Then
                                            For j = 1 To Len(Custom.Fields.Item("Dims").Value.ToString)
                                                If Mid(Custom.Fields.Item("Dims").Value.ToString, j, 1) = "[" Then StartMark = j
                                                If Mid(Custom.Fields.Item("Dims").Value.ToString, j, 1) = "]" Then
                                                    Accessories(31, Accesscnt) = Accessories(31, Accesscnt) & Mid(Custom.Fields.Item("Dims").Value.ToString, _
                                                                                    StartMark, j - StartMark) & "=" & FormatDimText(DimstringA, _
                                                                                    Custom.Fields.Item("Seperator").Value.ToString, Custom.Fields.Item("Dim_Type").Value.ToString, _
                                                                                    Custom.Fields.Item("Custom_Prefix").Value.ToString) & "]"
                                                End If
                                            Next
                                        Else
                                            Accessories(31, Accesscnt) = "none"
                                        End If
                                        Accessories(0, Accesscnt) = Custom.Fields.Item("Model_Number").Value.ToString
                                        Accessories(1, Accesscnt) = Custom.Fields.Item("Custom_Prefix").Value.ToString
                                        Accessories(3, Accesscnt) = 0
                                        Accessories(4, Accesscnt) = "none"
                                        Accessories(5, Accesscnt) = Custom.Fields.Item("File").Value.ToString
                                        Accessories(6, Accesscnt) = Custom.Fields.Item("Layer Name").Value.ToString
                                        Accessories(2, Accesscnt) = "none"
                                        Accessories(7, Accesscnt) = Custom.Fields.Item("Notes").Value.ToString
                                        Accessories(8, Accesscnt) = False
                                        Accessories(9, Accesscnt) = 0
                                        Accessories(10, Accesscnt) = 0
                                        Accessories(33, Accesscnt) = 0
                                        Accessories(11, Accesscnt) = 0
                                        Accessories(12, Accesscnt) = 0
                                        Accessories(13, Accesscnt) = 0
                                        Accessories(14, Accesscnt) = 0
                                        Accessories(15, Accesscnt) = 0
                                        Accessories(16, Accesscnt) = 0
                                        Accessories(17, Accesscnt) = 0
                                        Accessories(18, Accesscnt) = 0
                                        Accessories(19, Accesscnt) = 0
                                        Accessories(20, Accesscnt) = 0
                                        Accessories(21, Accesscnt) = 0
                                        Accessories(22, Accesscnt) = 0
                                        Accessories(23, Accesscnt) = 0
                                        Accessories(24, Accesscnt) = 0
                                        Accessories(25, Accesscnt) = 0
                                        Accessories(26, Accesscnt) = 0
                                        Accessories(27, Accesscnt) = 0
                                        Accessories(28, Accesscnt) = 0
                                        Accessories(29, Accesscnt) = 0
                                        Accessories(30, Accesscnt) = 0
                                        Accessories(32, Accesscnt) = Custom.Fields.Item("Prefix").Value.ToString

                                        Exit Do
                                    End If
                                End If
                            End If
                    End Select
                    Custom.MoveNext()       'move to the next row in the table
                    If Custom.EOF = True Then
                        CustomItems(0, Customcnt) = "none"
                        Exit Do
                    End If
                Loop
                '        Acces.Close
                'adds another to the accessory count
                Accesscnt = Accesscnt + 1
            End If
        Next i
        sTrace = "CustomGet: Close acces table" : LogDebug(sTrace)
        Custom.Close()

        sTrace = "CustomGet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_CustomGet:

        Call LogUnstructuredError("Error at " & sTrace & " in Database " & Right(AryConfigFile(2), Len(AryConfigFile(2)) - 9), LOG_ERROR)

    End Sub 'CustomGet

    Function FindChar(Instr As String, ChartoFind As String, CharStart As Integer) As Integer '**Added 6/2/2015
        FindChar = 0
        Dim i As Integer
        For i = CharStart To Len(Instr)
            If Mid(Instr, i, 1) = ChartoFind Then
                FindChar = i
                Exit Function
            End If
        Next
    End Function

    Function FormatDimText(TextIn As String, Seperator As String, StringType As String, Prefix As String) As String '**Added 6/2/2015
        Dim i As Integer
        Dim mark As Integer
        Dim TempText As String
        Dim Feet As String
        Dim Inches As String
        Dim Numerator As Long
        Dim Denominator As Long
        Dim Inchdecimal As Double

        FormatDimText = ""
        On Error GoTo FormatDimTextError
        mark = 0
        If LCase(StringType) = "text" Then
            For i = 1 To Len(TextIn)
                If Mid(TextIn, i, 1) = "_" Then
                    FormatDimText = FormatDimText & Mid(TextIn, mark + 1, i - mark - 1) & Seperator
                    mark = i
                End If
            Next
            FormatDimText = FormatDimText & Mid(TextIn, mark + 1, Len(TextIn) - mark)
        End If

        mark = 0
        If LCase(StringType) = "arch" Then
            TempText = TextIn
            If Left(TempText, 1) = "-" Then TempText = Mid(TempText, 2, Len(TempText) - 1)
            'If TempText Like "#.#" = False Then GoTo FormatDimText

            For i = 1 To Len(TempText)
                If Mid(TempText, i, 1) = "." Then mark = i
            Next

            Feet = (TempText - (TempText Mod 12)) / 12
            Inches = (TempText - (TempText Mod 1)) - Feet * 12

            If mark <> 0 Then
                Inchdecimal = TempText - Feet * 12 - Inches
                Numerator = Math.Round((Inchdecimal * 1000) / 125)
                Select Case Numerator
                    Case 2
                        Numerator = "1"
                        Denominator = "4"
                        Exit Select
                    Case 6
                        Numerator = "3"
                        Denominator = "4"
                        Exit Select
                    Case 4
                        Numerator = "1"
                        Denominator = "2"
                        Exit Select
                    Case Else
                        Denominator = "8"
                End Select
            End If

            If Feet <> "0" Then
                FormatDimText = Feet & "'"
                If Inches <> "0" Or Inchdecimal <> 0 Then FormatDimText = FormatDimText & "-"
            End If
            If Inches <> "0" Then
                FormatDimText = FormatDimText & Inches
                If Inchdecimal <> 0 Then
                    FormatDimText = FormatDimText & " " & Numerator & "/" & Denominator
                End If
                FormatDimText = FormatDimText & """"
            Else
                If Inchdecimal <> 0 Then
                    FormatDimText = FormatDimText & Numerator & "/" & Denominator
                End If
            End If

        End If

        On Error GoTo 0
        Exit Function
FormatDimTextError:
        FormatDimText = "error"
    End Function

    Sub WieghtGet()

        Dim i As Integer
        Dim linecnt As Integer
        Dim AnchType As String = ""

        sTrace = "WieghtGet: Enter" : LogDebug(sTrace)

        lngWieght(0) = 0
        lngWieght(1) = 0
        lngWieght(2) = 0
        lngWieght(3) = 0
        lngWieght(4) = 0
        lngWieght(5) = 0
        lngWieght(6) = 0
        lngWieght(7) = 0
        lngWieght(8) = 0
        lngWieght(9) = 0
        lngWieght(10) = 0
        lngWieght(11) = 0
        lngWieght(12) = 0

        On Error GoTo Err_WieghtGet

        sTrace = "WieghtGet: Locate anchorage in config" : LogDebug(sTrace)
        For i = 0 To intLineCount - 1
            If Left(AryConfigFile(i), 9) = "Anchorage" Then
                linecnt = Len(AryConfigFile(i))
                AnchType = Right(AryConfigFile(i), (linecnt - 10))
            End If
        Next i

        sTrace = "WieghtGet: Locate anchorage in database" : LogDebug(sTrace)
        Anchor.MoveFirst()
        Do While Anchor.EOF = False
            If Anchor.Fields("Model Number").Value.ToString = ModelNum Then
                If Anchor.Fields("Anchorage").Value.ToString = AnchType Then
                    If LCase(Anchor.Fields("Hand").Value.ToString) = LCase(Hand) Or LCase(Anchor.Fields("Hand").Value.ToString) = "none" Then Exit Do
                End If
            End If
            Anchor.MoveNext()
        Loop

        On Error GoTo 0

SteelSet:
        sTrace = "WieghtGet: Steel set" : LogDebug(sTrace)
        On Error Resume Next
        For i = 0 To Accesscnt - 1
            If Anchor.Fields("Plan").Value.ToString = "A" Then
                lngWieght(0) = CLng(Accessories(9, i)) + lngWieght(0)
                lngWieght(1) = CLng(Accessories(10, i)) + lngWieght(1)
                lngWieght(3) = CLng(Accessories(11, i)) + lngWieght(3)
                lngWieght(2) = CLng(Accessories(33, i)) + lngWieght(2)  'added 01/14/09
                lngWieght(4) = CLng(Accessories(12, i)) + lngWieght(4)
                lngWieght(5) = CLng(Accessories(13, i)) + lngWieght(5)
                lngWieght(6) = CLng(Accessories(14, i)) + lngWieght(6)
                lngWieght(7) = CLng(Accessories(15, i)) + lngWieght(7)
                lngWieght(8) = CLng(Accessories(16, i)) + lngWieght(8)
                lngWieght(9) = CLng(Accessories(17, i)) + lngWieght(9)
                lngWieght(10) = CLng(Accessories(18, i)) + lngWieght(10)
                lngWieght(11) = CLng(Accessories(19, i)) + lngWieght(11)
                lngWieght(12) = CLng(Accessories(20, i)) + lngWieght(12)
            End If
            If Anchor.Fields("Plan").Value.ToString = "B" Then
                lngWieght(0) = CLng(Accessories(9, i)) + lngWieght(0)
                lngWieght(1) = CLng(Accessories(10, i)) + lngWieght(1)
                lngWieght(2) = CLng(Accessories(33, i)) + lngWieght(2)  'added 01/14/09
                lngWieght(3) = CLng(Accessories(21, i)) + lngWieght(3)
                lngWieght(4) = CLng(Accessories(22, i)) + lngWieght(4)
                lngWieght(5) = CLng(Accessories(23, i)) + lngWieght(5)
                lngWieght(6) = CLng(Accessories(24, i)) + lngWieght(6)
                lngWieght(7) = CLng(Accessories(25, i)) + lngWieght(7)
                lngWieght(8) = CLng(Accessories(26, i)) + lngWieght(8)
                lngWieght(9) = CLng(Accessories(27, i)) + lngWieght(9)
                lngWieght(10) = CLng(Accessories(28, i)) + lngWieght(10)
                lngWieght(11) = CLng(Accessories(29, i)) + lngWieght(11)
                lngWieght(12) = CLng(Accessories(30, i)) + lngWieght(12)
            End If
        Next i
        'added 01/14/09
        If Anchor.Fields("Plan").Value.ToString = "A" Then
            lngWieght(0) = CLng(Drive.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(Drive.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(Drive.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(Drive.Fields.Item("Plan-A Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(Drive.Fields.Item("Plan-A Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(Drive.Fields.Item("Plan-A Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(Drive.Fields.Item("Plan-A Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(Drive.Fields.Item("Plan-A Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(Drive.Fields.Item("Plan-A Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(Drive.Fields.Item("Plan-A Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(Drive.Fields.Item("Plan-A Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(Drive.Fields.Item("Plan-A Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(Drive.Fields.Item("Plan-A Point 10").Value) + lngWieght(12)
        End If
        If Anchor.Fields("Plan").Value.ToString = "B" Then
            lngWieght(0) = CLng(Drive.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(Drive.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(Drive.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(Drive.Fields.Item("Plan-B Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(Drive.Fields.Item("Plan-B Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(Drive.Fields.Item("Plan-B Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(Drive.Fields.Item("Plan-B Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(Drive.Fields.Item("Plan-B Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(Drive.Fields.Item("Plan-B Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(Drive.Fields.Item("Plan-B Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(Drive.Fields.Item("Plan-B Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(Drive.Fields.Item("Plan-B Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(Drive.Fields.Item("Plan-B Point 10").Value) + lngWieght(12)
        End If
        If Anchor.Fields("Plan").Value.ToString = "A" Then
            lngWieght(0) = CLng(InConns.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(InConns.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(InConns.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(InConns.Fields.Item("Plan-A Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(InConns.Fields.Item("Plan-A Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(InConns.Fields.Item("Plan-A Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(InConns.Fields.Item("Plan-A Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(InConns.Fields.Item("Plan-A Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(InConns.Fields.Item("Plan-A Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(InConns.Fields.Item("Plan-A Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(InConns.Fields.Item("Plan-A Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(InConns.Fields.Item("Plan-A Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(InConns.Fields.Item("Plan-A Point 10").Value) + lngWieght(12)
        End If
        If Anchor.Fields("Plan").Value.ToString = "B" Then
            lngWieght(0) = CLng(InConns.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(InConns.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(InConns.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(InConns.Fields.Item("Plan-B Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(InConns.Fields.Item("Plan-B Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(InConns.Fields.Item("Plan-B Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(InConns.Fields.Item("Plan-B Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(InConns.Fields.Item("Plan-B Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(InConns.Fields.Item("Plan-B Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(InConns.Fields.Item("Plan-B Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(InConns.Fields.Item("Plan-B Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(InConns.Fields.Item("Plan-B Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(InConns.Fields.Item("Plan-B Point 10").Value) + lngWieght(12)
        End If
        If Anchor.Fields("Plan").Value.ToString = "A" Then
            lngWieght(0) = CLng(OutConns.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(OutConns.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(OutConns.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(OutConns.Fields.Item("Plan-A Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(OutConns.Fields.Item("Plan-A Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(OutConns.Fields.Item("Plan-A Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(OutConns.Fields.Item("Plan-A Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(OutConns.Fields.Item("Plan-A Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(OutConns.Fields.Item("Plan-A Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(OutConns.Fields.Item("Plan-A Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(OutConns.Fields.Item("Plan-A Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(OutConns.Fields.Item("Plan-A Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(OutConns.Fields.Item("Plan-A Point 10").Value) + lngWieght(12)
        End If
        If Anchor.Fields("Plan").Value.ToString = "B" Then
            lngWieght(0) = CLng(OutConns.Fields.Item("Shipping Weight").Value) + lngWieght(0)
            lngWieght(1) = CLng(OutConns.Fields.Item("Operating Weight").Value) + lngWieght(1)
            lngWieght(2) = CLng(OutConns.Fields.Item("Heaviest Section Add").Value) + lngWieght(2)
            lngWieght(3) = CLng(OutConns.Fields.Item("Plan-B Point 1").Value) + lngWieght(3)
            lngWieght(4) = CLng(OutConns.Fields.Item("Plan-B Point 2").Value) + lngWieght(4)
            lngWieght(5) = CLng(OutConns.Fields.Item("Plan-B Point 3").Value) + lngWieght(5)
            lngWieght(6) = CLng(OutConns.Fields.Item("Plan-B Point 4").Value) + lngWieght(6)
            lngWieght(7) = CLng(OutConns.Fields.Item("Plan-B Point 5").Value) + lngWieght(7)
            lngWieght(8) = CLng(OutConns.Fields.Item("Plan-B Point 6").Value) + lngWieght(8)
            lngWieght(9) = CLng(OutConns.Fields.Item("Plan-B Point 7").Value) + lngWieght(9)
            lngWieght(10) = CLng(OutConns.Fields.Item("Plan-B Point 8").Value) + lngWieght(10)
            lngWieght(11) = CLng(OutConns.Fields.Item("Plan-B Point 9").Value) + lngWieght(11)
            lngWieght(12) = CLng(OutConns.Fields.Item("Plan-B Point 10").Value) + lngWieght(12)
        End If
        'end add

        On Error GoTo Err_WieghtGet

        sTrace = "WieghtGet: Assign anchor fields" : LogDebug(sTrace)
        lngWieght(0) = CLng(Anchor.Fields("Shipping Wieght").Value) + lngWieght(0)
        lngWieght(1) = CLng(Anchor.Fields("Operating Wieght").Value) + lngWieght(1)
        lngWieght(2) = CLng(Anchor.Fields("Heaviest Section").Value) + lngWieght(2)
        lngWieght(3) = CLng(Anchor.Fields("Point 1").Value) + lngWieght(3)
        lngWieght(4) = CLng(Anchor.Fields("Point 2").Value) + lngWieght(4)
        lngWieght(5) = CLng(Anchor.Fields("Point 3").Value) + lngWieght(5)
        lngWieght(6) = CLng(Anchor.Fields("Point 4").Value) + lngWieght(6)
        lngWieght(7) = CLng(Anchor.Fields("Point 5").Value) + lngWieght(7)
        lngWieght(8) = CLng(Anchor.Fields("Point 6").Value) + lngWieght(8)
        lngWieght(9) = CLng(Anchor.Fields("Point 7").Value) + lngWieght(9)
        lngWieght(10) = CLng(Anchor.Fields("Point 8").Value) + lngWieght(10)
        lngWieght(11) = CLng(Anchor.Fields("Point 9").Value) + lngWieght(11)
        lngWieght(12) = CLng(Anchor.Fields("Point 10").Value) + lngWieght(12)

        sTrace = "WieghtGet: Exit" : LogDebug(sTrace)

        Exit Sub

Err_WieghtGet:

        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'WieghtGet


    Sub FlowSet()

        ' Looks for the Flowrate in the configuration file.
        ' If no flow rate is found a defualt of 100 GPM is assigned

        Dim i As Integer
        Dim linecnt As Integer

        sTrace = "FlowSet: Enter" : LogDebug(sTrace)

        On Error GoTo Err_FlowSet
        sTrace = "FlowSet: Locate flow rate in config" : LogDebug(sTrace)

        For i = 0 To intLineCount - 1
            If Left(AryConfigFile(i), 9) = "FlowRate=" Then
                linecnt = Len(AryConfigFile(i))
                If Right(AryConfigFile(i), (linecnt - 9)) <> "" Then
                    Flow = CLng(Right(AryConfigFile(i), (linecnt - 9)))
                Else
                    Flow = 100
                End If
            End If
        Next i

        sTrace = "FlowSet: Exit" : LogDebug(sTrace)

        Exit Sub
Err_FlowSet:

        sTrace = "Err_FlowSet: Enter" : LogDebug(sTrace)
        Call LogUnstructuredError("Error at " & sTrace, LOG_FATAL_ERROR)

    End Sub 'FlowSet


    Sub DataCompile()

        'Subroutine compiles all unit data from set databases

        Dim LineLen As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim EQMark As Integer
        Dim g, h, i, j, k, l As Integer
        Dim ConnDim As String
        Dim ConnLayerIn As String
        Dim ConnNotesIn As String
        Dim ConnLayerOut As String
        Dim ConnNotesOut As String
        Dim DwgIndex As Integer
        Dim IndexFound As Boolean

        sTrace = "DataCompile: Enter" : LogDebug(sTrace)

        On Error GoTo Err_DataCompile

        LineLen = 0
        ConnDim = ""
        ConnLayerIn = ""
        ConnNotesIn = ""
        ConnLayerOut = ""
        ConnNotesOut = ""

        DwgCount = 2
        ReDim DwgData(5, DwgCount - 1)

        'sets the drawing, layers and dims for the unit and steel print from the set databases
        sTrace = "DataCompile: Set drawing, layers and dims for the unit and steel print" : LogDebug(sTrace)
        If LCase(Data.Fields.Item("Unit Print File").Value.ToString) <> "none" Then DwgData(0, 0) = Data.Fields.Item("Unit Print File").Value.ToString
        If LCase(Data.Fields.Item("Layers On").Value.ToString) <> "none" Then DwgData(1, 0) = Data.Fields.Item("Layers On").Value.ToString
        If LCase(Data.Fields.Item("Unit Print Dims").Value.ToString) <> "none" Then DwgData(2, 0) = Data.Fields.Item("Unit Print Dims").Value.ToString
        If LCase(Data.Fields.Item("Notes").Value.ToString) <> "none" Then DwgData(3, 0) = Data.Fields.Item("Notes").Value.ToString
        If LCase(Anchor.Fields.Item("File").Value.ToString) <> "none" Then DwgData(0, 1) = Anchor.Fields.Item("File").Value.ToString
        If LCase(Anchor.Fields.Item("Layer").Value.ToString) <> "none" Then DwgData(1, 1) = Anchor.Fields.Item("Layer").Value.ToString
        If LCase(Anchor.Fields.Item("Dims").Value.ToString) <> "none" Then DwgData(2, 1) = Anchor.Fields.Item("Dims").Value.ToString
        If LCase(Anchor.Fields.Item("notes").Value.ToString) <> "none" Then DwgData(3, 1) = Anchor.Fields.Item("notes").Value.ToString


        LineLen = Len(InConns.Fields.Item("File").Value)

        'sets the data for the inlet connenction if the inlet is set for the unit print
        sTrace = "DataCompile: Set inlet conn if the inlet is set for the unit print" : LogDebug(sTrace)
        If LCase(InConns.Fields.Item("File").Value.ToString) = "unit" Then

            If LCase(InConns.Fields.Item("Dims-In").Value.ToString) <> "none" Then
                Call FlowAdjustConn(InConns.Fields("Dims-In").Value, InConns.Fields("Inlet").Value.ToString, ConnLayerIn, ConnNotesIn, ConnDim, CLng(InConns.Fields("Min-In").Value), CLng(InConns.Fields("Max-In").Value))
                DwgData(2, 0) = DwgData(2, 0) & ConnDim
            End If
            If LCase(InConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                DwgData(1, 0) = DwgData(1, 0) + ConnLayerIn
                If LCase(InConns.Fields("Layer Name").Value.ToString) <> "none" Then DwgData(1, 0) = DwgData(1, 0) + InConns.Fields("Layer Name").Value
            End If
            If LCase(InConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                DwgData(3, 0) = DwgData(3, 0) + ConnNotesIn
                If LCase(InConns.Fields("Notes").Value.ToString) <> "none" Then DwgData(3, 0) = DwgData(3, 0) + InConns.Fields("Notes").Value
            End If

        End If

        LineLen = Len(OutConns.Fields.Item("File").Value)
        'sets the data for the outlet connenction if the outlet is set for the unit print
        sTrace = "DataCompile: Set outlet conn if the outlet is set for the unit print" : LogDebug(sTrace)
        If LCase(OutConns.Fields.Item("File").Value.ToString) = "unit" Then

            If LCase(OutConns.Fields.Item("Dims-Out").Value.ToString) <> "none" Then
                Call FlowAdjustConn(OutConns.Fields("Dims-Out").Value, OutConns.Fields("Outlet").Value.ToString, ConnLayerOut, ConnNotesOut, ConnDim, CLng(OutConns.Fields("Min-Out").Value), CLng(OutConns.Fields("Max-Out").Value))
                DwgData(2, 0) = DwgData(2, 0) & ConnDim
            End If
            If LCase(OutConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                DwgData(1, 0) = DwgData(1, 0) + ConnLayerOut
                If LCase(OutConns.Fields("Layer Name").Value.ToString) <> "none" Then DwgData(1, 0) = DwgData(1, 0) + OutConns.Fields("Layer Name").Value
            End If
            If LCase(OutConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                DwgData(3, 0) = DwgData(3, 0) + ConnNotesOut
                If LCase(OutConns.Fields("Notes").Value.ToString) <> "none" Then DwgData(3, 0) = DwgData(3, 0) + OutConns.Fields("Notes").Value
            End If

        End If

        LineLen = Len(InConns.Fields.Item("File").Value)

        'sets the data for the inlet connenction if the inlet is other than the unit print
        sTrace = "DataCompile: Set inlet conn if the inlet is other than the unit print" : LogDebug(sTrace)
        If LCase(InConns.Fields.Item("File").Value.ToString) <> "unit" Then

            For h = 1 To LineLen
                If Mid(InConns.Fields.Item("File").Value.ToString, h, 1) = "[" Then bmark = h
                If Mid(InConns.Fields.Item("File").Value.ToString, h, 1) = "=" Then EQMark = h
                If Mid(InConns.Fields.Item("File").Value.ToString, h, 1) = "]" Then
                    emark = h
                    If LCase(Mid(InConns.Fields.Item("File").Value.ToString, bmark + 1, 4)) = "unit" Then

                        If LCase(InConns.Fields.Item("Dims-In").Value.ToString) <> "none" Then
                            ConnDim = GetDBdims(InConns.Fields("Dims-In").Value, InConns, "Dims-In", "unit")
                            Call FlowAdjustConn(ConnDim, InConns.Fields("Inlet").Value.ToString, ConnLayerIn, ConnNotesIn, ConnDim, InConns.Fields("Min-In").Value, InConns.Fields("Max-In").Value)
                            DwgData(2, 0) = DwgData(2, 0) & ConnDim
                        End If
                        ConnDim = ""
                        If LCase(InConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                            DwgData(1, 0) = DwgData(1, 0) + ConnLayerIn
                            ConnDim = GetDBdims(InConns.Fields("Layer Name").Value, InConns, "Layer Name", "unit")
                            DwgData(1, 0) = DwgData(1, 0) + ConnDim
                        End If
                        ConnDim = ""
                        If LCase(InConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                            DwgData(3, 0) = DwgData(3, 0) + ConnNotesIn
                            ConnDim = GetDBdims(InConns.Fields("Notes").Value, InConns, "Notes", "unit")
                            DwgData(3, 0) = DwgData(3, 0) + ConnDim
                        End If
                        ConnDim = ""

                    Else
                        If LCase(Mid(InConns.Fields.Item("File").Value.ToString, bmark + 1, 5)) = "steel" Then

                            If LCase(InConns.Fields.Item("Dims-In").Value.ToString) <> "none" Then
                                ConnDim = GetDBdims(InConns.Fields("Dims-In").Value, InConns, "Dims-In", "steel")
                                Call FlowAdjustConn(ConnDim, InConns.Fields("Inlet").Value.ToString, ConnLayerIn, ConnNotesIn, ConnDim, InConns.Fields("Min-In").Value, InConns.Fields("Max-In").Value)
                                DwgData(2, 1) = DwgData(2, 1) & ConnDim
                            End If
                            ConnDim = ""
                            If LCase(InConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, 1) = DwgData(1, 1) + ConnLayerIn
                                ConnDim = GetDBdims(InConns.Fields("Layer Name").Value, InConns, "Layer Name", "steel")
                                DwgData(1, 1) = DwgData(1, 1) + ConnDim
                            End If
                            ConnDim = ""
                            If LCase(InConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, 1) = DwgData(3, 1) + ConnNotesIn
                                ConnDim = GetDBdims(InConns.Fields("Notes").Value, InConns, "Notes", "steel")
                                DwgData(3, 1) = DwgData(3, 1) + ConnDim
                            End If
                            ConnDim = ""
                        Else
                            IndexFound = False
                            For g = 0 To DwgCount - 1
                                If Mid(InConns.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1) = DwgData(0, g) Then
                                    DwgIndex = g
                                    IndexFound = True
                                End If
                            Next g

                            If IndexFound = False Then
                                DwgCount = DwgCount + 1
                                DwgIndex = DwgCount - 1
                            End If

                            ReDim Preserve DwgData(5, DwgCount - 1)
                            DwgData(0, DwgIndex) = Mid(InConns.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1)
                            DwgData(5, DwgIndex) = GetDBdims(InConns.Fields("Prefix").Value, InConns, "Prefix", Mid(InConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            If LCase(InConns.Fields.Item("Dims-In").Value.ToString) <> "none" Then
                                ConnDim = GetDBdims(InConns.Fields("Dims-In").Value, InConns, "Dims-In", Mid(InConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                Call FlowAdjustConn(ConnDim, InConns.Fields("Inlet").Value.ToString, ConnLayerIn, ConnNotesIn, ConnDim, InConns.Fields("Min-In").Value, InConns.Fields("Max-In").Value)
                                DwgData(2, DwgIndex) = DwgData(2, DwgIndex) & ConnDim
                            End If
                            ConnDim = ""
                            If LCase(InConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) + ConnLayerIn
                                ConnDim = GetDBdims(InConns.Fields("Layer Name").Value, InConns, "Layer Name", Mid(InConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) + ConnDim
                            End If
                            ConnDim = ""
                            If LCase(InConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) + ConnNotesIn
                                ConnDim = GetDBdims(InConns.Fields("Notes").Value, InConns, "Notes", Mid(InConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) + ConnDim
                            End If
                        End If
                    End If
                End If
            Next h
        End If

        LineLen = Len(OutConns.Fields.Item("File").Value)

        'sets the data for the outlet connenction if the outlet is other than the unit print
        sTrace = "DataCompile: Set outlet conn if the outlet is other than the unit print" : LogDebug(sTrace)
        If LCase(OutConns.Fields.Item("File").Value.ToString) <> "unit" Then

            For h = 1 To LineLen
                If Mid(OutConns.Fields.Item("File").Value.ToString, h, 1) = "[" Then bmark = h
                If Mid(OutConns.Fields.Item("File").Value.ToString, h, 1) = "=" Then EQMark = h
                If Mid(OutConns.Fields.Item("File").Value.ToString, h, 1) = "]" Then
                    emark = h
                    If LCase(Mid(OutConns.Fields.Item("File").Value.ToString, bmark + 1, 4)) = "unit" Then

                        If LCase(OutConns.Fields.Item("Dims-Out").Value.ToString) <> "none" Then
                            ConnDim = GetDBdims(OutConns.Fields("Dims-Out").Value, OutConns, "Dims-Out", "unit")
                            Call FlowAdjustConn(ConnDim, OutConns.Fields("Outlet").Value.ToString, ConnLayerOut, ConnNotesOut, ConnDim, OutConns.Fields("Min-Out").Value, OutConns.Fields("Max-Out").Value)
                            DwgData(2, 0) = DwgData(2, 0) & ConnDim
                        End If
                        ConnDim = ""
                        If LCase(OutConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                            DwgData(1, 0) = DwgData(1, 0) + ConnLayerOut
                            ConnDim = GetDBdims(OutConns.Fields("Layer Name").Value, OutConns, "Layer Name", "unit")
                            DwgData(1, 0) = DwgData(1, 0) + ConnDim
                        End If
                        ConnDim = ""
                        If LCase(OutConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                            DwgData(3, 0) = DwgData(3, 0) + ConnNotesIn + ConnNotesOut
                            ConnDim = GetDBdims(OutConns.Fields("Notes").Value, OutConns, "Notes", "unit")
                            DwgData(3, 0) = DwgData(3, 0) + ConnDim
                        End If
                        ConnDim = ""

                    Else
                        If LCase(Mid(OutConns.Fields.Item("File").Value.ToString, bmark + 1, 5)) = "steel" Then

                            If LCase(OutConns.Fields.Item("Dims-Out").Value.ToString) <> "none" Then
                                ConnDim = GetDBdims(OutConns.Fields("Dims-Out").Value, OutConns, "Dims-Out", "steel")
                                Call FlowAdjustConn(ConnDim, OutConns.Fields("Outlet").Value.ToString, ConnLayerOut, ConnNotesOut, ConnDim, OutConns.Fields("Min-Out").Value, OutConns.Fields("Max-Out").Value)
                                DwgData(2, 1) = DwgData(2, 1) & ConnDim
                            End If
                            ConnDim = ""
                            If LCase(OutConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, 1) = DwgData(1, 1) + ConnLayerIn + ConnLayerOut
                                ConnDim = GetDBdims(OutConns.Fields("Layer Name").Value, OutConns, "Layer Name", "steel")
                                DwgData(1, 1) = DwgData(1, 1) + ConnDim
                            End If
                            ConnDim = ""
                            If LCase(OutConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, 1) = DwgData(3, 1) + ConnNotesOut
                                ConnDim = GetDBdims(OutConns.Fields("Notes").Value, OutConns, "Notes", "steel")
                                DwgData(3, 1) = DwgData(3, 1) + ConnDim
                            End If
                            ConnDim = ""
                        Else
                            IndexFound = False
                            For g = 0 To DwgCount - 1
                                If Mid(OutConns.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1) = DwgData(0, g) Then
                                    DwgIndex = g
                                    IndexFound = True
                                End If
                            Next g

                            If IndexFound = False Then
                                DwgCount = DwgCount + 1
                                DwgIndex = DwgCount - 1
                            End If
                            ReDim Preserve DwgData(5, DwgCount - 1)
                            DwgData(0, DwgIndex) = Mid(OutConns.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1)
                            DwgData(5, DwgIndex) = GetDBdims(OutConns.Fields("Prefix").Value, OutConns, "Prefix", Mid(OutConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            ConnDim = ""
                            If LCase(OutConns.Fields.Item("Dims-Out").Value.ToString) <> "none" Then
                                ConnDim = GetDBdims(OutConns.Fields("Dims-Out").Value, OutConns, "Dims-Out", Mid(OutConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                Call FlowAdjustConn(ConnDim, OutConns.Fields("Outlet").Value.ToString, ConnLayerOut, ConnNotesOut, ConnDim, OutConns.Fields("Min-Out").Value, OutConns.Fields("Max-Out").Value)
                                DwgData(2, DwgIndex) = DwgData(2, DwgIndex) & ConnDim
                            End If
                            ConnDim = ""
                            If LCase(OutConns.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) + ConnLayerOut
                                ConnDim = GetDBdims(OutConns.Fields("Layer Name").Value, OutConns, "Layer Name", Mid(OutConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) + ConnDim
                            End If
                            ConnDim = ""
                            If LCase(OutConns.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) + ConnNotesIn + ConnNotesOut
                                ConnDim = GetDBdims(OutConns.Fields("Notes").Value, OutConns, "Notes", Mid(OutConns.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) + ConnDim
                            End If
                        End If
                    End If
                End If
            Next h
        End If

        'sets the data for the drive if the drive is on the unit print
        sTrace = "DataCompile: Set drive drive if the drive is on the unit print" : LogDebug(sTrace)
        If LCase(Drive.Fields.Item("File").Value.ToString) = "unit" Then
            If LCase(Drive.Fields.Item("Dims").Value.ToString) <> "none" Then
                DwgData(2, 0) = DwgData(2, 0) + GetDBdims(Drive.Fields("Dims").Value, Drive, "Dims", "unit")
            End If
            If LCase(Drive.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                DwgData(1, 0) = DwgData(1, 0) + GetDBdims(Drive.Fields("Layer Name").Value, Drive, "Layer Name", "unit")
            End If
            If LCase(Drive.Fields.Item("Notes").Value.ToString) <> "none" Then
                DwgData(3, 0) = DwgData(3, 0) + GetDBdims(Drive.Fields("Notes").Value, Drive, "Notes", "unit")
            End If
        End If

        'sets the data for the drive if the drive is other than unit print
        sTrace = "DataCompile: Set drive drive if the drive is other than the unit print" : LogDebug(sTrace)
        LineLen = Len(Drive.Fields.Item("File").Value)
        If LCase(Drive.Fields.Item("File").Value.ToString) <> "unit" Then
            For h = 1 To LineLen
                If Mid(Drive.Fields.Item("File").Value.ToString, h, 1) = "[" Then bmark = h
                If Mid(Drive.Fields.Item("File").Value.ToString, h, 1) = "=" Then EQMark = h
                If Mid(Drive.Fields.Item("File").Value.ToString, h, 1) = "]" Then
                    emark = h
                    If LCase(Mid(Drive.Fields.Item("File").Value.ToString, bmark + 1, 4)) = "unit" Then

                        If LCase(Drive.Fields.Item("Dims").Value.ToString) <> "none" Then
                            DwgData(2, 0) = DwgData(2, 0) & GetDBdims(Drive.Fields("Dims").Value, Drive, "Dims", "unit")
                        End If
                        If LCase(Drive.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                            DwgData(1, 0) = DwgData(1, 0) + GetDBdims(Drive.Fields("Layer Name").Value, Drive, "Layer Name", "unit")
                        End If
                        If LCase(Drive.Fields.Item("Notes").Value.ToString) <> "none" Then
                            DwgData(3, 0) = DwgData(3, 0) + GetDBdims(Drive.Fields("Notes").Value, Drive, "Notes", "unit")
                        End If
                    Else
                        If LCase(Mid(Drive.Fields.Item("File").Value.ToString, bmark + 1, 5)) = "steel" Then

                            If LCase(Drive.Fields.Item("Dims").Value.ToString) <> "none" Then
                                DwgData(2, 0) = DwgData(2, 0) & GetDBdims(Drive.Fields("Dims").Value, Drive, "Dims", "steel")
                            End If
                            If LCase(Drive.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, 0) = DwgData(1, 0) + GetDBdims(Drive.Fields("Layer Name").Value, Drive, "Layer Name", "steel")
                            End If
                            If LCase(Drive.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, 0) = DwgData(3, 0) + GetDBdims(Drive.Fields("Notes").Value, Drive, "Notes", "steel")
                            End If
                        Else
                            IndexFound = False
                            For g = 0 To DwgCount - 1
                                If Mid(Drive.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1) = DwgData(0, g) Then
                                    DwgIndex = g
                                    IndexFound = True
                                End If
                            Next g

                            If IndexFound = False Then
                                DwgCount = DwgCount + 1
                                DwgIndex = DwgCount - 1
                            End If

                            ReDim Preserve DwgData(5, DwgCount - 1)
                            DwgData(0, DwgIndex) = Mid(Drive.Fields.Item("File").Value.ToString, EQMark + 1, emark - EQMark - 1)
                            DwgData(5, DwgIndex) = GetDBdims(Drive.Fields("Prefix").Value, Drive, "Prefix", Mid(Drive.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            If LCase(Drive.Fields.Item("Dims").Value.ToString) <> "none" Then
                                DwgData(2, DwgIndex) = DwgData(2, DwgIndex) & GetDBdims(Drive.Fields("Dims").Value, Drive, "Dims", Mid(Drive.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            End If
                            If LCase(Drive.Fields.Item("Layer Name").Value.ToString) <> "none" Then
                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) & GetDBdims(Drive.Fields("Layer Name").Value, Drive, "Layer Name", Mid(Drive.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            End If
                            If LCase(Drive.Fields.Item("Notes").Value.ToString) <> "none" Then
                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) & GetDBdims(Drive.Fields("Notes").Value, Drive, "Notes", Mid(Drive.Fields("File").Value.ToString, bmark + 1, EQMark - 1 - bmark))
                            End If
                        End If
                    End If
                End If
            Next h
        End If

        'runs through the accesory array and gathers the data for accesories
        sTrace = "DataCompile: Set data for accessories" : LogDebug(sTrace)
        For i = 0 To Accesscnt - 1
            LineLen = Len(Accessories(5, i))
            If Accessories(0, i) <> "none" Then
                If Accessories(8, i) = True Then Call FlowAdjustAcc(i)
                If LCase(Accessories(5, i)) = "unit" Then
                    If LCase(Accessories(6, i)) <> "none" Then DwgData(1, 0) = DwgData(1, 0) & getdims(i, 6, "unit")
                    If LCase(Accessories(31, i)) <> "none" Then DwgData(2, 0) = DwgData(2, 0) & getdims(i, 31, "unit")
                    If LCase(Accessories(7, i)) <> "none" Then DwgData(3, 0) = DwgData(3, 0) & getdims(i, 7, "unit")
                    If LCase(Accessories(2, i)) <> "none" Then DwgData(4, 0) = DwgData(4, 0) & getdims(i, 2, "unit")
                End If
                If LCase(Accessories(5, i)) = "steel" Then
                    If LCase(Accessories(6, i)) <> "none" Then DwgData(1, 1) = DwgData(1, 1) & getdims(i, 6, "steel")
                    If LCase(Accessories(31, i)) <> "none" Then DwgData(2, 1) = DwgData(2, 1) & getdims(i, 31, "steel")
                    If LCase(Accessories(7, i)) <> "none" Then DwgData(3, 1) = DwgData(3, 1) & getdims(i, 7, "steel")
                    If LCase(Accessories(2, i)) <> "none" Then DwgData(4, 1) = DwgData(4, 1) & getdims(i, 2, "steel")
                End If
                If LCase(Accessories(5, i)) = "[unit]" Then
                    If LCase(Accessories(6, i)) <> "none" Then DwgData(1, 0) = DwgData(1, 0) & getdims(i, 6, "unit")
                    If LCase(Accessories(31, i)) <> "none" Then DwgData(2, 0) = DwgData(2, 0) & getdims(i, 31, "unit")
                    If LCase(Accessories(7, i)) <> "none" Then DwgData(3, 0) = DwgData(3, 0) & getdims(i, 7, "unit")
                    If LCase(Accessories(2, i)) <> "none" Then DwgData(4, 0) = DwgData(4, 0) & getdims(i, 2, "unit")
                End If
                If LCase(Accessories(5, i)) = "[steel]" Then
                    If LCase(Accessories(6, i)) <> "none" Then DwgData(1, 1) = DwgData(1, 1) & getdims(i, 6, "steel")
                    If LCase(Accessories(31, i)) <> "none" Then DwgData(2, 1) = DwgData(2, 1) & getdims(i, 31, "steel")
                    If LCase(Accessories(7, i)) <> "none" Then DwgData(3, 1) = DwgData(3, 1) & getdims(i, 7, "steel")
                    If LCase(Accessories(2, i)) <> "none" Then DwgData(4, 1) = DwgData(4, 1) & getdims(i, 2, "steel")
                End If
                If LCase(Accessories(5, i)) <> "unit" Then
                    If LCase(Accessories(5, i)) <> "[unit]" Then
                        If LCase(Accessories(5, i)) <> "steel" Then
                            If LCase(Accessories(5, i)) <> "[steel]" Then
                                For h = 1 To LineLen
                                    If Mid(Accessories(5, i), h, 1) = "[" Then bmark = h
                                    If Mid(Accessories(5, i), h, 1) = "=" Then EQMark = h
                                    If Mid(Accessories(5, i), h, 1) = "]" Then
                                        emark = h
                                        If LCase(Mid(Accessories(5, i), bmark + 1, 4)) = "unit" Then
                                            DwgData(1, 0) = DwgData(1, 0) & getdims(i, 6, "unit")
                                            DwgData(2, 0) = DwgData(2, 0) & getdims(i, 31, "unit")
                                            DwgData(3, 0) = DwgData(3, 0) & getdims(i, 7, "unit")
                                            DwgData(4, 0) = DwgData(4, 0) & getdims(i, 2, "unit")
                                        Else
                                            If LCase(Mid(Accessories(5, i), bmark + 1, 5)) = "steel" Then
                                                DwgData(1, 1) = DwgData(1, 1) & getdims(i, 6, "unit")
                                                DwgData(2, 1) = DwgData(2, 1) & getdims(i, 31, "unit")
                                                DwgData(3, 1) = DwgData(3, 1) & getdims(i, 7, "unit")
                                                DwgData(4, 1) = DwgData(4, 1) & getdims(i, 2, "unit")
                                            Else
                                                IndexFound = False
                                                For g = 0 To DwgCount - 1
                                                    If Mid(Accessories(5, i), EQMark + 1, emark - EQMark - 1) = DwgData(0, g) Then
                                                        DwgIndex = g
                                                        IndexFound = True
                                                    End If
                                                Next g

                                                If IndexFound = False Then
                                                    DwgCount = DwgCount + 1
                                                    DwgIndex = DwgCount - 1
                                                End If
                                                ReDim Preserve DwgData(5, DwgCount - 1)
                                                DwgData(0, DwgIndex) = Mid(Accessories(5, i), EQMark + 1, emark - EQMark - 1)
                                                DwgData(1, DwgIndex) = DwgData(1, DwgIndex) & getdims(i, 6, Mid(Accessories(5, i), bmark + 1, EQMark - 1 - bmark))
                                                DwgData(2, DwgIndex) = DwgData(2, DwgIndex) & getdims(i, 31, Mid(Accessories(5, i), bmark + 1, EQMark - 1 - bmark))
                                                DwgData(3, DwgIndex) = DwgData(3, DwgIndex) & getdims(i, 7, Mid(Accessories(5, i), bmark + 1, EQMark - 1 - bmark))
                                                DwgData(4, DwgIndex) = DwgData(4, DwgIndex) & getdims(i, 2, Mid(Accessories(5, i), bmark + 1, EQMark - 1 - bmark))
                                                DwgData(5, DwgIndex) = getdims(i, 32, Mid(Accessories(5, i), bmark + 1, EQMark - 1 - bmark))
                                            End If
                                        End If
                                    End If
                                Next h
                            End If
                        End If
                    End If
                End If
            End If
        Next i

        For i = 0 To Accesscnt - 1
            For j = 0 To 32
                Accessories(j, i) = ""
            Next j
        Next i

        sTrace = "DataCompile: Exit" : LogDebug(sTrace)

        Exit Sub

Err_DataCompile:

        Call LogUnstructuredError("Error at " & sTrace, LOG_ERROR)

    End Sub 'DataCompile


    'Sub FlowAdjustConn(StrDimIn As Object, Conntype As String, strLayerOut As String, strNotesOut As String, strDimOut As Object, MinConn As Long, MaxConn As Long)   3.0.2.40
    Sub FlowAdjustConn(StrDimIn As Object, Conntype As String, strLayerOut As String, strNotesOut As String, ByRef strDimOut As Object, MinConn As Long, MaxConn As Long)

        Dim Alinelen As Integer
        Dim h, i As Integer
        Dim lenAdjust As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim EQMark As Integer
        Dim ComMark As Integer
        Dim FlowFound As Boolean
        Dim TempText As String
        Dim TempDim As String
        Dim TempUnit As String
        Dim lngtempdim As Double
        Dim tempdimlen As Integer
        Dim lenTempdim As Integer
        Dim lentempdim2 As Integer
        Dim AdjustDim As String

        Dim lngMinFlow As Long
        Dim lngmaxFlow As Long
        Dim lngSelectedConn As Long

        sTrace = "FlowAdjustConn: Enter" : LogDebug(sTrace)
        sTrace = "FlowAdjustConn: StrDimIn = " & StrDimIn : LogDebug(sTrace)

        TempDim = ""
        TempUnit = ""
        AdjustDim = ""
        FlowFound = False
        Alinelen = Len(StrDimIn)

        ' This subroutine adjusts the dimension for dimension labeled flow sensitive
        On Error GoTo Err_FlowAdjustConn

        sTrace = "FlowAdjustConn: Parse StrDimIn" : LogDebug(sTrace)
        For i = 1 To Alinelen
            If Mid(StrDimIn, i, 1) = "[" Then bmark = i
            If Mid(StrDimIn, i, 1) = "," Then ComMark = i
            If Mid(StrDimIn, i, 1) = "=" Then EQMark = i
            If LCase(Mid(StrDimIn, bmark + 1, 6)) = "*flow*" Then FlowFound = True
            If Right(StrDimIn, 1) = "]" Then
                If i + lenAdjust = Len(StrDimIn) And FlowFound = True Then GoTo NextStep
            End If
            If Mid(StrDimIn, i, 1) = "]" And FlowFound = True Then
NextStep:
                sTrace = "FlowAdjustConn: NextStep" : LogDebug(sTrace)
                emark = i
                TempDim = Mid(StrDimIn, EQMark + 1, emark - EQMark - 1)
                tempdimlen = Len(TempDim)
                For h = 1 To tempdimlen
                    If Mid(TempDim, h, 1) Like "#" = False Then
                        lngtempdim = Left(TempDim, h - 1)
                        TempUnit = Right(TempDim, tempdimlen - (h - 1))
                        GoTo UnitStripEnd
                    End If
                Next h
UnitStripEnd:
                sTrace = "FlowAdjustConn: UnitStripEnd Loop1" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    If LCase(FlowDB.Fields("Connection").Value.ToString) = LCase(Conntype) Then
                        If LCase(FlowDB.Fields("WebDesc").Value.ToString) = "inlet" Then
                            If FlowDB.Fields("Size").Value.ToString = InletSize Then GoTo FlowEnd
                        End If

                        If LCase(FlowDB.Fields("WebDesc").Value.ToString) = "outlet" Then
                            If FlowDB.Fields("Size").Value = OutletSize Then GoTo FlowEnd
                        End If
                    End If
                    FlowDB.MoveNext()
                Loop

                sTrace = "FlowAdjustConn: UnitStripEnd Loop2" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    lngMinFlow = FlowDB.Fields("Min-Flow").Value
                    lngmaxFlow = FlowDB.Fields("max-flow").Value
                    If LCase(FlowDB.Fields("Connection").Value.ToString) = LCase(Conntype) Then
                        If lngMinFlow < Flow Then
                            If lngmaxFlow > Flow Then
                                lngSelectedConn = FlowDB.Fields("Size").Value
                                If FlowDB.Fields("Size").Value < MaxConn Then
                                    If FlowDB.Fields("Size").Value > MinConn Or FlowDB.Fields("Size").Value = MinConn Then
                                        GoTo FlowEnd
                                    End If
                                End If
                            End If
                        End If
                    End If
                    FlowDB.MoveNext()
                Loop

                If lngSelectedConn < MinConn Then lngSelectedConn = MinConn
                If lngSelectedConn > MaxConn Then lngSelectedConn = MaxConn

                sTrace = "FlowAdjustConn: UnitStripEnd Loop3" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    If LCase(FlowDB.Fields("Connection").Value.ToString) = LCase(Conntype) Then
                        If FlowDB.Fields("Size").Value = lngSelectedConn Then GoTo FlowEnd
                    End If
                    FlowDB.MoveNext()
                Loop

FlowBypass:
                sTrace = "FlowAdjustConn: FlowBypass" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    If LCase(FlowDB.Fields("Connection").Value.ToString) = LCase(Conntype) Then
                        If FlowDB.Fields("Size").Value Then GoTo FlowEnd
                    End If
                    FlowDB.MoveNext()
                Loop
FlowEnd:
                sTrace = "FlowAdjustConn: FlowEnd" : LogDebug(sTrace)
                lenTempdim = 0
                lentempdim2 = 0
                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "x" Then
                    AdjustDim = lngtempdim + FlowDB.Fields("X-Change").Value & TempUnit
                    lenTempdim = Len(AdjustDim)
                    lentempdim2 = Len(TempDim)
                End If

                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "y" Then
                    AdjustDim = lngtempdim + FlowDB.Fields("Y-Change").Value & TempUnit
                    lentempdim2 = Len(TempDim)
                    lenTempdim = Len(AdjustDim)
                End If

                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "s" Then
                    AdjustDim = FlowDB.Fields("Size").Value & TempUnit
                    lentempdim2 = Len(TempDim)
                    lenTempdim = Len(AdjustDim)
                End If

                TempText = Left(StrDimIn, EQMark) & AdjustDim & Right(StrDimIn, Alinelen - (emark - 1))
                StrDimIn = TempText
                FlowFound = False
                If LCase(FlowDB.Fields("Layer").Value.ToString) <> "none" Then strLayerOut = FlowDB.Fields("layer").Value.ToString
                If LCase(FlowDB.Fields("Notes").Value.ToString) <> "none" Then strLayerOut = FlowDB.Fields("Notes").Value.ToString
                If lenTempdim > lentempdim2 Then i = i + (lenTempdim - lentempdim2)
                If lenTempdim < lentempdim2 Then i = i - (lentempdim2 - lenTempdim)
                lenAdjust = lenTempdim - lentempdim2
                Alinelen = Len(StrDimIn)
            End If
        Next i
        strDimOut = StrDimIn

        sTrace = "FlowAdjustConn: strDimOut = " & strDimOut : LogDebug(sTrace)
        sTrace = "FlowAdjustConn: Exit" : LogDebug(sTrace)

        Exit Sub

Err_FlowAdjustConn:

        Call LogUnstructuredError("Error at " & sTrace & " while trying to retrive flow data on " & StrDimIn & ". Please check DB values.", LOG_FATAL_ERROR)

    End Sub 'FlowAdjustConn



    Sub FlowAdjustAcc(AccIndex As Object)

        Dim StrDimIn As String
        Dim Alinelen As Integer
        Dim lenAdjust As Integer
        Dim h, i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim EQMark As Integer
        Dim ComMark As Integer
        Dim FlowFound As Boolean
        Dim TempText As String
        Dim TempDim As String
        Dim TempUnit As String
        Dim lngtempdim As Double
        Dim tempdimlen As Integer

        Dim strLayerOut As String
        Dim AdjustDim As String
        Dim lenTempdim As Integer     'was not declared but is used in this scope. not sure.
        Dim lentempdim2 As Integer    'was not declared but is used in this scope. not sure.

        sTrace = "FlowAdjustAcc: Enter" : LogDebug(sTrace)

        FlowFound = False
        TempUnit = ""
        AdjustDim = ""
        StrDimIn = Accessories(31, AccIndex)
        Alinelen = Len(StrDimIn)


        On Error GoTo Err_FlowAdjustAcc
        For i = 1 To Alinelen
            If Mid(StrDimIn, i, 1) = "[" Then bmark = i
            If Mid(StrDimIn, i, 1) = "," Then ComMark = i
            If Mid(StrDimIn, i, 1) = "=" Then EQMark = i
            If LCase(Mid(StrDimIn, bmark + 1, 6)) = "*flow*" Then FlowFound = True
            If Right(StrDimIn, 1) = "]" Then
                If i + lenAdjust = Len(StrDimIn) And FlowFound = True Then GoTo NextStep
            End If
            If Mid(StrDimIn, i, 1) = "]" And FlowFound = True Then
NextStep:
                sTrace = "FlowAdjustAcc: NextStep Loop1" : LogDebug(sTrace)
                emark = i
                TempDim = Mid(StrDimIn, EQMark + 1, emark - EQMark - 1)
                tempdimlen = Len(TempDim)
                For h = 1 To tempdimlen
                    If Mid(TempDim, h, 1) Like "#" = False Then
                        lngtempdim = Left(TempDim, h - 1)
                        TempUnit = Right(TempDim, tempdimlen - (h - 1))
                        Exit For
                    End If
                Next h
                sTrace = "FlowAdjustAcc: NextStep Loop2" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    If LCase(FlowDB.Fields("Connection").Value) = LCase(Accessories(1, AccIndex)) Then
                        If LCase(FlowDB.Fields("WebDesc").Value) = "equalizer" Then
                            If FlowDB.Fields("Size").Value = EqSize Then GoTo FlowEnd
                        End If

                        If LCase(FlowDB.Fields("WebDesc").Value) = "bypass" Then
                            If FlowDB.Fields("Size").Value = BySize Then GoTo FlowEnd
                        End If
                    End If
                    FlowDB.MoveNext()
                Loop
                sTrace = "FlowAdjustAcc: NextStep Loop3" : LogDebug(sTrace)
                FlowDB.MoveFirst()
                Do While FlowDB.EOF = False
                    If LCase(FlowDB.Fields("Connection").Value) = LCase(Accessories(1, AccIndex)) Then
                        If FlowDB.Fields("Min-Flow").Value < Flow Then
                            If FlowDB.Fields("max-flow").Value > Flow Or FlowDB.Fields("max-flow").Value = Flow Then
                                GoTo FlowEnd
                            End If
                        End If
                    End If
                    FlowDB.MoveNext()
                Loop
FlowEnd:
                sTrace = "FlowAdjustAcc: FlowEnd" : LogDebug(sTrace)
                lenTempdim = 0
                lentempdim2 = 0
                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "x" Then
                    AdjustDim = lngtempdim + FlowDB.Fields("X-Change").Value & TempUnit
                    lenTempdim = Len(AdjustDim)
                    lentempdim2 = Len(TempDim)
                End If

                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "y" Then
                    AdjustDim = lngtempdim + FlowDB.Fields("Y-Change").Value & TempUnit
                    lentempdim2 = Len(TempDim)
                    lenTempdim = Len(AdjustDim)
                End If

                If Mid(LCase(StrDimIn), ComMark + 1, 1) = "s" Then
                    AdjustDim = FlowDB.Fields("Size").Value & TempUnit
                    lentempdim2 = Len(TempDim)
                    lenTempdim = Len(AdjustDim)
                End If

                TempText = Left(StrDimIn, EQMark) & AdjustDim & Right(StrDimIn, Alinelen - (emark - 1))
                StrDimIn = TempText
                Alinelen = Len(StrDimIn)
                FlowFound = False
                If LCase(FlowDB.Fields("Layer").Value) <> "none" Then strLayerOut = FlowDB.Fields("layer").Value
                If LCase(FlowDB.Fields("Notes").Value) <> "none" Then strLayerOut = FlowDB.Fields("Notes").Value
                If lenTempdim > lentempdim2 Then i = i + (lenTempdim - lentempdim2)
                If lenTempdim < lentempdim2 Then i = i - (lentempdim2 - lenTempdim)
                lenAdjust = lenTempdim - lentempdim2
            End If
        Next i
        Accessories(31, AccIndex) = StrDimIn

        sTrace = "FlowAdjustAcc: Exit" : LogDebug(sTrace)

        Exit Sub

Err_FlowAdjustAcc:

        Call LogUnstructuredError("Error at " & sTrace & " " & StrDimIn, LOG_ERROR)

    End Sub 'FlowAdjustAcc



    Function GetDBdims(Textin As Object, DBin As ADODB.Recordset, Field As String, Fileno As String) As String

        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim Alinelen As Integer
        Dim FileNolen As Integer
        Dim DBcombine As Object
        Dim sRet As String

        sTrace = "GetDBdims: Enter" : LogDebug(sTrace)
        sTrace = "GetDBdims: Textin = " & Textin : LogDebug(sTrace)

        On Error GoTo Err_GetDBdims

        sRet = ""

        DBcombine = DBin.Fields(Field).Value
        FileNolen = Len(Fileno)

        'parses out the db dim entries to get the dimensional data.
        Alinelen = Len(DBcombine)

        For i = 1 To Alinelen
            If Mid(DBcombine, i, 1) = "[" Then bmark = i
            If Mid(DBcombine, i, 1) = "]" Then
                emark = i
                If LCase(Mid((Mid(DBcombine, bmark, emark - bmark)), 2, FileNolen)) = Fileno Then
                    If LCase(Mid(DBcombine, bmark + FileNolen + 2, (emark - bmark - FileNolen - 2))) <> "none" Then
                        'GetDBdims = GetDBdims & "[" & Mid(DBcombine, bmark + FileNolen + 2, (emark - bmark - FileNolen - 2)) & "]"
                        sRet = sRet & "[" & Mid(DBcombine, bmark + FileNolen + 2, (emark - bmark - FileNolen - 2)) & "]"
                    End If
                End If
                If LCase(Mid((Mid(DBcombine, bmark, emark - bmark)), 2, FileNolen + 6)) = "*flow*" & Fileno Then
                    If LCase(Mid(DBcombine, bmark + FileNolen + 2, (emark - bmark - FileNolen - 2))) <> "none" Then
                        'GetDBdims = GetDBdims & "[*flow*" & Mid(DBcombine, bmark + FileNolen + 8, (emark - bmark - FileNolen - 8)) & "]"
                        sRet = sRet & "[*flow*" & Mid(DBcombine, bmark + FileNolen + 8, (emark - bmark - FileNolen - 8)) & "]"
                    End If
                End If
            End If
        Next i

        GetDBdims = sRet

        sTrace = "GetDBdims: sRet = " & sRet : LogDebug(sTrace)
        sTrace = "GetDBdims: Exit" : LogDebug(sTrace)

        Exit Function

Err_GetDBdims:

        Call LogUnstructuredError("Error at " & sTrace & " " & Textin, LOG_FATAL_ERROR)

    End Function 'GetDBdims



    Function getdims(AccesNo As Object, CatNo As Object, Fileno As String) As String

        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim Alinelen As Integer
        Dim FileNolen As Integer
        Dim sRet As String

        sTrace = "getdims: Enter" : LogDebug(sTrace)

        sRet = ""
        FileNolen = Len(Fileno)

        'parses out the db dim entries to get the dimensional data.
        On Error GoTo Err_getdims
        Alinelen = Len(Accessories(CatNo, AccesNo))
        sTrace = "getdims: Accessories(" & CatNo & ", " & AccesNo & ") = " & Accessories(CatNo, AccesNo) : LogDebug(sTrace)

        For i = 1 To Alinelen
            If Mid(Accessories(CatNo, AccesNo), i, 1) = "[" Then bmark = i
            If Mid(Accessories(CatNo, AccesNo), i, 1) = "]" Then
                emark = i
                If LCase(Mid((Mid(Accessories(CatNo, AccesNo), bmark, emark - bmark)), 2, FileNolen)) = Fileno Then
                    If LCase(Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 2, (emark - bmark - FileNolen - 2))) <> "none" Then
                        'getdims = getdims & "[" & Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 2, (emark - bmark - FileNolen - 2)) & "]"
                        sRet = sRet & "[" & Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 2, (emark - bmark - FileNolen - 2)) & "]"
                    End If
                End If
                If LCase(Mid((Mid(Accessories(CatNo, AccesNo), bmark, emark - bmark)), 2, FileNolen + 6)) = "*flow*" & Fileno Then
                    If LCase(Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 8, (emark - bmark - FileNolen - 8))) <> "none" Then
                        'getdims = getdims & "[" & Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 8, (emark - bmark - FileNolen - 8)) & "]"
                        sRet = sRet & "[" & Mid(Accessories(CatNo, AccesNo), bmark + FileNolen + 8, (emark - bmark - FileNolen - 8)) & "]"
                    End If
                End If
            End If
        Next i

        getdims = sRet

        sTrace = "getdims: sRet = " & sRet : LogDebug(sTrace)
        sTrace = "getdims: Exit" : LogDebug(sTrace)

        Exit Function

Err_getdims:

        On Error Resume Next
        DB.Close()
        On Error GoTo 0
        Call LogUnstructuredError("Error at " & sTrace & " " & AccesNo & " " & CatNo & " " & Fileno, LOG_FATAL_ERROR)

    End Function 'getdims



    Sub ConfigureDwg()

        'this sub is used to open, modify and save the drawings

        Dim h, i, j As Integer
        Dim AcadApp As AcadApplication
        Dim dwgfile As AcadDocument                  'the curently modified document
        Dim TitleBlock As AcadBlockReference         'the block refrence for the title block (if applicable)
        'Dim varlayer As Object                      'used in iterating through layers
        Dim varAttributes 'As Variant                'used to iterate through attributes
        Dim AcadAttribute 'As AcadAttribute
        'Dim varobjects As Object                    'used to iterate through drawing objects
        Dim strDwgFileName As String
        Dim TagLen As Integer
        'Dim penfile As String
        Dim dwgFilename As String
        Dim DeleteSet As AcadSelectionSet
        Dim Layer As AcadLayer
        Dim mode As Integer
        Dim corner1(0 To 2) As Double
        Dim corner2(0 To 2) As Double
        Dim LayerGroup(0) As Int16
        Dim LayerValue(0) As Object
        Dim oLayerGroup As Object
        Dim oLayerValue As Object
        'Dim SetCount As Integer
        'Dim bRet As Boolean
        Dim TitleSave As String
        Dim CopySave As String
        Dim AcadGrp As AcadGroup                     '3.0.2.0
        Dim IdxGrp As Integer                        '3.0.2.0
        Dim C2WCommand As String           'Added 2/4/2021 -EJJ C2W intgration

        sTrace = "ConfigureDwg: Enter" : LogDebug(sTrace)

        Dim x, y As Integer
        If bDebug Then
            For x = 0 To UBound(DwgData, 1)
                For y = 0 To UBound(DwgData, 2)
                    sTrace = "ConfigureDwg: " & "DwgData(" & x & "," & y & ") = " & DwgData(x, y) : LogDebug(sTrace)
                Next
            Next
        End If

        dwgfile = Nothing
        AcadApp = Nothing

        TitleSave = OEMBlock
        CopySave = OEMCopy
        mode = AcSelect.acSelectionSetCrossing
        corner1(0) = 28 : corner1(1) = 17 : corner1(2) = 0
        corner2(0) = -3.3 : corner2(1) = -3.6 : corner2(2) = 0
        'penfile = sPgmsPath & "\monochrome.ctb"
        LayerGroup(0) = 8

        If AryConfigFile(1) = "" Then AryConfigFile(1) = EmailAddress

        If FileType = "pdf" Then
            Try
                Kill(sOutPath & "*.pdf")
            Catch ex As Exception
                'no files to delete
            End Try
        End If

        sTrace = "ConfigureDwg: Create Acad App" : LogDebug(sTrace)
        Try
            AcadApp = New AcadApplication
        Catch ex As System.Runtime.InteropServices.COMException
            GoTo Err_ConfigureDwg
        End Try

        sTrace = "ConfigureDwg: Set Acad App Properties" : LogDebug(sTrace)
        AcadApp.Visible = True
        AcadApp.ActiveDocument.Close(False) 'close default Drawing1.dwg

        Thread.Sleep(10000)

        'iterate through all the dwgdata entries
        For i = 0 To DwgCount - 1
            'checks for a file name in the drawing data array entry
            If DwgData(0, i) <> "" Then
                'opens the specified drawing in the drawing data array
                sTrace = "ConfigureDwg: dwgfile = AcadApp.Documents.Open(" & sDataPath & DwgData(0, i) & ")" : LogDebug(sTrace)
                Try
                    dwgfile = AcadApp.Documents.Open(sDataPath & DwgData(0, i))
                Catch ex As System.Runtime.InteropServices.COMException
                    GoTo Err_ConfigureDwg
                End Try

                'While Not AcadApp.GetAcadState.IsQuiescent
                'End While
                Thread.Sleep(10000)

                sTrace = "ConfigureDwg: dwgfile.ActiveLayout properties" : LogDebug(sTrace)
                dwgfile.ActiveLayout.ShowPlotStyles = False      'Specifies whether plot styles are to be used in the plot (True) or use the plot styles assigned to objects in the drawing (False).
                dwgfile.ActiveLayout.PlotType = AcPlotType.acExtents
                dwgfile.ActiveLayout.CenterPlot = True
                strDwgFileName = "DRAWING"
                'On Error GoTo 0
                '3.0.2.0 - INC019869 - 07/23/2014 - Begin
                'AutoCad returns the groups in index order - store them in an array
                sTrace = "ConfigureDwg: Loading " & dwgfile.Groups.Count.ToString() & " Group Index Translations" : LogDebug(sTrace)
                IdxGrp = 0
                ReDim AryGrp(IdxGrp)
                For Each AcadGrp In dwgfile.Groups()
                    ReDim Preserve AryGrp(IdxGrp)
                    AryGrp(IdxGrp) = AcadGrp.Name
                    sTrace = "ConfigureDwg: " & IdxGrp & " - " & AryGrp(IdxGrp) : LogDebug(sTrace)
                    IdxGrp += 1
                Next
                '3.0.2.0 - INC019869 - 07/23/2014 - End
                sTrace = "ConfigureDwg: Call SetDims" : LogDebug(sTrace)
                Call SetDims(i, dwgfile)
                Err.Clear() 'if there was an error it should have been cleared in LogError, but just in case...
                sTrace = "ConfigureDwg: Call SetLayers" : LogDebug(sTrace)
                Call SetLayers(i, dwgfile)
                Err.Clear()
                sTrace = "ConfigureDwg: Call Set Notes" : LogDebug(sTrace)
                Call SetNotes(i, dwgfile)
                Err.Clear()
                sTrace = "ConfigureDwg: Call InsertTitleBlock" : LogDebug(sTrace)

                C2WCommand = "C2W_SAVEAS2" & vbCr & SharePointSite & strDwgFileName & ".dwg" & vbCr & "1" & vbCr & "AutoSub Request" & vbCr ' Added 2/4/2021 -EJJ C2W Int.

                If OEMBlock <> "" Then InsertTitleBlock(sDataPath & OEMBlock, sDataPath & OEMCopy, dwgfile, C2WCommand) 'Added C2W argument 2/4/2021 -Ejj
                OEMBlock = TitleSave
                OEMCopy = CopySave
                'iterates through all the block refrences
                sTrace = "ConfigureDwg: Iterates through all the block refrences" : LogDebug(sTrace)
                For h = 0 To dwgfile.ModelSpace.Count - 1
                    If dwgfile.ModelSpace.Item(h).ObjectName = "AcDbBlockReference" Then
                        If LCase(dwgfile.ModelSpace.Item(h).Name) = "bacstblk" Then                '3.0.2.0 - INC020053 - 07/23/2014
                            sTrace = "ConfigureDwg: TitleBlock Found" : LogDebug(sTrace)
                            TitleBlock = dwgfile.ModelSpace.Item(h)
                            varAttributes = TitleBlock.GetAttributes
                            j = 0
                            For Each AcadAttribute In varAttributes
                                If varAttributes(j).TagString = "ORDER" Then varAttributes(j).TextString = AryConfigFile(1)
                                'added for config version 11/10/2020  -ejj
                                If varAttributes(j).TagString = "DATE" Then
                                    If ConfigVersion = "Nothing" Then
                                        varAttributes(j).TextString = Now()  
                                    Else
                                       varAttributes(j).TextString = ConfigVersion
                                    End If
                                End if
                                If varAttributes(j).TagString = "DESC" Then varAttributes(j).TextString = strCustData
                                If varAttributes(j).TagString = "DWGNUM" Then
                                    Select Case i
                                        Case 0
                                            strDwgFileName = ("UP-" & AryConfigFile(1) & strFileSuffix)
                                        Case 1
                                            strDwgFileName = ("SS-" & AryConfigFile(1) & strFileSuffix)
                                        Case Else
                                            TagLen = Len(DwgData(5, i))
                                            strDwgFileName = (Mid(DwgData(5, i), 2, TagLen - 2) & "-" & AryConfigFile(1) & strFileSuffix)
                                    End Select
                                    varAttributes(j).TextString = strDwgFileName
                                End If
                                j = j + 1
                            Next
                        End If
                        If LCase(dwgfile.ModelSpace.Item(h).Name) = "fileno" Then                  '3.0.2.0 - INC020053 - 07/23/2014
                            sTrace = "ConfigureDwg: Set title block" : LogDebug(sTrace)
                            TitleBlock = dwgfile.ModelSpace.Item(h)
                            varAttributes = TitleBlock.GetAttributes
                            j = 0
                            For Each AcadAttribute In varAttributes
                                If varAttributes(j).TagString = "FILENO" Then varAttributes(j).TextString = strDwgFileName
                                j = j + 1
                            Next
                        End If
                        If LCase(dwgfile.ModelSpace.Item(h).Name) = "model weights" Then           '3.0.2.0 - INC020053 - 07/23/2014
                            sTrace = "ConfigureDwg: Set model wieghts block" : LogDebug(sTrace)
                            TitleBlock = dwgfile.ModelSpace.Item(h)
                            varAttributes = TitleBlock.GetAttributes
                            j = 0
                            For Each AcadAttribute In varAttributes
                                If varAttributes(j).TagString = "MODELNUM" Then varAttributes(j).TextString = ModelCallout
                                If varAttributes(j).TagString = "SHIPWEIGHT" Then varAttributes(j).TextString = lngWieght(0)
                                If varAttributes(j).TagString = "OPWEIGHT" Then varAttributes(j).TextString = lngWieght(1)
                                If varAttributes(j).TagString = "HVYWEIGHT" Then varAttributes(j).TextString = lngWieght(2)
                                j = j + 1
                            Next
                        End If
                        If LCase(dwgfile.ModelSpace.Item(h).Name) = "anchorage weights" Then       '3.0.2.0 - INC020053 - 07/23/2014
                            sTrace = "ConfigureDwg: Set anchorage wieghts block" : LogDebug(sTrace)
                            TitleBlock = dwgfile.ModelSpace.Item(h)
                            varAttributes = TitleBlock.GetAttributes
                            j = 0
                            For Each AcadAttribute In varAttributes
                                If varAttributes(j).TagString = "MODELNUM" Then varAttributes(j).TextString = ModelCallout
                                If varAttributes(j).TagString = "HVYWEIGHT" Then varAttributes(j).TextString = lngWieght(2)
                                If varAttributes(j).TagString = "SHIPWEIGHT" Then varAttributes(j).TextString = lngWieght(0)
                                If varAttributes(j).TagString = "OPWEIGHT" Then varAttributes(j).TextString = lngWieght(1)
                                If varAttributes(j).TagString = "POINT1" Then varAttributes(j).TextString = lngWieght(3)
                                If varAttributes(j).TagString = "POINT2" Then varAttributes(j).TextString = lngWieght(4)
                                If varAttributes(j).TagString = "POINT3" Then varAttributes(j).TextString = lngWieght(5)
                                If varAttributes(j).TagString = "POINT4" Then varAttributes(j).TextString = lngWieght(6)
                                If varAttributes(j).TagString = "POINT5" Then varAttributes(j).TextString = lngWieght(7)
                                If varAttributes(j).TagString = "POINT6" Then varAttributes(j).TextString = lngWieght(8)
                                If varAttributes(j).TagString = "POINT7" Then varAttributes(j).TextString = lngWieght(9)
                                If varAttributes(j).TagString = "POINT8" Then varAttributes(j).TextString = lngWieght(10)
                                If varAttributes(j).TagString = "POINT9" Then varAttributes(j).TextString = lngWieght(11)
                                If varAttributes(j).TagString = "POINT10" Then varAttributes(j).TextString = lngWieght(12)
                                j = j + 1
                            Next
                        End If
                    End If
                Next h
                sTrace = "ConfigureDwg: Delete Frozen Layers" : LogDebug(sTrace)
                For Each Layer In dwgfile.Layers
                    If Layer.Freeze = True Then
                        If Layer.Lock = True Then
                            Layer.Lock = Not (Layer.Lock)
                        End If
                        LayerValue(0) = Layer.Name
                        'On Error GoTo 0
                        'DeleteSet = dwgfile.SelectionSets.Add("SSET")
                        'Call DeleteSet.Select(AcSelect.acSelectionSetAll, corner1, corner2, LayerGroup, LayerValue)
                        'DeleteSet.Erase()
                        'DeleteSet.Delete()
                        oLayerGroup = LayerGroup
                        LayerValue(0) = Layer.Name
                        oLayerValue = LayerValue
                        DeleteSet = dwgfile.SelectionSets.Add("SSET")
                        DeleteSet.Select(AcSelect.acSelectionSetAll, corner1, corner2, oLayerGroup, oLayerValue)
                        DeleteSet.Erase()
                        DeleteSet.Delete()
                    End If
                Next Layer
                dwgfile.PurgeAll()
                'wait for plot to finish before continuing
                dwgfile.SetVariable("BACKGROUNDPLOT", 0)
                sTrace = "ConfigureDwg: Save Drawing File " & sOutPath & strDwgFileName & "." & FileType : LogDebug(sTrace)
                Select Case FileType
                    Case "dwg"
                        dwgfile.SaveAs(sOutPath & strDwgFileName & ".dwg", AcSaveAsType.ac2000_dwg)
                    Case "dxf"
                        dwgfile.SaveAs(sOutPath & strDwgFileName & ".dxf", AcSaveAsType.ac2000_dxf)
                    Case "dwf"
                        dwgfile.Plot.PlotToFile(sOutPath & strDwgFileName & ".dwf", sPgmsPath & "\DWF6 ePlot.pc3")
                    Case "pdf"
                        dwgfile.Plot.PlotToFile(sOutPath & strDwgFileName & ".pdf", sPgmsPath & "\DWG To PDF.pc3")
                    Case Else
                        dwgfile.SaveAs(sOutPath & strDwgFileName & ".dwg")
                End Select
                LogDrawing(strDwgFileName & "." & FileType)
                sTrace = "ConfigureDwg: Close Drawing file" : LogDebug(sTrace)

                'CAD2WIN addition starts here: EJJ 1/25/2021
                
                C2WCommand = C2WCommand & ";Config Version|" & ConfigVersion & ";BAC Job Number|" _
                            & mid(AryConfigFile(1),1,8) & ";BAC Line Number|" & mid(AryConfigFile(1),9,2) & vbCr

                sTrace = "ConfigureDwg: C2W Command is:" &  C2WCommand : LogDebug(sTrace)
                
                dwgfile.sendcommand (C2WCommand)
                'CAD2WIN end

                dwgfile.Close(False)
            End If
        Next i

        dwgfile = Nothing

        For i = 0 To DwgCount - 1
            For j = 0 To 5
                DwgData(j, i) = ""
            Next j
        Next i

        sTrace = "ConfigureDwg: Quit Acad" : LogDebug(sTrace)
        AcadApp.Quit()
        AcadApp = Nothing

        sTrace = "ConfigureDwg: Exit" : LogDebug(sTrace)

        Exit Sub

Err_ConfigureDwg:

        'We don't want to leave any zombies lying around.
        dwgFilename = sDataPath & DwgData(0, i)

        If AcadApp Is Nothing Then
        Else
            If AcadApp.Documents.Count > 0 Then
                dwgfile.Close(False)
            End If

            If dwgfile Is Nothing Then
            Else
                dwgfile = Nothing
            End If

            AcadApp.Quit()
            AcadApp = Nothing
        End If

        Call LogUnstructuredError("Error at " & sTrace & " -- " & dwgfile.FullName, LOG_FATAL_ERROR)

    End Sub 'ConfigureDwg



    Sub SetDims(Index As Object, ByRef Dwg As AcadDocument)

        'parses the extraced dim data, combines the data then applies the dims to the drawing

        Dim LineLen As Integer
        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim BoolEmark As Boolean
        Dim EQMark As Integer
        Dim Cmark As Integer
        Dim DimGroup As String
        Dim DimGroupIdx As Integer                  '3.0.2.0
        Dim DimNumber As Object
        Dim DimText As String
        Dim DimNumLen As Integer

        sTrace = "SetDims: Enter" : LogDebug(sTrace)

        LineLen = Len(DwgData(2, Index))
        BoolEmark = False

        On Error GoTo Err_SetDims
        sTrace = "SetDims: Parse DwgData(2," & Index & ") " & DwgData(2, Index) : LogDebug(sTrace)
        For i = 1 To LineLen
            If Mid(DwgData(2, Index), i, 1) = "[" Then bmark = i
            If Mid(DwgData(2, Index), i, 1) = "," And BoolEmark = False Then Cmark = i
            If Mid(DwgData(2, Index), i, 1) = "=" And BoolEmark = False Then
                EQMark = i
                BoolEmark = True
            End If
            If Mid(DwgData(2, Index), i, 1) = "]" Then
                emark = i
                If LCase(Mid(DwgData(2, Index), bmark + 1, 10)) = "*flow*unit" Then
                    DimGroup = Mid(DwgData(2, Index), bmark + 12, Cmark - bmark - 12)
                Else
                    If LCase(Mid(DwgData(2, Index), bmark + 1, 6)) = "*flow*" Then
                        DimGroup = Mid(DwgData(2, Index), bmark + 7, Cmark - bmark - 7)
                    Else
                        If LCase(Mid(DwgData(2, Index), bmark + 1, 4)) = "unit" Then
                            DimGroup = Mid(DwgData(2, Index), bmark + 6, Cmark - bmark - 6)
                        Else
                            DimGroup = Mid(DwgData(2, Index), bmark + 1, Cmark - bmark - 1)
                        End If
                    End If
                End If
                DimNumber = Mid(DwgData(2, Index), Cmark + 1, EQMark - Cmark - 1)
                DimNumLen = Len(DimNumber)
                If Left(LCase(DimNumber), 1) = "x" Then DimNumber = 1
                If Left(LCase(DimNumber), 1) = "y" Then DimNumber = 2
                If Left(LCase(DimNumber), 1) = "s" Then DimNumber = 3
                DimText = Mid(DwgData(2, Index), EQMark + 1, emark - EQMark - 1)
                DimNumber = DimNumber - 1
                '3.0.2.0 - INC019869 - 07/23/2014 - Begin
                'On Error Resume Next
                'sTrace = "SetDims: Found DimGroup = " & DimGroup & "  DimNumber = " & DimNumber & " DimText = " & DimText : LogDebug(sTrace)
                'Dwg.Groups(DimGroup).Item(DimNumber).TextOverride = DimText
                'Dwg.Groups(DimGroup).Item(DimNumber).TextString = DimText
                DimGroupIdx = AryGrp.ToList().IndexOf(DimGroup)
                sTrace = "SetDims: Using DimGroup(" & DimGroupIdx & ") = " & DimGroup & "  DimNumber = " & DimNumber & "  DimText = " & DimText : LogDebug(sTrace)
                If Dwg.Groups(DimGroupIdx).Item(DimNumber).ObjectName = "AcDbRotatedDimension" Then
                    Dwg.Groups(DimGroupIdx).Item(DimNumber).TextOverride = DimText
                Else
                    Dwg.Groups(DimGroupIdx).Item(DimNumber).TextString = DimText
                End If
                'On Error GoTo Err_SetDims
                '3.0.2.0 - INC019869 - 07/23/2014 - End
                BoolEmark = False
            End If
        Next i

        sTrace = "SetDims: Exit" : LogDebug(sTrace)

        Exit Sub

Err_SetDims:

        Call LogUnstructuredError("Error at " & sTrace & " -- " & Dwg.FullName, LOG_ERROR)

    End Sub 'SetDims



    Sub SetLayers(Index As Integer, ByRef Dwg As AcadDocument)

        'this sub sets the layers on the "Dwg" drawing to Thaw after parsing the data extracted from the database

        Dim LineLen As Integer
        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim Layername As String

        sTrace = "SetLayers: Enter" : LogDebug(sTrace)

        'On Error GoTo Err_SetLayers

        LineLen = Len(DwgData(1, Index))
        Try
            Dwg.ActiveLayer = Dwg.Layers("0")
        Catch ex As System.Runtime.InteropServices.COMException
            GoTo Err_SetLayers
        End Try
        sTrace = "SetLayers: Layers On Parse DwgData(1," & Index & ") " & DwgData(1, Index) : LogDebug(sTrace)
        For i = 1 To LineLen
            If Mid(DwgData(1, Index), i, 1) = "[" Then bmark = i
            If Mid(DwgData(1, Index), i, 1) = "]" Then
                emark = i
                Layername = Mid(DwgData(1, Index), bmark + 1, emark - bmark - 1)
                If Left(Layername, 5) = "unit=" Then Layername = Mid(DwgData(1, Index), bmark + 6, emark - bmark - 6)
                If Left(Layername, 6) = "steel=" Then Layername = Mid(DwgData(1, Index), bmark + 7, emark - bmark - 7)
                sTrace = "SetLayers: Found Layername = " & Layername : LogDebug(sTrace)
                Try
                    'Dwg.Layers(Layername).LayerOn = True
                    Dwg.Layers.Item(Layername).LayerOn = True
                Catch ex As System.Runtime.InteropServices.COMException
                    GoTo Err_SetLayers
                End Try
                Try
                    'Dwg.Layers(Layername).Freeze = False
                    Dwg.Layers.Item(Layername).Freeze = False
                Catch ex As System.Runtime.InteropServices.COMException
                    GoTo Err_SetLayers
                End Try
                Try
                    Dwg.Regen(AcRegenType.acActiveViewport)
                Catch ex As System.Runtime.InteropServices.COMException
                    GoTo Err_SetLayers
                End Try
            End If
        Next i

        LineLen = Len(DwgData(4, Index))
        sTrace = "SetLayers: Layers Off Parse DwgData(4," & Index & ") " & DwgData(4, Index) : LogDebug(sTrace)
        For i = 1 To LineLen
            If Mid(DwgData(4, Index), i, 1) = "[" Then bmark = i
            If Mid(DwgData(4, Index), i, 1) = "]" Then
                emark = i
                Layername = Mid(DwgData(4, Index), bmark + 1, emark - bmark - 1)
                If Left(Layername, 5) = "unit=" Then Layername = Mid(DwgData(4, Index), bmark + 6, emark - bmark - 6)
                If Left(Layername, 6) = "steel=" Then Layername = Mid(DwgData(4, Index), bmark + 7, emark - bmark - 7)
                sTrace = "SetLayers: Found Layername = " & Layername : LogDebug(sTrace)
                Try
                    'Dwg.Layers(Layername).LayerOn = False
                    'Dwg.Layers(Layername).Freeze = True
                    Dwg.Layers.Item(Layername).LayerOn = False
                    Dwg.Layers.Item(Layername).Freeze = True
                    Dwg.Regen(AcRegenType.acActiveViewport)
                Catch ex As System.Runtime.InteropServices.COMException
                    GoTo Err_SetLayers
                End Try
            End If
        Next i

        sTrace = "SetLayers: Exit" : LogDebug(sTrace)

        Exit Sub

Err_SetLayers:

        Call LogUnstructuredError("Error at " & sTrace & " -- " & Dwg.FullName, LOG_ERROR)

    End Sub 'SetLayers



    Sub SetNotes(Index As Object, ByRef Dwg As AcadDocument)

        'parses the extraced notes data, combines the data then applies the notes to the drawing

        Dim LineLen As Integer
        Dim h, i As Integer
        Dim bmark As Integer
        Dim emark As Integer
        Dim NoteText As String
        Dim NoteCount As Integer
        Dim STnote As String
        Dim lngVersion As Double
        Dim lngTempver As Double
        Dim strVer As String
        Dim STNoteArray() As String
        Dim stnoteCnt As Integer
        Dim DimGroupIdx As Integer                  '3.0.2.0

        sTrace = "SetNotes: Enter" : LogDebug(sTrace)

        LineLen = Len(DwgData(3, Index))

        'On Error GoTo Err_SetNotes                 '3.0.2.0

        NoteCount = 1
        stnoteCnt = 1
        NoteText = ""
        ReDim STNoteArray(0)

        Try
            sTrace = "SetNotes: Parse DwgData(3," & Index & ") " & DwgData(3, Index) : LogDebug(sTrace)
            For i = 1 To LineLen
                If Mid(DwgData(3, Index), i, 1) = "[" Then bmark = i
                If Mid(DwgData(3, Index), i, 1) = "]" Then
                    emark = i
                    If Left(Mid(DwgData(3, Index), bmark + 1, emark - bmark - 1), 1) = "#" Then
                        STnote = ""
                        For h = 0 To stnoteCnt - 1
                            If Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2) = STNoteArray(h) Then GoTo FoundNote
                        Next h
                        stnoteCnt = stnoteCnt + 1
                        ReDim Preserve STNoteArray(stnoteCnt)
                        STNoteArray(stnoteCnt - 1) = Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2)
                        sTrace = "SetNotes: Found Note " & STNoteArray(stnoteCnt - 1) : LogDebug(sTrace)
                        Call StandardNotes(Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2), STnote, NoteCount)
                        NoteText = NoteText + STnote
FoundNote:
                    Else
                        If Left(Mid(DwgData(3, Index), bmark + 1, emark - bmark - 1), 5) = "unit=" Then
                            If Left(Mid(DwgData(3, Index), bmark + 6, emark - bmark - 6), 1) = "#" Then
                                STnote = ""
                                For h = 0 To stnoteCnt - 1
                                    If Mid(DwgData(3, Index), bmark + 6, emark - bmark - 7) = STNoteArray(h) Then GoTo FoundNote
                                Next h
                                stnoteCnt = stnoteCnt + 1
                                ReDim Preserve STNoteArray(stnoteCnt)
                                STNoteArray(stnoteCnt - 1) = Mid(DwgData(3, Index), bmark + 6, emark - bmark - 6)
                                sTrace = "SetNotes: Found Note " & STNoteArray(stnoteCnt - 1) : LogDebug(sTrace)
                                Call StandardNotes(Mid(DwgData(3, Index), bmark + 7, emark - bmark - 7), STnote, NoteCount)
                                NoteText = NoteText + STnote
                            Else
                                NoteText = NoteText & NoteCount & ") " & Mid(DwgData(3, Index), bmark + 6, emark - bmark - 6) & Chr(13) & Chr(10)
                                NoteCount = NoteCount + 1
                            End If
                        Else
                            If Left(Mid(DwgData(3, Index), bmark + 1, emark - bmark - 1), 1) = "#" Then
                                STnote = ""
                                For h = 0 To stnoteCnt - 1
                                    If Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2) = STNoteArray(h) Then GoTo FoundNote
                                Next h
                                stnoteCnt = stnoteCnt + 1
                                ReDim Preserve STNoteArray(stnoteCnt)
                                STNoteArray(stnoteCnt - 1) = Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2)
                                sTrace = "SetNotes: Found Note " & STNoteArray(stnoteCnt - 1) : LogDebug(sTrace)
                                Call StandardNotes(Mid(DwgData(3, Index), bmark + 2, emark - bmark - 2), STnote, NoteCount)
                                NoteText = NoteText + STnote
                            Else
                                NoteText = NoteText & NoteCount & ") " & Mid(DwgData(3, Index), bmark + 1, emark - bmark - 1) & Chr(13) & Chr(10)
                                NoteCount = NoteCount + 1
                            End If
                        End If
                    End If
                End If
            Next i

            sTrace = "SetNotes: Parse Version Table" : LogDebug(sTrace)
            lngVersion = 0
            tblVersion.MoveFirst()
            Do While tblVersion.EOF = False
                lngTempver = tblVersion.Fields("Version").Value
                If lngTempver > lngVersion Then
                    lngVersion = tblVersion.Fields("Version").Value
                End If
                tblVersion.MoveNext()
            Loop

            strVer = Format(lngVersion, "#0.00")
            NoteText = "Notes" & Chr(13) & Chr(10) & NoteText
        Catch e As Exception
            'GoTo Err_SetNotes
            Call LogStructuredError("Error at " & sTrace & " -- " & Dwg.FullName, LOG_ERROR, e)
        End Try

        '3.0.2.0 - INC019869 - 07/23/2014 - Begin
        'On Error GoTo 0

        'On Error Resume Next
        'sTrace = "SetNotes: Groups(DATAVERSION)" : LogDebug(sTrace)
        'Dwg.Groups("DATAVERSION").Item(0).TextString = "Data Version " & strVer
        'On Error GoTo 0

        'On Error Resume Next
        'sTrace = "SetNotes: Groups(NOTES)" : LogDebug(sTrace)
        'Dwg.Groups("NOTES").Item(0).TextString = NoteText

        Try
            DimGroupIdx = AryGrp.ToList().IndexOf("DATAVERSION")
            sTrace = "SetNotes: DATAVERSION Groups(" & DimGroupIdx & ").Item(0)" : LogDebug(sTrace)
            Dwg.Groups(DimGroupIdx).Item(0).TextString = "Data Version " & strVer
        Catch e As Exception
            Call LogStructuredError("Warning at " & sTrace & " -- " & Dwg.FullName, LOG_WARNING, e)
            'GoTo Err_SetNotes
        End Try

        Try
            DimGroupIdx = AryGrp.ToList().IndexOf("NOTES")
            sTrace = "SetNotes: NOTES Groups(" & DimGroupIdx & ").Item(0)" : LogDebug(sTrace)
            Dwg.Groups(DimGroupIdx).Item(0).TextString = NoteText
        Catch e As Exception
            'GoTo Err_SetNotes
            Call LogStructuredError("Warning at " & sTrace & " -- " & Dwg.FullName, LOG_WARNING, e)
        End Try

        'On Error GoTo 0
        '3.0.2.0 - INC019869 - 07/23/2014 - End

        sTrace = "SetNotes: Exit" : LogDebug(sTrace)

        Exit Sub

Err_SetNotes:

        Call LogUnstructuredError("Error at " & sTrace & " -- " & Dwg.FullName, LOG_ERROR)

    End Sub 'SetNotes


    'Sub StandardNotes(NumberIn As String, NoteOutput As String, NoteCount As Integer)  3.0.2.0 - INC019869 - 07/23/2014
    Sub StandardNotes(ByRef NumberIn As String, ByRef NoteOutput As String, ByRef NoteCount As Integer)

        Dim NoteOut As String
        Dim i As Integer
        Dim bmark As Integer
        Dim emark As Integer

        sTrace = "StandardNotes: Enter" : LogDebug(sTrace)

        On Error GoTo Err_StandardNotes

        'sets notes from standard notes database
        sTrace = "StandardNotes: Open Notes database" : LogDebug(sTrace)
        DBN.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & sDataPath & "\All Product\Product Database\Standard Notes.mdb" & ";Persist Security Info=False"
        DBN.Open()
        sTrace = "StandardNotes: Open Notes table" : LogDebug(sTrace)
        NotesRecSet.Open("Notes", DBN, CursorTypeEnum.adOpenDynamic)

        sTrace = "StandardNotes: Locate note number " & NumberIn : LogDebug(sTrace)
        NoteOut = ""
        NotesRecSet.MoveFirst()
        Do While NotesRecSet.EOF = False
            If NotesRecSet.Fields.Item("NoteNum").Value = NumberIn Then
                NoteOut = NotesRecSet.Fields.Item("NoteText").Value
                sTrace = "StandardNotes: Found " & NoteOut : LogDebug(sTrace)
                Exit Do
            End If
            NotesRecSet.MoveNext()
        Loop
        Dim linelenS As Integer
        linelenS = Len(NoteOut)

        For i = 1 To linelenS
            If Mid(NoteOut, i, 1) = "[" Then bmark = i
            If Mid(NoteOut, i, 1) = "]" Then
                emark = i
                NoteOutput = NoteOutput & NoteCount & ") " & Mid(NoteOut, bmark + 1, emark - bmark - 1) & Chr(13) & Chr(10)
                NoteCount = NoteCount + 1
            End If
        Next i

        sTrace = "StandardNotes: Close Notes database" : LogDebug(sTrace)
        DBN.Close()

        sTrace = "StandardNotes: Exit" : LogDebug(sTrace)

        Exit Sub

Err_StandardNotes:

        On Error Resume Next
        DBN.Close()
        On Error GoTo 0
        Call LogUnstructuredError("Error at " & sTrace, LOG_ERROR)

    End Sub 'StandardNotes



    Sub InsertTitleBlock(OEMborder As String, OEMCopyright As String, ByRef Dwg As AcadDocument, ByRef C2WCommand As String)  'Added C2W argument 2/4/2021 -Ejj

        Dim i, j As Integer
        Dim Border As AcadBlockReference
        Dim BlockOrigin As Object
        Dim BlockScale As Double
        Dim TitleBlock As AcadBlockReference
        Dim CopyBlock As AcadBlockReference
        Dim OldTitleBlock As AcadBlockReference
        Dim StrModel1 As String
        Dim StrModel2 As String
        Dim AcadAttribute 'As AcadAttribute
        Dim varAttributes 'As Variant
        Dim OldCopyBlock As AcadBlockReference
        Dim CopyOrigin(2) As Double
        Dim CopyScale As Double

        sTrace = "InsertTitleBlock: Enter" : LogDebug(sTrace)

        'OEMborder = "L:\APPS\Auto Submittal Program\Product Data\All Product\Title_Blocks\Frick\bacstblk.dwg"
        On Error GoTo Err_InsertTitleBlock

        StrModel1 = ""
        StrModel2 = ""
        Border = Nothing
        OldTitleBlock = Nothing
        OldCopyBlock = Nothing

        sTrace = "InsertTitleBlock: Locate Blocks" : LogDebug(sTrace)
        For i = 0 To Dwg.ModelSpace.Count - 1
            If Dwg.ModelSpace.Item(i).ObjectName = "AcDbBlockReference" Then
                If LCase(Dwg.ModelSpace.Item(i).Name) = "border" Then
                    Border = Dwg.ModelSpace.Item(i)
                    sTrace = "InsertTitleBlock: Found Border" : LogDebug(sTrace)
                End If
                If LCase(Dwg.ModelSpace.Item(i).Name) = "bacstblk" Then
                    OldTitleBlock = Dwg.ModelSpace.Item(i)
                    sTrace = "InsertTitleBlock: Found OldTitleBlock" : LogDebug(sTrace)
                End If
                If LCase(Dwg.ModelSpace.Item(i).Name) = "copyr0" Then
                    OldCopyBlock = Dwg.ModelSpace.Item(i)
                    sTrace = "InsertTitleBlock: Found OldCopyBlock" : LogDebug(sTrace)
                End If
            End If
        Next i

        sTrace = "InsertTitleBlock: Locate OldTitleBlock Attributes" : LogDebug(sTrace)
        If OldTitleBlock Is Nothing Then
        Else
            varAttributes = OldTitleBlock.GetAttributes
            j = 0
            For Each AcadAttribute In varAttributes
                If varAttributes(j).TagString = "MODEL1" Then StrModel1 = varAttributes(j).TextString
                If varAttributes(j).TagString = "MODEL2" Then StrModel2 = varAttributes(j).TextString
                j = j + 1
            Next
        End If

        C2WCommand = C2WCommand &  "Description 1|" & StrModel1 & ";Description 2|" & StrModel2 'added 2/4/2021 for C2W  -EJJ  

        CopyOrigin(0) = OldCopyBlock.InsertionPoint(0)
        CopyOrigin(1) = OldCopyBlock.InsertionPoint(1)
        CopyOrigin(2) = OldCopyBlock.InsertionPoint(2)
        CopyScale = OldCopyBlock.XScaleFactor

        BlockOrigin = Border.InsertionPoint
        BlockScale = Border.XScaleFactor / 10

        sTrace = "InsertTitleBlock: Delete old blocks" : LogDebug(sTrace)
        '3.0.2.0 - INC020053 - 07/23/2014 - Begin
        'OldTitleBlock.Delete()
        'OldCopyBlock.Delete()
        If OldTitleBlock Is Nothing Then
            sTrace = "InsertTitleBlock: OldTitleBlock Not Found" : LogDebug(sTrace)
        Else
            OldTitleBlock.Delete()
        End If
        If OldCopyBlock Is Nothing Then
            sTrace = "InsertTitleBlock: OldCopyBlock Not Found" : LogDebug(sTrace)
        Else
            OldCopyBlock.Delete()
        End If
        '3.0.2.0 - INC020053 - 07/23/2014 - End

        Dwg.PurgeAll()
        Dwg.Regen(AcRegenType.acActiveViewport)

        sTrace = "InsertTitleBlock: Insert TitleBlock" : LogDebug(sTrace)
        TitleBlock = Dwg.ModelSpace.InsertBlock(BlockOrigin, OEMborder, BlockScale, BlockScale, BlockScale, 0, "")
        sTrace = "InsertTitleBlock: Insert CopyBlock" : LogDebug(sTrace)
        CopyBlock = Dwg.ModelSpace.InsertBlock(CopyOrigin, OEMCopyright, CopyScale, CopyScale, CopyScale, 0, "")

        sTrace = "InsertTitleBlock: Locate TitleBlock Attributes" : LogDebug(sTrace)
        varAttributes = TitleBlock.GetAttributes
        j = 0
        For Each AcadAttribute In varAttributes
            If varAttributes(j).TagString = "MODEL1" Then varAttributes(j).TextString = StrModel1
            If varAttributes(j).TagString = "MODEL2" Then varAttributes(j).TextString = StrModel2
            j = j + 1
        Next
        Dwg.Regen(AcRegenType.acAllViewports)

        OEMborder = ""
        OEMCopyright = ""

        On Error GoTo 0

        sTrace = "InsertTitleBlock: Exit" : LogDebug(sTrace)

        Exit Sub

Err_InsertTitleBlock:

        Call LogUnstructuredError("Error at " & sTrace & " -- " & Dwg.FullName, LOG_ERROR)

    End Sub 'InsertTitleBlock



    Sub LogDrawing(sMsg As String)

        File.AppendAllText(sOutPath & sDrawingFileName, sMsg & Environment.NewLine)

    End Sub 'LogDrawing



    Sub LogDebug(sMsg As String)

        If bDebug Then
            File.AppendAllText(sOutPath & sDebugFileName, Now() & "  " & sMsg & Environment.NewLine)
        End If

    End Sub 'LogDebug



    Sub LogUnstructuredError(ErrorBody As String, ErrorType As String)

        File.AppendAllText(sOutPath & sErrorFileName, ErrorBody & Environment.NewLine)
        File.AppendAllText(sOutPath & sErrorFileName, ErrorType & " " & Err.Number & " " & Err.Description & Environment.NewLine)
        File.AppendAllText(sOutPath & sErrorFileName, "" & Environment.NewLine)

        If bDebug Then
            File.AppendAllText(sOutPath & sDebugFileName, Now() & "  " & UCase(ErrorType) & " " & Err.Number & " " & Err.Description & Environment.NewLine)
        End If

        Select Case ErrorType
            Case LOG_WARNING
                WarningCount += 1
                Err.Clear()
            Case LOG_ERROR
                ErrorCount += 1
                Err.Clear()
            Case LOG_FATAL_ERROR
                ErrorCount += 1
                Call EndProgram()
        End Select

    End Sub 'LogUnstructuredError



    '3.0.2.0 - INC019869 - 07/23/2014 - Begin
    Sub LogStructuredError(ErrorBody As String, ErrorType As String, e As Exception)

        File.AppendAllText(sOutPath & sErrorFileName, ErrorBody & Environment.NewLine)
        File.AppendAllText(sOutPath & sErrorFileName, ErrorType & " " & e.Message & Environment.NewLine)
        File.AppendAllText(sOutPath & sErrorFileName, "" & Environment.NewLine)

        If bDebug Then
            File.AppendAllText(sOutPath & sDebugFileName, Now() & "  " & UCase(ErrorType) & " " & e.Message & Environment.NewLine)
        End If

        Select Case ErrorType
            Case LOG_WARNING
                WarningCount += 1
                Err.Clear()
            Case LOG_ERROR
                ErrorCount += 1
                Err.Clear()
            Case LOG_FATAL_ERROR
                ErrorCount += 1
                Call EndProgram()
        End Select

    End Sub 'LogStructuredError
    '3.0.2.0 - INC019869 - 07/23/2014 - End



    Sub EndProgram()

        Dim iRet As Integer

        sTrace = "EndProgram: Enter" : LogDebug(sTrace)

        On Error Resume Next

        Acces = Nothing
        InConns = Nothing
        OutConns = Nothing
        Drive = Nothing
        Data = Nothing
        Anchor = Nothing
        FlowDB = Nothing
        NotesRecSet = Nothing
        tblVersion = Nothing
        TransTable = Nothing
        DB = Nothing
        DBN = Nothing

        If ErrorCount > 0 Then
            iRet = ErrorCount + 100
            Call ErrorContext()
        Else
            iRet = DwgCount
        End If

        sTrace = "Return code: " & iRet : LogDebug(sTrace)
        sTrace = "EndProgram: Exit" : LogDebug(sTrace)

        Environment.Exit(iRet)

    End Sub 'EndProgram



    Sub ErrorContext()

        Dim i As Integer

        sTrace = "ErrorContext: Enter" : LogDebug(sTrace)

        File.AppendAllText(sOutPath & sErrorFileName, "" & Environment.NewLine)

        'input file name
        File.AppendAllText(sOutPath & sErrorFileName, sInFile & Environment.NewLine)

        File.AppendAllText(sOutPath & sErrorFileName, "" & Environment.NewLine)

        'configuration array
        File.AppendAllText(sOutPath & sErrorFileName, "Configuration Array" & Environment.NewLine)
        For i = 0 To intLineCount - 1
            File.AppendAllText(sOutPath & sErrorFileName, i.ToString & " " & AryConfigFile(i) & Environment.NewLine)
        Next i

        File.AppendAllText(sOutPath & sErrorFileName, "" & Environment.NewLine)

        'drawing context
        If DwgData(0, 0) <> "" Then
            File.AppendAllText(sOutPath & sErrorFileName, "Drawing Context" & Environment.NewLine)
            For i = 0 To DwgCount - 1
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " =                " & DwgData(0, i) & Environment.NewLine)
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " Layers On =      " & DwgData(1, i) & Environment.NewLine)
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " Dimensions =     " & DwgData(2, i) & Environment.NewLine)
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " Notes =          " & DwgData(3, i) & Environment.NewLine)
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " Layers Off =     " & DwgData(4, i) & Environment.NewLine)
                File.AppendAllText(sOutPath & sErrorFileName, "Drawing " & i.ToString & " Drawing Prefix = " & DwgData(5, i) & Environment.NewLine)
            Next i
        End If

        sTrace = "ErrorContext: Exit" : LogDebug(sTrace)

    End Sub 'ErrorContext



End Module
