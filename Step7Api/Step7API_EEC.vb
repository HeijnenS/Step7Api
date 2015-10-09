Public Class Step7API_EEC




    'Option Explicit

    Dim fso As Scripting.FileSystemObject
    Dim m_Project As S7Project
    Dim a As String

    Private Function CreateHardware() As Boolean
        'On Error Resume Next
        On Error GoTo 0
        Dim Sta As S7Station
        Dim fullFilename As String

        fullFilename = TextBox1.Text & "\HardwareConfig\HardwareConfigMain.cfg"
        Sta = m_Project.Stations.Import(fullFilename)
        If Err.Number = 0 Then
            CreateHardware = True
        Else
            txtStatus.Text = "Error while adding the hardware: (" & Err.Number & ") " & Err.Description
            CreateHardware = False
        End If
    End Function

    Private Function ImportSources() As Boolean
        'met creeren van een nieuw program wordt automatisch een nieuwe symboltable aangemaakt.
        'On Error Resume Next
        On Error Resume Next
        Dim Program As Object

        Dim Source As IS7Source
        Dim SourceFolder As IS7Container3
        Dim Symbol As IS7SymbolTable
        Dim lResult As Long

        Dim name
        Dim Filename As String
        Dim Attribs As Microsoft.VisualBasic.FileAttribute
        Dim ListFiles
        Dim strExtension As String
        Dim myProgram As IS7Program
        Dim ProgramFolder As S7ProgramType


        ListFiles = New Collection
        myProgram = Nothing
        For Each myProgram In m_Project.Programs
            If myProgram.Type = S7ProgramType.S7 Then  ': 1327367 : S7ProgramType : Tabelle1.ImportSources
                'myProgram.name = "iDea"
                Exit For
            End If
        Next
        Program = myProgram
        Symbol = Program.SymbolTable
        lResult = Symbol.Import(TextBox1.Text & "\TagList\TaglistMain.sdf", S7SymImportFlags.S7SymImportOverwriteNameLeading) 'Import symbol list
        Filename = Dir(TextBox1.Text + "\SourceCode\")     ' first call to Dir() initializes the list

        While Filename <> ""
            Attribs = GetAttr(TextBox1.Text + "\SourceCode\" & Filename)      ' to be added, a file must have the right set of attributes
            If Not Attribs And vbSystem Or vbHidden Then
                strExtension = UCase(Split(Filename, ".")(UBound(Split(Filename, "."))))
                If strExtension = "AWL" Or strExtension = "SCL" Or strExtension = "INP" Then
                    ListFiles.Add(Filename, TextBox1.Text + "\SourceCode\" & Filename)
                End If
            End If
            ' fetch next filename
            Filename = Dir()
        End While

        'Check all S7 software items, that can be found in Program.Next
        Dim Container As IS7SWItem
        For Each Container In Program.Next
            If (Container.Type = S7SWObjType.S7Container) Then             'If S7Container is found -> check the concrete type of this container
                Dim SWContainer As IS7Container
                SWContainer = Container
                If SWContainer.ConcreteType = S7ContainerType.S7SourceContainer Then
                    For Each name In ListFiles
                        Source = SWContainer.Next.Add(Split(name, ".")(0), S7SWObjType.S7Source, TextBox1.Text + "\SourceCode\" + name)
                    Next name
                End If
            End If
        Next

        If Err.Number = 0 Then
            ImportSources = True
        Else
            ImportSources = False
        End If

    End Function


    Private Function CompileSCLSources() As Boolean
        On Error GoTo 0
        Dim test As Object
        Dim myProgram As IS7Program
        Dim Program As IS7Program
        Dim SourceFolder As IS7SWItem
        Dim MakeFile As S7Source
        Dim myFileName As String


        If m_Project Is Nothing Then
            'Exit Sub
            'btnOpenProject_Click
        End If

        myProgram = Nothing
        For Each myProgram In m_Project.Programs
            If myProgram.Type = S7ProgramType.S7 Then  ': 1327367 : S7ProgramType : Tabelle1.ImportSources
                myProgram.Name = "iDea"
                Exit For
            End If
        Next
        Program = myProgram

        'Check all S7 software items, that can be found in Program.Next (Sources & Blocks folder)
        Dim Container As IS7SWItem

        SourceFolder = Program.Next.Item("Sources")
        MakeFile = SourceFolder.Next.Item("Make")
        Dim s7SWItems As SimaticLib.S7SWItems
        s7SWItems = MakeFile.Compile()

        If Err.Number = 0 Then
            CompileSCLSources = True
        Else
            CompileSCLSources = False
        End If


    End Function

    Private Function CompileAWL() As Boolean
        Dim mySource As S7Source
        Dim SourceFolder As IS7SWItem
        Dim AWL0 As Collection
        Dim AWL1 As Collection
        Dim AWL2 As Collection
        Dim AWL3 As Collection
        Dim Program As IS7Program
        Dim myProgram As IS7Program
        Dim myFileName As String

        AWL0 = New Collection
        AWL1 = New Collection
        AWL2 = New Collection
        AWL3 = New Collection


        If m_Project Is Nothing Then
            'Exit Sub
            'btnOpenProject_Click
        End If

        myProgram = Nothing
        For Each myProgram In m_Project.Programs
            If myProgram.Type = S7ProgramType.S7 Then  ': 1327367 : S7ProgramType : Tabelle1.ImportSources
                myProgram.Name = "iDea"
                Exit For
            End If
        Next
        Program = myProgram

        'Check all S7 software items, that can be found in Program.Next (Sources & Blocks folder)
        'Dim Container As IS7SWItem
        SourceFolder = Program.Next.Item("Sources")
        For Each mySource In SourceFolder.Next
            'compile Lad elements
            If mySource.ConcreteType = S7SourceType.S7AWL Then
                'create a collection that is sorted by name length and compile order is longest first
                If InStr(1, mySource.Name, "@") > 0 Then
                    myFileName = Split(mySource.Name, "@")(1)
                    Select Case Len(myFileName)
                        Case Is > 6 : AWL3.Add(mySource)
                        Case Is > 3 : AWL2.Add(mySource)
                        Case Is > 0 : AWL1.Add(mySource)
                    End Select
                Else
                    If mySource.Name = "Main" Then
                        AWL0.Add(mySource)
                    Else
                        MsgBox(mySource.Name & "wordt niet gecompileerd (mist @)")
                    End If
                End If
            End If
        Next

        For Each mySource In AWL3
            mySource.Compile()
        Next

        For Each mySource In AWL2
            mySource.Compile()
        Next

        For Each mySource In AWL1
            mySource.Compile()
        Next

        For Each mySource In AWL0
            mySource.Compile()
        Next


        If Err.Number = 0 Then
            CompileAWL = True
        Else
            CompileAWL = False
        End If

    End Function

    Public Sub FindTodayFolder()
        Dim dateString As String
        Dim dateToday As String
        Dim subFolder As Folder
        Dim myFolder As Folder
        Dim fso As FileSystemObject
        Dim myFile As File

        dateToday = DateTime.Today()
        fso = New FileSystemObject
        For Each subFolder In fso.GetFolder(TextBox1.Text).SubFolders 'subfolder = project
            On Error Resume Next
            myFolder = subFolder.SubFolders.Item("HardwareConfig")
            If Err.Number = 0 Then
                myFile = myFolder.Files.Item("HardwareConfigMain.cfg")
                If Split(myFile.DateLastModified, " ")(0) = dateToday Then
                    TextBox1.Text = subFolder.Path
                    Exit Sub
                End If
            Else
                Err.Clear()
            End If
        Next
    End Sub


    Private Sub FindParentFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindParentFolder.Click
        Dim fldr As New FolderBrowserDialog
        Dim sItem As String

        With fldr
            .Description = "Select a Folder"
            .SelectedPath = "\\srvia03\WorkspaceDevelopment"
            If fldr.ShowDialog() = DialogResult.OK Then
                sItem = fldr.SelectedPath
                TextBox1.Text = sItem
            End If
        End With
NextCode:
        'GetFolder = sItem
        fldr = Nothing
    End Sub


    Private Sub Generate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Generate.Click
        'Alles opnieuw genereren er komt dan automatisch een nieuwe symbol lijst en sources/HW  worden automatisch verwijderd
        Dim numOfProjects As Integer
        Dim myProject As S7Project
        Dim S As Simatic
        Dim projectName As String
        Dim i As Integer
        Dim subFolder As Folder
        Dim present(0 To 3) As Boolean

        fso = Nothing
        fso = New Scripting.FileSystemObject

        txtStatus.Text = "Start Generatie"
        Me.Refresh()

        'Determine project name
        projectName = ""
        If InStr(1, TextBox1.Text, "\") > 0 Then 'check if there is a \ to split on to avoid errors
            'split on '\' and restreive the last part after the last '\', this is your folder name
            projectName = Split(TextBox1.Text, "\")(Len(TextBox1.Text) - Len(Replace(TextBox1.Text, "\", "")))
        End If

        present(0) = False
        present(1) = False
        present(2) = False
        present(3) = False
        'Exit program if sub folders are not present
        For Each subFolder In fso.GetFolder(TextBox1.Text).SubFolders
            If UCase(subFolder.Name) = "HARDWARECONFIG" Then present(0) = True
            If UCase(subFolder.Name) = "HMIGENERATOR" Then present(1) = True
            If UCase(subFolder.Name) = "SOURCECODE" Then present(2) = True
            If UCase(subFolder.Name) = "TAGLIST" Then present(3) = True
            If present(0) And present(1) And present(2) And present(3) Then Exit For
            txtStatus.Text = "Zoek door folder " & subFolder.Name
            Me.Refresh()
        Next

        txtStatus.Text = ""
        Me.Refresh()

        If Not (present(0)) Or Not (present(1)) Or Not (present(2)) Or Not (present(3)) Then
            txtStatus.Text = "controleer mapstructuur " & TextBox1.Text
            Me.Refresh()
            Exit Sub
        End If

        'Delete all previous projects in current Simatic
        S = New Simatic

        On Error Resume Next
        myProject = S.Projects.Item(projectName)
        If Err.Number = 0 Then
            'Err.Clear
            myProject = Nothing
            S.Projects.Remove(projectName)
            If Err.Number <> 0 Then
                txtStatus.Text = "Kan project [" & projectName & "] niet verwijderen"
                Me.Refresh()
                Exit Sub
            End If
        Else
            'project bestaat nog niet
        End If

        m_Project = S.Projects.Add(projectName, "", S7ProjectType.S7Project)
        While m_Project.Stations.Count <> 0
            Dim myStation As S7Station
            myStation = m_Project.Stations.Item(1)
            txtStatus.Text = "verwijder " & myStation.Name & " uit geheugen"
            Me.Refresh()
            m_Project.Stations.Remove(1)
        End While
        txtStatus.Text = "Hardware opzetten"
        Me.Refresh()
        If Not (CreateHardware()) Then Exit Sub
        txtStatus.Text = "Sources importeren"
        Me.Refresh()
        If Not (ImportSources()) Then Exit Sub
        txtStatus.Text = "SCL Sources compileren"
        Me.Refresh()
        If Not (CompileSCLSources()) Then Exit Sub
        txtStatus.Text = "AWL Sources compileren"
        Me.Refresh()
        If Not (CompileAWL()) Then Exit Sub
        txtStatus.Text = "Gereed!"
    End Sub

    Private Sub Step7API_EEC_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim dateString As String
        Dim dateToday As String
        Dim subFolder As Folder
        Dim myFolder As Folder
        Dim fso As FileSystemObject
        Dim myFile As File

        dateToday = Date.Today
        fso = New FileSystemObject
        For Each subFolder In fso.GetFolder(TextBox1.Text).SubFolders 'subfolder = project
            On Error Resume Next
            myFolder = subFolder.SubFolders.Item("HardwareConfig")
            If Err.Number = 0 Then
                myFile = myFolder.Files.Item("HardwareConfigMain.cfg")
                If Split(myFile.DateLastModified, " ")(0) = dateToday Then
                    TextBox1.Text = subFolder.Path
                    Exit Sub
                End If
            Else
                Err.Clear()
            End If
        Next
    End Sub
End Class
