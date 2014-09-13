Inport dll here
Imports System.IO
Imports System.Reflection
Imports System.Xml
Imports System.Text
Imports System.Configuration
Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class Form1
    

    Dim sbXML As New StringBuilder()

    Dim writer As XmlWriter
    Dim settings As New XmlWriterSettings()
    Dim _filePath As String
    Dim dataAccess As New DataAccess()

    Dim SaveInDBAndCreateMSGFile As Boolean = CType(ConfigurationManager.AppSettings("SaveInDBAndCreateMSGFile"), Boolean)

    Public Property FilePath() As String
        Get
            Return _filePath
        End Get
        Set(ByVal value As String)
            _filePath = value
        End Set
    End Property

    Public Property ServerPath() As String
        Get
            Return _serverPath
        End Get
        Set(ByVal value As String)
            _serverPath = value
        End Set
    End Property

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        AsposeBackgroundWorker = New BackgroundWorker()
        AsposeBackgroundWorker.WorkerReportsProgress = True
        AsposeBackgroundWorker.WorkerSupportsCancellation = True

    End Sub

    Private Sub setLicense()
        Try
            SetLicense
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ExitApplication()
        End Try

    End Sub

    Private Sub ExitApplication()
        Application.Exit()
        Environment.Exit(1)
    End Sub

    Private Sub DisplayPSTFilesRecursively(ByVal e As DoWorkEventArgs)
        Try
            If File.Exists(pstFilePath) Then
                ' Load the Outlook PST file
                Dim pst As PersonalStorage = PersonalStorage.FromFile(pstFilePath)

                ' Get the folders and messages information
                Dim folderInfo As FolderInfo = pst.RootFolder

                DisplayFolderContents(folderInfo, pst, pst.DisplayName, e)
            Else
                MessageBox.Show("File does not exist under path " & pstFilePath)
                ExitApplication()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub


    Private Sub DisplayFolderContents(ByVal folderInfo As FolderInfo, ByVal pst As PersonalStorage,
                                      ByVal displayName As String, ByVal e As DoWorkEventArgs)
        Dim messageInfoCollection As MessageInfoCollection = folderInfo.GetContents()
        For Each messageInfo As MessageInfo In messageInfoCollection
            Try
                ProcessMessages(pst, messageInfo, displayName)
            Catch ex As Exception
                Continue For ' continue for loop on any exception in "ProcessMessages" function
            End Try
        Next messageInfo
        ' Call this method recursively for each subfolder
        If folderInfo.HasSubFolders() = True Then
            Dim TotalFolderCount As Integer = folderInfo.GetSubFolders().Count

            For i As Integer = 0 To TotalFolderCount - 1
                Dim subfolderInfo As FolderInfo = DirectCast(folderInfo.GetSubFolders().Item(i), FolderInfo)
                Dim name As String = folderInfo.GetSubFolders().Item(i).DisplayName

                'Periodically report progress to the main thread so that it can update the UI.
                AsposeBackgroundWorker.ReportProgress((i / TotalFolderCount) * 100, New ProgressAdditionalInfo(TotalFolderCount, name))

                If AsposeBackgroundWorker.CancellationPending Then
                    e.Cancel = True
                    AsposeBackgroundWorker.ReportProgress(0)
                    Return
                End If
                DisplayFolderContents(subfolderInfo, pst, name, e)
            Next
        End If
    End Sub


    Private Sub ProcessMessages(ByVal pst As PersonalStorage, ByVal msgInfo As MessageInfo,
                                ByVal displayName As String)
        Dim message As MapiMessage = pst.ExtractMessage(msgInfo)
        Dim formattedString As String = message.Subject.Replace(":", "-") _
                                                       .Replace("/", "_") _
                                                       .Replace("\", "_") _
                                                       .Replace("*", "_") _
                                                       .Replace("?", "_") _
                                                       .Replace("<", "_") _
                                                       .Replace(">", "_") _
                                                       .Replace("|", "_") _
                                                       .Replace("#", "No")

        If Not Directory.Exists(FolderToExtractMSGFiles) Then
            Directory.CreateDirectory(FolderToExtractMSGFiles)
        End If

        'DEBUG ASSERT WHILE TESTING 
        'Debug.Assert(displayName.IndexOf("0083730") = -1) 'for testing
        
        Dim actualDisplayName As String = displayName
        Dim success As Boolean = False
        Dim junkWords As Boolean = False

        If displayName.Trim().ToLower().IndexOf("deleted") >= 0 Or displayName.Trim().ToLower().IndexOf("pending") >= 0 Or
                displayName.Trim().ToLower().IndexOf("assessment") >= 0 Or displayName.Trim().ToLower().IndexOf("CRE") >= 0 Then
            junkWords = True
        End If

        If Not junkWords Then
            Dim match As Match = Regex.Match(displayName, "^[\d-\d]+") 'Ex => 64-3444
            If match.Success Then
                displayName = match.Value.Trim()
                success = True
            End If

            If Not success Then 'Ex => C-2011-032
                Dim match4 As Match = Regex.Match(displayName, "^[A-Z][-\d-\d]+")
                If match4.Success Then
                    displayName = match4.Value.Trim()
                    success = True
                End If
            End If

            If Not success Then
                Dim match1 As Match = Regex.Match(displayName, "^[\sA-Za-z0-9]+", RegexOptions.IgnoreCase)
                If match1.Success Then
                    displayName = match1.Value.Split("-"c)(0).Trim().ToString()
                    displayName = displayName.Split(" ")(0).Trim().ToString()
                    success = True
                End If
            End If

            If Not success Then 'Ex => P743473
                Dim match2 As Match = Regex.Match(displayName, "^([P]{1}\d{2,20})+", RegexOptions.IgnoreCase)
                If match2.Success Then
                    displayName = match2.Value.Split("-"c)(0).Trim().ToString()
                    displayName = displayName.Split(" ")(0).Trim().ToString()
                    success = True
                End If
            End If

            If Not success Then
                Dim match3 As Match = Regex.Match(displayName, "^(\d{1,15})+") 'Ex => 345455
                If match3.Success Then
                    displayName = match3.Value.Split("-"c)(0).Trim().ToString()
                    success = True
                End If
            End If

            If success Then
                Dim emailFolder As String = FolderToExtractMSGFiles & "\" & displayName

                If Not Directory.Exists(emailFolder) Then
                    Directory.CreateDirectory(emailFolder)
                End If

                Dim formatNew As String

                If formattedString.Length >= 10 Then
                    formatNew = formattedString.Substring(0, 10) & "..." & Guid.NewGuid().ToString().Substring(0, 10) & ".msg"
                Else
                    formatNew = formattedString.Substring(0, formattedString.Length) & "..." & Guid.NewGuid().ToString().Substring(0, 10) & ".msg"
                End If

                FilePath = emailFolder & "\" & formatNew

                Dim srvPath As String = ServerPath & displayName & "/" & formatNew

                If SaveInDBAndCreateMSGFile Then
                    
                End If
            End If
        End If
    End Sub

    Private Sub btnStart_Click(sender As System.Object, e As System.EventArgs) Handles btnStart.Click
        Try
            setLicense()
            If Not AsposeBackgroundWorker.IsBusy Then
                Me.btnStart.Enabled = False
                Me.btnCancel.Enabled = True

                lblProgress.Text = "Calculations in process..."

                ShowProgressNotificationOnTaskBar()

                'Kickoff the worker thread to begin it's DoWork function.
                AsposeBackgroundWorker.RunWorkerAsync()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            ExitApplication()
        End Try
    End Sub

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        If AsposeBackgroundWorker.IsBusy Then

            'Notify the worker thread that a cancel has been requested.
            'The cancel will not actually happen until the thread in the
            'DoWork checks the m_oWorker.CancellationPending flag. 
            AsposeBackgroundWorker.CancelAsync()
        End If
    End Sub

    Private Sub ShowProgressNotificationOnTaskBar()
        'Me.Hide()
        Me.ShowInTaskbar = True

        If NotifyIcon1.BalloonTipText = String.Empty Then
            NotifyIcon1 = New System.Windows.Forms.NotifyIcon()
            NotifyIcon1.BalloonTipText = "0%"
            NotifyIcon1.Icon = New Drawing.Icon("App.ico")
            NotifyIcon1.Visible = True

        End If
    End Sub

    Private Sub AsposeBackgroundWorker_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles AsposeBackgroundWorker.DoWork
        Dim worker As BackgroundWorker = TryCast(sender, BackgroundWorker)
        DisplayPSTFilesRecursively(e)
    End Sub

    Private Sub AsposeBackgroundWorker_ProgressChanged(sender As System.Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles AsposeBackgroundWorker.ProgressChanged
        progress.Value = e.ProgressPercentage
        Dim progressValue As String = "Processing....." & progress.Value.ToString() & "%"

        'lblProgressPercentage.Text = progressValue
        Me.NotifyIcon1.ShowBalloonTip(360000, "MSG Extraction Progress", progressValue, ToolTipIcon.Info)

        If TypeOf e.UserState Is ProgressAdditionalInfo Then
            Dim pai As ProgressAdditionalInfo = CType(e.UserState, ProgressAdditionalInfo)
            lblProgress.Text = "Folder Name = " & pai.DisplayNameProgress
            lblFolderCount.Text = "Total Folder Count : " & pai.FolderCount.ToString()
        End If
    End Sub

    Private Sub AsposeBackgroundWorker_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles AsposeBackgroundWorker.RunWorkerCompleted
        If (e.Cancelled = True) Then
            Me.lblProgress.Text = "PST File Processing canceled!!!"
            lblProgress.ForeColor = Color.Red
            If MessageBox.Show("Processing is Cancelled. Press OK to close the window and Exit. ") = DialogResult.OK Then
                ExitApplication()
            End If
        ElseIf Not (e.Error Is Nothing) Then
            MessageBox.Show(e.Error.Message)
            ExitApplication()
        Else
            'Everything completed normally.
            Me.lblProgress.Text = "PST File Processing Completed..."
            lblProgress.ForeColor = Color.Green
            progress.Value = 0 ' reset progress back to 0 after task completion
            'lblProgressPercentage.Text = "Processing.....0%"
            ShowWindow()

            If MessageBox.Show("Processing is Completed. Press OK to close the window and Exit. ") = DialogResult.OK Then
                ExitApplication()
            End If
        End If

        Me.btnStart.Enabled = True
        Me.btnCancel.Enabled = False
    End Sub

    Private Sub ShowWindow()
        Me.Show()
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub Form1_Resize(sender As System.Object, e As System.EventArgs) Handles MyBase.Resize
        If WindowState = FormWindowState.Minimized Then
            ShowProgressNotificationOnTaskBar()
        End If
    End Sub
End Class
