Imports System.IO
Imports System.Speech.Recognition
Imports System.Speech.Synthesis
Imports System.Text


Public Class Form1

    Private csvFiles As List(Of String)
    'Weather Search Engine
    Private ReadOnly weather As New SpeechRecognitionEngine()

    Private ReadOnly cast As New DictationGrammar()

    Private ReadOnly word As New SpeechSynthesizer()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Interval = 12000
        Timer1.Start()
        Timer2.Start()
        Timer2.Interval = 8000
        AxWindowsMediaPlayer1.Visible = True

        AxWindowsMediaPlayer1.URL = "C:\Users\justin.ross\source\repos\Guardian Watch\Audio\raven.mp4"

        weather.LoadGrammarAsync(cast)
        weather.RequestRecognizerUpdate()
        weather.SetInputToDefaultAudioDevice()
        weather.InitialSilenceTimeout = TimeSpan.FromSeconds(2.5)
        weather.BabbleTimeout = TimeSpan.FromSeconds(1.5)
        weather.EndSilenceTimeout = TimeSpan.FromSeconds(1.2)
        weather.EndSilenceTimeoutAmbiguous = TimeSpan.FromSeconds(1.5)
        AddHandler weather.SpeechRecognized, AddressOf Weather_SpeechRecognized
        'weather.RecognizeAsync(RecognizeMode.Multiple)

        FileSystemWatcher1.Path = My.Computer.FileSystem.SpecialDirectories.Desktop
        FileSystemWatcher2.Path = My.Computer.FileSystem.SpecialDirectories.MyDocuments
    End Sub

    Private Sub FileSystemWatcher1_Created(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher1.Created
        ListBox1.Items.Add(e.FullPath.ToString)
        My.Computer.Audio.Play(AppDomain.CurrentDomain.BaseDirectory & "../../Audio/alert.wav", AudioPlayMode.Background)
    End Sub

    Private Sub FileSystemWatcher2_Created(sender As Object, e As FileSystemEventArgs) Handles FileSystemWatcher2.Created
        ListBox1.Items.Add(e.FullPath.ToString)
        My.Computer.Audio.Play(AppDomain.CurrentDomain.BaseDirectory & "../../Audio/alert.wav", AudioPlayMode.Background)
    End Sub


    Private Sub DeleteThreat_Click(sender As Object, e As EventArgs) Handles DeleteThreat.Click
        Dim li As Integer
        Dim fn As String

        For li = ListBox1.SelectedIndices.Count - 1 To 0 Step -1

            fn = ListBox1.Items(ListBox1.SelectedIndices(li))

            If File.Exists(fn) Then

                File.Delete(fn)

            End If

            ListBox1.Items.RemoveAt(ListBox1.SelectedIndices(li))
        Next
        My.Computer.Audio.Play(AppDomain.CurrentDomain.BaseDirectory & "../../Audio/threat.wav", AudioPlayMode.Background)
    End Sub


    Private Sub EnableMic_Click(sender As Object, e As EventArgs) Handles EnableMic.Click
        weather.RecognizeAsync(RecognizeMode.Multiple)
    End Sub

    Private Sub DisableMic_Click(sender As Object, e As EventArgs) Handles DisableMic.Click
        weather.RecognizeAsyncStop()
    End Sub

    Private Sub SystemCleaner()
        Process.Start(AppDomain.CurrentDomain.BaseDirectory & "../../Clean.bat")

        'Clear my internet & system tracks

        For i As Integer = 0 To 1
            Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 255", vbNormalFocus)
        Next
        Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 4351", vbNormalFocus)
        For i As Integer = 0 To 1
            Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 2", vbNormalFocus)
        Next
        Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 16384", vbNormalFocus)
        For i As Integer = 0 To 1
            Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 16", vbNormalFocus)
        Next
        Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 1", vbNormalFocus)
        For i As Integer = 0 To 1
            Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 32", vbNormalFocus)
        Next
        Shell("rundll32.exe InetCpl.cpl,ClearMyTracksByProcess 8", vbNormalFocus)

        Dim reg4 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Internet Explorer", True)
        reg4.CreateSubKey("TypedURLs") '//Typedurls-Url typed
        reg4.DeleteSubKeyTree("TypedURLs") '//delete subkey and all subkeys
        '//in it      
        Dim reg6 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer", True)
        reg6.CreateSubKey("RecentDocs") '//Recentdocs-Recent documents opened
        reg6.DeleteSubKeyTree("RecentDocs") '//delete subkey and all subkeys
        '//in it
        Dim reg7 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer", True)
        reg7.CreateSubKey("ComDlg32") '//Comdlg32-Common dialog
        reg7.DeleteSubKeyTree("ComDlg32") '//delete subkey and all subkeys in
        '//it
        Dim reg8 As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Microsoft\Windows\CurrentVersion\Explorer", True)
        reg8.CreateSubKey("StreamMRU")         '//StreamMru-Stream history.
        reg8.DeleteSubKeyTree("StreamMRU") '//delete subkey and all subkeys
        '//in it
        On Error GoTo err
        Dim di As New DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.InternetCache))

        ' Create the directory only if it does not already exist.
        If di.Exists = False Then
            di.Create()
        End If

        File.SetAttributes(Environment.GetFolderPath(
            Environment.SpecialFolder.InternetCache).ToString,
            FileAttributes.Normal)
        Dim Cache1 As String
        Dim Cache2() As String

        Cache2 = Directory.GetFiles(Environment.GetFolderPath(
            Environment.SpecialFolder.InternetCache))
        For Each Cache1 In Cache2 '//Get all files in Temporary internet

            File.SetAttributes(Cache1, FileAttributes.Normal)
            File.Delete(Cache1)
        Next

        di.Delete(True)

        Dim dm As New DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.History))

        ' Create the directory only if it does not already exist.
        If dm.Exists = False Then
            dm.Create()
        End If

        File.SetAttributes(Environment.GetFolderPath(
           Environment.SpecialFolder.History).ToString, FileAttributes.Normal)
        Dim history1 As String
        Dim history2() As String
        history2 = Directory.GetFiles(Environment.GetFolderPath(
            Environment.SpecialFolder.History))
        For Each history1 In history2 '//Get all files in history folder and
            '//then set their attribute to normal, and then delete them.
            File.SetAttributes(history1, FileAttributes.Temporary)
            File.Delete(history1)
        Next

        dm.Delete(True)

        Dim dr As New DirectoryInfo(
            Environment.GetFolderPath(Environment.SpecialFolder.Recent))

        ' Create the directory only if it does not already exist.
        If dr.Exists = False Then
            dr.Create()
        End If

        File.SetAttributes(Environment.GetFolderPath(
            Environment.SpecialFolder.Recent).ToString, FileAttributes.Normal)
        Dim recentk As String
        Dim Recentl() As String

        Recentl = Directory.GetFiles(Environment.GetFolderPath(
            Environment.SpecialFolder.Recent))
        For Each recentk In Recentl '//Get all files in recent folder and then
            '//set their attribute to normal, and then delete them.
            File.SetAttributes(recentk, FileAttributes.Normal)
            File.Delete(recentk)
        Next

        dr.Delete(True)

        Dim ex As New DirectoryInfo(Environment.GetFolderPath(
            Environment.SpecialFolder.Cookies))

        ' Create the directory only if it does not already exist.
        If ex.Exists = False Then
            ex.Create()
        End If

        File.SetAttributes(Environment.GetFolderPath(
            Environment.SpecialFolder.Cookies).ToString, FileAttributes.Normal)
        Dim Cookie1 As String '//Cookie1
        Dim Cookie2() As String '//Cookie2
        Cookie2 = Directory.GetFiles(
            Environment.GetFolderPath(Environment.SpecialFolder.Cookies))
        For Each Cookie1 In Cookie2 '//Get all files that have .txt extension
            '//and set their attribute to normal and then delete them.
            If InStr(Cookie1, ".txt", CompareMethod.Text) Then
                File.SetAttributes(Cookie1, FileAttributes.Normal)
                File.Delete(Cookie1)
            End If
        Next

        ex.Delete(True)

        Dim ax As New DirectoryInfo("C:\Documents and Settings\" +
            Environment.UserName + "\Local Settings\Temp")

        ' Create the directory only if it does not already exist.
        If ax.Exists = False Then
            ax.Create()
        End If

        File.SetAttributes("C:\Documents and Settings\" + Environment.UserName + "\Local Settings\Temp", FileAttributes.Normal)
        Dim temp1 As String
        Dim temp2() As String
        temp2 = Directory.GetFiles("C:\Documents and Settings\" +
            Environment.UserName + "\Local Settings\Temp")
        For Each temp1 In temp2 '//Get all files in Temp folder and then set
            '//their attribute to normal, and then delete them.
            File.SetAttributes(temp1, FileAttributes.Normal)
            File.Delete(temp1)
        Next
        ' The true indicates that if subdirectories
        ' or files are in this directory, they are to be deleted as well.

        ' Delete the directory.
        ax.Delete(True)

        If File.Exists("C:\Windows\System32\cleanmgr.exe") Then
            Dim Drive As String = "C"
            Dim _process = New Process()

            With _process
                .StartInfo.FileName = "cleanmgr.exe"
                .StartInfo.WorkingDirectory = "C:\Windows\System32"
                .StartInfo.Arguments = "/verylowdisk"
                .EnableRaisingEvents = True
                .StartInfo.CreateNoWindow = False
                .StartInfo.UseShellExecute = False
                .StartInfo.WindowStyle = ProcessWindowStyle.Hidden
                .StartInfo.RedirectStandardOutput = True
                .StartInfo.RedirectStandardError = True
                .Start()
                .WaitForExit()
            End With
        Else
            MessageBox.Show("Failed to find Disk cleaner.")
        End If
err:
    End Sub

    Private Sub DeleteHistory_Click(sender As Object, e As EventArgs) Handles DeleteHistory.Click
        SystemCleaner()
    End Sub

    Private Sub SysInfo_Click(sender As Object, e As EventArgs) Handles SysInfo.Click
        Using proc As New Process()

            proc.StartInfo.FileName = "msinfo32"
            proc.StartInfo.Arguments = "/pch"
            proc.StartInfo.RedirectStandardOutput = True
            proc.StartInfo.UseShellExecute = False
            proc.Start()
        End Using
    End Sub


    'Weather Engine
    Private Sub Weather_SpeechRecognized(sender As Object, e As SpeechRecognizedEventArgs)
        Dim weather As String = e.Result.Text.ToString

        Select Case (weather)
            Case weather

                If weather <> "+" Then

                    'This section will take spoken word, write to file then execute search.
                    Dim sb As New StringBuilder
                    'Be sure to change the text file path below to your path if you are new to this program.
                    sb.AppendLine(weather)
                    File.WriteAllText("C:\Users\justin.ross\source\repos\Guardian Watch\creation.txt", sb.ToString())
                    weather = "https://www.youtube.com/search?q= " + " " & Uri.EscapeUriString(weather)

                    Dim proc As New Process()
                    Dim startInfo As New ProcessStartInfo(weather)
                    proc.StartInfo = startInfo
                    proc.Start()
                    SendKeys.Send($"^{{w}}")
                    Return
                End If

        End Select

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        FileSystemWatcher1.Dispose()

    End Sub

    Private Sub ClearItems_Click(sender As Object, e As EventArgs) Handles ClearItems.Click
        ListBox1.Items.Clear()
    End Sub

    Private Sub SaveList_Click(sender As Object, e As EventArgs) Handles SaveList.Click
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim sb As New StringBuilder()
            For Each o As Object In ListBox1.Items
                sb.AppendLine(o)
            Next
            File.WriteAllText($"{Environ$("USERPROFILE")}\Desktop\SavedItems.txt", sb.ToString())
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        AxWindowsMediaPlayer1.Visible = False
        Timer1.Stop()
    End Sub

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Timer2.Stop()
        My.Computer.Audio.Play(AppDomain.CurrentDomain.BaseDirectory & "../../Audio/guardian.wav", AudioPlayMode.Background)
    End Sub

    Public Sub New()
        ' This call is required by the Windows Form Designer.
        InitializeComponent()
        ' display the list of running processes
        UpdateProcessList()
    End Sub

    Private Sub KillProcess_Click(sender As Object, e As EventArgs) Handles KillProcess.Click
        My.Computer.Audio.Play(AppDomain.CurrentDomain.BaseDirectory & "../../Audio/destroyprocess.wav", AudioPlayMode.Background)
        If ListBox1.SelectedItems.Count <= 0 Then
            MessageBox.Show("Click on a process name to select it.", "No Process Selected")
            Return
        End If
        ' loop through the running processes looking for a match
        ' by comparing process name to the name selected in the listbox
        Dim p As Process
        For Each p In Process.GetProcesses()
            Dim arr() As String =
            ListBox1.SelectedItem.ToString().Split(CChar("-"))
            Dim sProcess As String = arr(0).Trim()
            Dim iId As Integer = Convert.ToInt32(arr(1).Trim())
            If p.ProcessName = sProcess And p.Id = iId Then

                p.Kill()

            End If
        Next
        ' update the list to show the killed process
        ' has been removed
        UpdateProcessList()
    End Sub

    Private Sub UpdateList_Click(sender As Object, e As EventArgs) Handles UpdateList.Click
        UpdateProcessList()
    End Sub

    Private Sub UpdateProcessList()
        ' clear the existing list of any items
        ListBox1.Items.Clear()
        ' loop through the running processes and add
        'each to the list
        Dim p As Process
        For Each p In Process.GetProcesses()
            ListBox1.Items.Add(p.ProcessName & " - " & p.Id.ToString())
        Next
        ListBox1.Sorted = True
        ' display the number of running processes in
        ' a status message at the bottom of the page 
        ListBox1.Text = "Processes running: " &
        ListBox1.Items.Count.ToString()
    End Sub


    Private Sub LoadListBox(path As String)
        Dim di As New DirectoryInfo(path)
        di.EnumerateFiles("*.*", SearchOption.AllDirectories)
        csvFiles = (From csv In di.EnumerateFiles("*.*", SearchOption.AllDirectories)
                    Where csv.CreationTime.Date >= DateTimePicker1.Value.Date AndAlso
                     csv.CreationTime.Date <= DateTimePicker2.Value.Date
                    Select csv.FullName).ToList

        ListBox1.DataSource = csvFiles
    End Sub

    Private Sub FindNewFiles_Click(sender As Object, e As EventArgs) Handles FindNewFiles.Click
        Dim path As String = My.Computer.FileSystem.SpecialDirectories.Temp
        LoadListBox(path)
    End Sub
End Class
