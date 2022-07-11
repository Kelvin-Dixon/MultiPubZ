Imports System.Text.RegularExpressions
Module MultiPubZ

    Public Sub Main(ByVal args() As String)
        If args.Length > 0 Then ' Check for Filename Supplied
            Dim TestFile As String = args(0)
            Dim InputFolder As String = My.Settings.InputFolder + "\"
            Dim ErrorFolder As String = My.Settings.ErrorFolder + "\"
            Dim MatchedFolder As String = My.Settings.MatchedFolder + "\"
            Dim Publist() = My.Settings.PubList.Split(",")
            Dim DaysToAdd As Integer = My.Settings.DaysToAdd
            Dim PubToMatch As String = My.Settings.PubToMatch
            Dim DateFormat As String = "yyyyMMdd"
            Dim SrcFile As String = InputFolder + TestFile
            Dim DstFile As String
            'Console.WriteLine("InputFolder: " & InputFolder)
            'Console.WriteLine("ErrorFolder: " & ErrorFolder)
            'Console.WriteLine("MatchedFolder: " & MatchedFolder)
            'Console.WriteLine("PubList: " & Publist)
            'Console.WriteLine("FileName Supplied")
            'Console.WriteLine(testFileName)
            If FileIO.FileSystem.FileExists(SrcFile) Then 'Is this a real File?
                Dim Result = Regex.Match(TestFile, "(?<Publication>.+)-(?<PubDate>\d{4}\d{2}\d{2})-(?<Section>Z)(?<Pages>\d{2}).pdf", RegexOptions.IgnoreCase)
                If Result.Success Then ' Filename Matches Pattern
                    'Console.WriteLine("Filename Matches")
                    Dim Publication = Result.Groups.Item("Publication").Value
                    If Publication = PubToMatch Then
                        Dim CurrentDate = Result.Groups.Item("PubDate").Value
                        Dim FutureDate = Date.ParseExact(CurrentDate, DateFormat, Globalization.CultureInfo.InvariantCulture).AddDays(DaysToAdd).ToString(DateFormat)
                        Dim Section = Result.Groups.Item("Section").Value
                        Dim Pages = Result.Groups.Item("Pages").Value
                        Dim NewName As String = ""
                        'Console.WriteLine("Publication: {0}", Publication)
                        'Console.WriteLine("CurrentDate: {0}", CurrentDate)
                        'Console.WriteLine("FutureDate: {0}", FutureDate)
                        'Console.WriteLine("Section: {0}", Section)
                        'Console.WriteLine("Pages: {0}", Pages)
                        For Each Pub As String In Publist
                            NewName = Pub + "-" + FutureDate + "-" + Section + Pages + ".pdf"
                            DstFile = MatchedFolder + NewName
                            'Console.WriteLine("Copyto: {0}", DstFile)
                            FileIO.FileSystem.CopyFile(SrcFile, DstFile, True)
                        Next
                        DstFile = MatchedFolder + TestFile
                        'Console.WriteLine("Moveto: {0}", DstFile)
                        FileIO.FileSystem.MoveFile(SrcFile, DstFile, True)
                    Else
                        'Console.WriteLine("Wrong Publication")
                        DstFile = ErrorFolder + TestFile
                        'Console.WriteLine("Moveto: {0}", DstFile)
                        FileIO.FileSystem.MoveFile(SrcFile, DstFile, True)
                    End If
                Else
                    'Console.WriteLine("Filename Does Not Match")
                    DstFile = ErrorFolder + TestFile
                    'Console.WriteLine("Moveto: {0}", DstFile)
                    FileIO.FileSystem.MoveFile(SrcFile, DstFile, True)
                End If
            Else
                Console.WriteLine("Filename Does not Exist")
            End If
        Else
            Console.WriteLine("Filename Not Specified")
        End If
    End Sub
End Module
