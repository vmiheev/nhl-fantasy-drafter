Attribute VB_Name = "Module1"
Option Explicit
Sub RankingSkaters()
    
    Worksheets("year1").Activate
    
    'Count the players from last year
    Dim PlayerCount As Integer
    
    PlayerCount = Cells(Rows.Count, "A").End(xlUp).Row - 1
    'MsgBox ("There are " & PlayerCount & " skaters.")
    
    'Declare the arrays for player names and stats
    Dim PlayerName() As String
    Dim PlayerTeam() As String
    Dim PlayerPosition() As String
    Dim PlayerGamesYear1() As Integer
    Dim PlayerGoalsYear1() As Integer
    Dim PlayerAssistsYear1() As Integer
    Dim PlayerPlusMinusYear1() As Integer
    Dim PlayerSOGYear1() As Integer
    Dim PlayerPPPointsYear1() As Integer
    Dim PlayerHitsYear1() As Integer
    Dim PlayerGamesYear2() As Integer
    Dim PlayerGoalsYear2() As Integer
    Dim PlayerAssistsYear2() As Integer
    Dim PlayerPlusMinusYear2() As Integer
    Dim PlayerSOGYear2() As Integer
    Dim PlayerPPPointsYear2() As Integer
    Dim PlayerHitsYear2() As Integer
    Dim PlayerGamesYear3() As Integer
    Dim PlayerGoalsYear3() As Integer
    Dim PlayerAssistsYear3() As Integer
    Dim PlayerPlusMinusYear3() As Integer
    Dim PlayerSOGYear3() As Integer
    Dim PlayerPPPointsYear3() As Integer
    Dim PlayerHitsYear3() As Integer
    
    ReDim PlayerName(PlayerCount)
    ReDim PlayerTeam(PlayerCount)
    ReDim PlayerPosition(PlayerCount)
    ReDim PlayerGamesYear1(PlayerCount)
    ReDim PlayerGoalsYear1(PlayerCount)
    ReDim PlayerAssistsYear1(PlayerCount)
    ReDim PlayerPlusMinusYear1(PlayerCount)
    ReDim PlayerSOGYear1(PlayerCount)
    ReDim PlayerPPPointsYear1(PlayerCount)
    ReDim PlayerHitsYear1(PlayerCount)
    ReDim PlayerGamesYear2(PlayerCount)
    ReDim PlayerGoalsYear2(PlayerCount)
    ReDim PlayerAssistsYear2(PlayerCount)
    ReDim PlayerPlusMinusYear2(PlayerCount)
    ReDim PlayerSOGYear2(PlayerCount)
    ReDim PlayerPPPointsYear2(PlayerCount)
    ReDim PlayerHitsYear2(PlayerCount)
    ReDim PlayerGamesYear3(PlayerCount)
    ReDim PlayerGoalsYear3(PlayerCount)
    ReDim PlayerAssistsYear3(PlayerCount)
    ReDim PlayerPlusMinusYear3(PlayerCount)
    ReDim PlayerSOGYear3(PlayerCount)
    ReDim PlayerPPPointsYear3(PlayerCount)
    ReDim PlayerHitsYear3(PlayerCount)
    
    'Loop through players in the year1 worksheet and save their names and stats into arrays
    Dim i As Integer
    
    For i = 0 To (PlayerCount - 1)
        
        PlayerName(i) = Cells(i + 2, 1).Value
        PlayerTeam(i) = Cells(i + 2, 2).Value
        PlayerPosition(i) = Cells(i + 2, 3).Value
        PlayerGamesYear1(i) = Cells(i + 2, 4).Value
        PlayerGoalsYear1(i) = Cells(i + 2, 5).Value
        PlayerAssistsYear1(i) = Cells(i + 2, 6).Value
        PlayerPlusMinusYear1(i) = Cells(i + 2, 8).Value
        PlayerSOGYear1(i) = Cells(i + 2, 10).Value
        PlayerPPPointsYear1(i) = Cells(i + 2, 12).Value + Cells(i + 2, 13).Value
        PlayerHitsYear1(i) = Cells(i + 2, 16).Value
        
    Next i
    
    'Count the players from year 2
    Worksheets("year2").Activate
    
    Dim PlayerCount2 As Integer
    
    PlayerCount2 = Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    'Loop over the players from year 2, and if they played last year then save their stats into the appropriate array slot
    Dim j As Integer
    
    For j = 0 To (PlayerCount2 - 1)
        
        For i = 0 To (PlayerCount - 1)
            
            If Cells(j + 2, 1).Value = PlayerName(i) Then
                
                PlayerGamesYear2(i) = Cells(j + 2, 4).Value
                PlayerGoalsYear2(i) = Cells(j + 2, 5).Value
                PlayerAssistsYear2(i) = Cells(j + 2, 6).Value
                PlayerPlusMinusYear2(i) = Cells(j + 2, 8).Value
                PlayerSOGYear2(i) = Cells(j + 2, 10).Value
                PlayerPPPointsYear2(i) = Cells(j + 2, 12).Value + Cells(j + 2, 13).Value
                PlayerHitsYear2(i) = Cells(j + 2, 16).Value
            
            End If
            
        Next i
        
    Next j
    
    'Count the players from year 2
    Worksheets("year3").Activate
    
    Dim PlayerCount3 As Integer
    
    PlayerCount3 = Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    'Loop over the players from year 3, and if they played last year then save their stats into the appropriate array slot
    
    For j = 0 To (PlayerCount3 - 1)
        
        For i = 0 To (PlayerCount - 1)
            
            If Cells(j + 2, 1).Value = PlayerName(i) Then
                
                PlayerGamesYear3(i) = Cells(j + 2, 4).Value
                PlayerGoalsYear3(i) = Cells(j + 2, 5).Value
                PlayerAssistsYear3(i) = Cells(j + 2, 6).Value
                PlayerPlusMinusYear3(i) = Cells(j + 2, 8).Value
                PlayerSOGYear3(i) = Cells(j + 2, 10).Value
                PlayerPPPointsYear3(i) = Cells(j + 2, 12).Value + Cells(j + 2, 13).Value
                PlayerHitsYear3(i) = Cells(j + 2, 16).Value
                
            End If
            
        Next i
        
    Next j

    'Create column headers for Rankings tab
    Worksheets("Rankings").Activate
    
    Cells(1, 1).Value = "Name"
    Cells(1, 2).Value = "Team"
    Cells(1, 3).Value = "Position"
    Cells(1, 4).Value = "Games"
    Cells(1, 5).Value = "G"
    Cells(1, 6).Value = "A"
    Cells(1, 7).Value = "+/-"
    Cells(1, 8).Value = "SOG"
    Cells(1, 9).Value = "PPP"
    Cells(1, 10).Value = "Hits"
    
    'Print names and stats into Rankings tab
    For i = 0 To (PlayerCount - 1)
        
        Cells(i + 2, 1).Value = PlayerName(i)
        Cells(i + 2, 2).Value = PlayerTeam(i)
        Cells(i + 2, 3).Value = PlayerPosition(i)
        Cells(i + 2, 4).Value = PlayerGamesYear1(i)
        Cells(i + 2, 5).Value = PlayerGoalsYear1(i)
        Cells(i + 2, 6).Value = PlayerAssistsYear1(i)
        Cells(i + 2, 7).Value = PlayerPlusMinusYear1(i)
        Cells(i + 2, 8).Value = PlayerSOGYear1(i)
        Cells(i + 2, 9).Value = PlayerPPPointsYear1(i)
        Cells(i + 2, 10).Value = PlayerHitsYear1(i)
        Cells(i + 2, 11).Value = PlayerGoalsYear2(i)
        Cells(i + 2, 12).Value = PlayerGoalsYear3(i)
    
    Next i
    
End Sub
