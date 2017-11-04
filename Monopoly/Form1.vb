Imports System.Math
Imports System.Collections.Generic
Imports System.Security.Cryptography
Imports System.IO
Imports System.Text


Public Class MonopolyMain
	Dim selectedPiecesSetup(7, 0) As String
	Dim DefaultPieces(8) As String
	Dim UpdatingPlayersSetup As Boolean
	Dim Pausecolour As System.Drawing.Color
	Dim CurrentHouseRules As HouseRules
	Dim PlayerNumber As Byte
	Dim CurrentPlayers() As Player
	Dim Tiles(39) As _Tile
	Dim TurnOrderCurrentTurn As Integer
	Dim TurnorderRolls(7) As Integer
	Dim rollable1 As Boolean
	Dim rollable2 As Boolean
	Dim die1amt As Integer
	Dim die2amt As Integer
	Dim Turn As Byte
	Dim GameDie1Target As Integer
	Dim GameDie2Target As Integer
	Dim rolling As Boolean
	Dim doublecount As Integer
	Dim FRed As Integer
	Dim FGreen As Integer
	Dim FBlue As Integer
	Dim movemax As Integer
	Dim CounterForTurnsInGame As Integer
	Dim rolledAmt As Integer
	Dim MoneyOnFreeParking As Integer
	Dim ChanceCards As New List(Of String)
	Dim ComChestCards As New List(Of String)
	Dim gamedie1active As Boolean
	Dim gamedie2active As Boolean
	Dim moveractive As Boolean
	Dim flasheractive As Boolean
	Dim SavedFileNames(50) As String
	Dim Trader1Cash As Integer
	Dim Trader2Cash As Integer
	Dim Trader1Cards As Integer
	Dim Trader2Cards As Integer
	Dim Trader2ID As Integer








	Structure _Rent
		Dim Base As Integer

		Dim OneHouse As Integer
		Dim TwoHouse As Integer
		Dim ThreeHouse As Integer
		Dim FourHouse As Integer
		Dim Hotel As Integer

		Dim OneStation As Integer
		Dim TwoStation As Integer
		Dim ThreeStation As Integer
		Dim FourStation As Integer

		Dim OneUtility As Integer
		Dim TwoUtility As Integer

		Dim TaxAmount As Integer



	End Structure
	Structure _Tile
		Dim GO As Boolean
		Dim Jail As Boolean
		Dim FreeParking As Boolean
		Dim Tax As Boolean
		Dim Station As Boolean
		Dim Utility As Boolean
		Dim GoToJail As Boolean
		Dim Chance As Boolean
		Dim ComChest As Boolean
		Dim IsProperty As Boolean

		Dim Title As String
		Dim Cost As String
		Dim MortgageValue As Integer
		Dim Colour As System.Drawing.Color
		Dim Rent As _Rent
		Dim HouseNo As Byte
		Dim Hotel As Boolean
		Dim HouseCost As Integer
		Dim HotelCost As Integer
		Dim Owner As String
		Dim Mortgaged As Boolean


	End Structure

	Structure Player
		Dim Name As String
		Dim Piece As String
		Dim Human As Boolean
		Dim Cash As Integer
		Dim GetOutJailCards As Integer
		Dim InJail As Boolean
		Dim TurnsInJail As Byte
		Dim TurnOrderRolledAmt As Integer
		Dim CurrentPos As Integer
		Dim Bankrupt As Boolean
	End Structure

	Structure HouseRules
		Dim FreeParkingMoney As Boolean
		Dim FreeParkingTaxCollect As Boolean
		Dim StartingCash As Integer
		Dim GOmulitplier As Byte
		Dim JailRent As Boolean
	End Structure

	Private Sub InitialiseChanceCards()
		ChanceCards.Add("Advance to Mayfair")
		ChanceCards.Add("Advance to Go")
		ChanceCards.Add("You are Assessed for Street Repairs £40 per House £115 per Hotel")
		ChanceCards.Add("Go to Jail. Move Directly to Jail. Do not pass 'Go' Do not Collect £200")
		ChanceCards.Add("Bank pays you Dividend of £50")
		ChanceCards.Add("Go back 3 Spaces")
		ChanceCards.Add("Pay School Fees of £150")
		ChanceCards.Add("Make General Repairs on all of Your Houses " & Chr(13) & "For each House pay £25" & Chr(13) & "For each Hotel pay £100")
		ChanceCards.Add("Speeding Fine £15")
		ChanceCards.Add("You have won a Crossword Competition Collect £100")
		ChanceCards.Add("Your Building and Loan Matures Collect £150")
		ChanceCards.Add("Get out of Jail Free")
		ChanceCards.Add("Avance to Trafalgar Square If you Pass 'Go' Collect £200")
		ChanceCards.Add("Take a Trip to Marylebone Station and if you Pass 'Go' Collect £200")
		ChanceCards.Add("Advance to Pall Mall If you Pass 'Go' Collect £200")
		ChanceCards.Add("'Drunk in Charge' Fine £20")
	End Sub

	Private Sub MonopolyMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		SetRelativePositions()
		PlayerOnePicture.Image = Nothing
		PlayerTwoPicture.Image = Nothing
		PlayerThreePicture.Image = Nothing
		PlayerFourPicture.Image = Nothing
		PlayerFivePicture.Image = Nothing
		PlayerSixPicture.Image = Nothing
		PlayerSevenPicture.Image = Nothing
		PlayerEightPicture.Image = Nothing
		For i = 0 To 7
			DefaultPieces(i) = PlayerOnePieceDrop.Items(i)
		Next
		DefaultPieces(8) = "None"
	End Sub

	Private Sub SetRelativePositions()
		Me.WindowState = FormWindowState.Maximized
		NoOfPlayersSlide.BackColor = Me.BackColor

		MainMenuPanel.Left = (Me.Width / 2) - (MainMenuPanel.Width / 2) - 20
		NewGamePanel.Left = (Me.Width / 2) - (NewGamePanel.Width / 2) - 20
		LoadingPanel.Left = (Me.Width / 2) - (LoadingPanel.Width / 2) - 20
		GamePanel.Parent = Me
		GamePanel.Size = Me.Size
		GamePanel.Left = 0
		GamePanel.Top = 0
		StatsPanel.Left = 0
		StatsPanel.Top = 0
	End Sub

	Private Sub MainQuit_Click(sender As Object, e As EventArgs) Handles MainQuit.Click
		Me.Close()
	End Sub

	Private Sub FormResized(sender As Object, e As EventArgs) Handles MyBase.Resize
		SetRelativePositions()
	End Sub

	Private Sub MainNewGame_Click(sender As Object, e As EventArgs) Handles MainNewGame.Click
		MainMenuPanel.Hide()
		NewGamePanel.Show()

	End Sub

	Private Sub FPMoney_CheckedChanged(sender As Object, e As EventArgs) Handles FPMoneyCheck.CheckedChanged
		If FPMoneyCheck.Checked = True Then
			FPTaxesCheck.Show()
		Else
			FPTaxesCheck.Hide()
		End If
	End Sub

	Private Sub NoOfPlayersSlide_Scroll(sender As Object, e As EventArgs) Handles NoOfPlayersSlide.Scroll
		PlayerNOlbl.Text = "Number of Players: " & NoOfPlayersSlide.Value
		Select Case NoOfPlayersSlide.Value
			Case 2
				PlayerThreePanel.Hide()
				PlayerFourPanel.Hide()
				PlayerFivePanel.Hide()
				PlayerSixPanel.Hide()
				PlayerSevenPanel.Hide()
				PlayerEightPanel.Hide()

				PlayerThreePicture.Image = Nothing
				If PlayerThreePieceDrop.Items.Contains("None") Then
					PlayerThreePieceDrop.SelectedItem = "None"
				Else
					PlayerThreePieceDrop.Items.Add("None")
					PlayerThreePieceDrop.SelectedItem = "None"
				End If
				PlayerFourPicture.Image = Nothing
				If PlayerFourPieceDrop.Items.Contains("None") Then
					PlayerFourPieceDrop.SelectedItem = "None"
				Else
					PlayerFourPieceDrop.Items.Add("None")
					PlayerFourPieceDrop.SelectedItem = "None"
				End If
				PlayerFivePicture.Image = Nothing
				If PlayerFivePieceDrop.Items.Contains("None") Then
					PlayerFivePieceDrop.SelectedItem = "None"
				Else
					PlayerFivePieceDrop.Items.Add("None")
					PlayerFivePieceDrop.SelectedItem = "None"
				End If
				PlayerSixPicture.Image = Nothing
				If PlayerSixPieceDrop.Items.Contains("None") Then
					PlayerSixPieceDrop.SelectedItem = "None"
				Else
					PlayerSixPieceDrop.Items.Add("None")
					PlayerSixPieceDrop.SelectedItem = "None"
				End If
				PlayerSevenPicture.Image = Nothing
				If PlayerSevenPieceDrop.Items.Contains("None") Then
					PlayerSevenPieceDrop.SelectedItem = "None"
				Else
					PlayerSevenPieceDrop.Items.Add("None")
					PlayerSevenPieceDrop.SelectedItem = "None"
				End If
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"
				PlayerSevenNametxt.Text = "Player7"
				PlayerSixNametxt.Text = "Player6"
				PlayerFiveNametxt.Text = "Player5"
				PlayerFourNametxt.Text = "Player4"
				PlayerThreeNametxt.Text = "Player3"
			Case 3
				PlayerThreePanel.Show()
				PlayerFourPanel.Hide()
				PlayerFivePanel.Hide()
				PlayerSixPanel.Hide()
				PlayerSevenPanel.Hide()
				PlayerEightPanel.Hide()
				PlayerFourPicture.Image = Nothing
				If PlayerFourPieceDrop.Items.Contains("None") Then
					PlayerFourPieceDrop.SelectedItem = "None"
				Else
					PlayerFourPieceDrop.Items.Add("None")
					PlayerFourPieceDrop.SelectedItem = "None"
				End If
				PlayerFivePicture.Image = Nothing
				If PlayerFivePieceDrop.Items.Contains("None") Then
					PlayerFivePieceDrop.SelectedItem = "None"
				Else
					PlayerFivePieceDrop.Items.Add("None")
					PlayerFivePieceDrop.SelectedItem = "None"
				End If
				PlayerSixPicture.Image = Nothing
				If PlayerSixPieceDrop.Items.Contains("None") Then
					PlayerSixPieceDrop.SelectedItem = "None"
				Else
					PlayerSixPieceDrop.Items.Add("None")
					PlayerSixPieceDrop.SelectedItem = "None"
				End If
				PlayerSevenPicture.Image = Nothing
				If PlayerSevenPieceDrop.Items.Contains("None") Then
					PlayerSevenPieceDrop.SelectedItem = "None"
				Else
					PlayerSevenPieceDrop.Items.Add("None")
					PlayerSevenPieceDrop.SelectedItem = "None"
				End If
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"
				PlayerSevenNametxt.Text = "Player7"
				PlayerSixNametxt.Text = "Player6"
				PlayerFiveNametxt.Text = "Player5"
				PlayerFourNametxt.Text = "Player4"
			Case 4
				PlayerThreePanel.Show()
				PlayerFourPanel.Show()
				PlayerFivePanel.Hide()
				PlayerSixPanel.Hide()
				PlayerSevenPanel.Hide()
				PlayerEightPanel.Hide()

				PlayerFivePicture.Image = Nothing
				If PlayerFivePieceDrop.Items.Contains("None") Then
					PlayerFivePieceDrop.SelectedItem = "None"
				Else
					PlayerFivePieceDrop.Items.Add("None")
					PlayerFivePieceDrop.SelectedItem = "None"
				End If
				PlayerSixPicture.Image = Nothing
				If PlayerSixPieceDrop.Items.Contains("None") Then
					PlayerSixPieceDrop.SelectedItem = "None"
				Else
					PlayerSixPieceDrop.Items.Add("None")
					PlayerSixPieceDrop.SelectedItem = "None"
				End If
				PlayerSevenPicture.Image = Nothing
				If PlayerSevenPieceDrop.Items.Contains("None") Then
					PlayerSevenPieceDrop.SelectedItem = "None"
				Else
					PlayerSevenPieceDrop.Items.Add("None")
					PlayerSevenPieceDrop.SelectedItem = "None"
				End If
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"
				PlayerSevenNametxt.Text = "Player7"
				PlayerSixNametxt.Text = "Player6"
				PlayerFiveNametxt.Text = "Player5"

			Case 5
				PlayerThreePanel.Show()
				PlayerFourPanel.Show()
				PlayerFivePanel.Show()
				PlayerSixPanel.Hide()
				PlayerSevenPanel.Hide()
				PlayerEightPanel.Hide()

				PlayerSixPicture.Image = Nothing
				If PlayerSixPieceDrop.Items.Contains("None") Then
					PlayerSixPieceDrop.SelectedItem = "None"
				Else
					PlayerSixPieceDrop.Items.Add("None")
					PlayerSixPieceDrop.SelectedItem = "None"
				End If
				PlayerSevenPicture.Image = Nothing
				If PlayerSevenPieceDrop.Items.Contains("None") Then
					PlayerSevenPieceDrop.SelectedItem = "None"
				Else
					PlayerSevenPieceDrop.Items.Add("None")
					PlayerSevenPieceDrop.SelectedItem = "None"
				End If
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"
				PlayerSevenNametxt.Text = "Player7"
				PlayerSixNametxt.Text = "Player6"

			Case 6
				PlayerThreePanel.Show()
				PlayerFourPanel.Show()
				PlayerFivePanel.Show()
				PlayerSixPanel.Show()
				PlayerSevenPanel.Hide()
				PlayerEightPanel.Hide()

				PlayerSevenPicture.Image = Nothing
				If PlayerSevenPieceDrop.Items.Contains("None") Then
					PlayerSevenPieceDrop.SelectedItem = "None"
				Else
					PlayerSevenPieceDrop.Items.Add("None")
					PlayerSevenPieceDrop.SelectedItem = "None"
				End If
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"
				PlayerSevenNametxt.Text = "Player7"
			Case 7
				PlayerThreePanel.Show()
				PlayerFourPanel.Show()
				PlayerFivePanel.Show()
				PlayerSixPanel.Show()
				PlayerSevenPanel.Show()
				PlayerEightPanel.Hide()
				PlayerEightPicture.Image = Nothing
				If PlayerEightPieceDrop.Items.Contains("None") Then
					PlayerEightPieceDrop.SelectedItem = "None"
				Else
					PlayerEightPieceDrop.Items.Add("None")
					PlayerEightPieceDrop.SelectedItem = "None"
				End If
				PlayerEightNametxt.Text = "Player8"

			Case 8
				PlayerThreePanel.Show()
				PlayerFourPanel.Show()
				PlayerFivePanel.Show()
				PlayerSixPanel.Show()
				PlayerSevenPanel.Show()
				PlayerEightPanel.Show()
		End Select
	End Sub

	Private Sub PlayerOnePieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerOnePieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerOnePieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerOnePicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerOnePicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerOnePicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerOnePicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerOnePicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerOnePicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerOnePicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerOnePicture.Image = My.Resources.Iron
			Case "None"
				PlayerOnePicture.Image = Nothing
		End Select
		selectedPiecesSetup(0, 0) = PlayerOnePieceDrop.SelectedItem
		CheckDupePieces(1)
	End Sub

	Private Sub PlayerTwoPieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerTwoPieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerTwoPieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerTwoPicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerTwoPicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerTwoPicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerTwoPicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerTwoPicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerTwoPicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerTwoPicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerTwoPicture.Image = My.Resources.Iron
			Case "None"
				PlayerTwoPicture.Image = Nothing
		End Select
		selectedPiecesSetup(1, 0) = PlayerTwoPieceDrop.SelectedItem
		CheckDupePieces(2)
	End Sub

	Private Sub PlayerThreePieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerThreePieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerThreePieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerThreePicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerThreePicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerThreePicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerThreePicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerThreePicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerThreePicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerThreePicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerThreePicture.Image = My.Resources.Iron
			Case "None"
				PlayerThreePicture.Image = Nothing
		End Select
		selectedPiecesSetup(2, 0) = PlayerThreePieceDrop.SelectedItem
		CheckDupePieces(3)
	End Sub

	Private Sub PlayerFourPieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerFourPieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerFourPieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerFourPicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerFourPicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerFourPicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerFourPicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerFourPicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerFourPicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerFourPicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerFourPicture.Image = My.Resources.Iron
			Case "None"
				PlayerFourPicture.Image = Nothing
		End Select
		selectedPiecesSetup(3, 0) = PlayerFourPieceDrop.SelectedItem
		CheckDupePieces(4)
	End Sub

	Private Sub PlayerFivePieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerFivePieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerFivePieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerFivePicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerFivePicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerFivePicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerFivePicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerFivePicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerFivePicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerFivePicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerFivePicture.Image = My.Resources.Iron
			Case "None"
				PlayerFivePicture.Image = Nothing

		End Select
		selectedPiecesSetup(4, 0) = PlayerFivePieceDrop.SelectedItem
		CheckDupePieces(5)
	End Sub

	Private Sub PlayerSixPieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerSixPieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerSixPieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerSixPicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerSixPicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerSixPicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerSixPicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerSixPicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerSixPicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerSixPicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerSixPicture.Image = My.Resources.Iron
			Case "None"
				PlayerSixPicture.Image = Nothing
		End Select
		selectedPiecesSetup(5, 0) = PlayerSixPieceDrop.SelectedItem
		CheckDupePieces(6)
	End Sub

	Private Sub PlayerSevenPieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerSevenPieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerSevenPieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerSevenPicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerSevenPicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerSevenPicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerSevenPicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerSevenPicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerSevenPicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerSevenPicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerSevenPicture.Image = My.Resources.Iron
			Case "None"
				PlayerSevenPicture.Image = Nothing
		End Select
		selectedPiecesSetup(6, 0) = PlayerSevenPieceDrop.SelectedItem
		CheckDupePieces(7)
	End Sub

	Private Sub PlayerEightPieceDrop_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PlayerEightPieceDrop.SelectedIndexChanged
		If UpdatingPlayersSetup = True Then
			Exit Sub
		End If
		Select Case PlayerEightPieceDrop.SelectedItem
			Case "Wheelbarrow"
				PlayerEightPicture.Image = My.Resources.Wheelbarrow
			Case "Battleship"
				PlayerEightPicture.Image = My.Resources.Ship
			Case "Racecar"
				PlayerEightPicture.Image = My.Resources.Car
			Case "Thimble"
				PlayerEightPicture.Image = My.Resources.Thimble
			Case "Boot"
				PlayerEightPicture.Image = My.Resources.Boot
			Case "Dog"
				PlayerEightPicture.Image = My.Resources.Dog
			Case "Top Hat"
				PlayerEightPicture.Image = My.Resources.Hat
			Case "Iron"
				PlayerEightPicture.Image = My.Resources.Iron
			Case "None"
				PlayerEightPicture.Image = Nothing
		End Select
		selectedPiecesSetup(7, 0) = PlayerEightPieceDrop.SelectedItem
		CheckDupePieces(8)

	End Sub

	Private Sub CheckDupePieces(ByVal current)
		UpdatingPlayersSetup = True

		PlayerOnePieceDrop.Items.Clear()
		PlayerTwoPieceDrop.Items.Clear()
		PlayerThreePieceDrop.Items.Clear()
		PlayerFourPieceDrop.Items.Clear()
		PlayerFivePieceDrop.Items.Clear()
		PlayerSixPieceDrop.Items.Clear()
		PlayerSevenPieceDrop.Items.Clear()
		PlayerEightPieceDrop.Items.Clear()

		PlayerOnePieceDrop.Items.AddRange(DefaultPieces)
		PlayerTwoPieceDrop.Items.AddRange(DefaultPieces)
		PlayerThreePieceDrop.Items.AddRange(DefaultPieces)
		PlayerFourPieceDrop.Items.AddRange(DefaultPieces)
		PlayerFivePieceDrop.Items.AddRange(DefaultPieces)
		PlayerSixPieceDrop.Items.AddRange(DefaultPieces)
		PlayerSevenPieceDrop.Items.AddRange(DefaultPieces)
		PlayerEightPieceDrop.Items.AddRange(DefaultPieces)

		For i = 0 To 7
			PlayerOnePieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerTwoPieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerThreePieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerFourPieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerFivePieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerSixPieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerSevenPieceDrop.Items.Remove(selectedPiecesSetup(i, 0))
			PlayerEightPieceDrop.Items.Remove(selectedPiecesSetup(i, 0))

		Next

		Select Case current
			Case 1
				PlayerOnePieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerOnePieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 2
				PlayerTwoPieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerTwoPieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 3
				PlayerThreePieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerThreePieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 4
				PlayerFourPieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerFourPieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 5
				PlayerFivePieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerFivePieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 6
				PlayerSixPieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerSixPieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 7
				PlayerSevenPieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerSevenPieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
			Case 8
				PlayerEightPieceDrop.Items.Add(selectedPiecesSetup(current - 1, 0))
				PlayerEightPieceDrop.SelectedItem = selectedPiecesSetup(current - 1, 0)
		End Select

		UpdatingPlayersSetup = False

	End Sub

	Private Sub InitialiseComChestCards()
		ComChestCards.Add("Income Tax refund Collect £20")
		ComChestCards.Add("From Sale of Stock you get £50")
		ComChestCards.Add("It is Your Birthday Collect $10 from each Player")
		ComChestCards.Add("Receive Interest on Shares £25")
		ComChestCards.Add("Get out of Jail Free")
		ComChestCards.Add("Advance to 'Go'")
		ComChestCards.Add("Pay Hospital £100")
		ComChestCards.Add("You have Won Second Prize in a Beauty Contest Collect £10")
		ComChestCards.Add("Bank Error in your Favour Collect £200")
		ComChestCards.Add("You Inherit £100")
		ComChestCards.Add("Go to Jail. Move Directly to Jail. Do not Pass 'Go'. Do not Collect £200")
		ComChestCards.Add("Pay your Insurance Premium £50")
		ComChestCards.Add("Doctor's Fee Pay £50")
		ComChestCards.Add("Go Back to Old Kent Road")
		ComChestCards.Add("Annuity Matures Collect £100")
	End Sub

	Private Sub ShuffleComChestCards()
		ComChestCards.Sort(New Randomizer(Of String)())
	End Sub

	Private Sub StartNewGame_Click(sender As Object, e As EventArgs) Handles StartNewGame.Click
		If SetHouseRules() = True Then
			PlayerNumber = NoOfPlayersSlide.Value
			If SetUpPlayers() = True Then
				Loadinglbl.Text = "LOADING..."
				LoadingPanel.Parent = Me
				LoadStartGame.Parent = LoadingPanel
				GamePanel.Parent = Me
				LoadingPanel.Show()
				NewGamePanel.Hide()
				InitialiseBoardTiles()
				UpdatePlayerStats()
				LoadGameBoard(True)
				InitialiseChanceCards()
				InitialiseComChestCards()
				ShuffleChanceCards()
				ShuffleComChestCards()
				LoadingBar.Value = 100
				Loadinglbl.Text = "LOADED!"
				LoadStartGame.Show()
				CounterForTurnsInGame = 0
				MoneyOnFreeParking = 100
			Else
				MsgBox("Invalid, Something is wrong with the players", MsgBoxStyle.Critical, "Invalid Player Setup")
			End If
		Else
			MsgBox("Invalid, This is not a number.", MsgBoxStyle.Critical, "Invalid House Rule")
		End If

	End Sub
	Private Sub InitialiseBoardTiles()
		Dim LoadingTile As Integer = 0

		'Everything from here is fixed, the starting state of a new game board
		'Loading tile is the current tile, Starting at 0 being GO working clockwise round to Mayfair as 39
		LoadingBar.Value = 0
		'GO Tile
		Tiles(LoadingTile).GO = True
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "GO"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = 0
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"
		'End of GO Tile

		LoadingTile += 1
		LoadingBar.Value += 1

		'Old Kent Road
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "OLD KENT ROAD"
		Tiles(LoadingTile).Cost = 60
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.SaddleBrown

		'Rent
		Tiles(LoadingTile).Rent.Base = 2
		Tiles(LoadingTile).Rent.OneHouse = 10
		Tiles(LoadingTile).Rent.TwoHouse = 30
		Tiles(LoadingTile).Rent.ThreeHouse = 90
		Tiles(LoadingTile).Rent.FourHouse = 160
		Tiles(LoadingTile).Rent.Hotel = 250
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 30
		Tiles(LoadingTile).HotelCost = 30
		Tiles(LoadingTile).Owner = ""
		'End of Old Kent Road

		LoadingTile += 1
		LoadingBar.Value += 1

		'First community chest
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = True
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "COMMUNITY CHEST"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.DeepSkyBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"
		'End of First community chest

		LoadingTile += 1
		LoadingBar.Value += 1

		'Whitechapel Road
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "WHITECHAPEL ROAD"
		Tiles(LoadingTile).Cost = 60
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.SaddleBrown


		'Rent
		Tiles(LoadingTile).Rent.Base = 4
		Tiles(LoadingTile).Rent.OneHouse = 20
		Tiles(LoadingTile).Rent.TwoHouse = 60
		Tiles(LoadingTile).Rent.ThreeHouse = 180
		Tiles(LoadingTile).Rent.FourHouse = 360
		Tiles(LoadingTile).Rent.Hotel = 450
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 30
		Tiles(LoadingTile).HotelCost = 30
		Tiles(LoadingTile).Owner = ""

		LoadingTile += 1
		LoadingBar.Value += 1
		'End of whitechapel road

		'INCOME TAX

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = True
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "INCOME TAX"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		Tiles(LoadingTile).Rent.TaxAmount = 200
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF INCOME TAX

		'KINGS CROSS STATION

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = True
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "KINGS CROSS STATION"
		Tiles(LoadingTile).Cost = 200
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 25
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 25
		Tiles(LoadingTile).Rent.TwoStation = 50
		Tiles(LoadingTile).Rent.ThreeStation = 100
		Tiles(LoadingTile).Rent.FourStation = 200
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF KINGS CROSS

		'THE ANGEL ISLINGTON

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "THE ANGEL ISLINGTON"
		Tiles(LoadingTile).Cost = 100
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.LightBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 6
		Tiles(LoadingTile).Rent.OneHouse = 30
		Tiles(LoadingTile).Rent.TwoHouse = 90
		Tiles(LoadingTile).Rent.ThreeHouse = 270
		Tiles(LoadingTile).Rent.FourHouse = 400
		Tiles(LoadingTile).Rent.Hotel = 550
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 50
		Tiles(LoadingTile).HotelCost = 50
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF THE ANGEL ISLINGTON

		'FIRST CHANCE
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = True
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "CHANCE"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.OrangeRed

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF FIRST CHANCE

		'EUSTON ROAD

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "EUSTON ROAD"
		Tiles(LoadingTile).Cost = 100
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.LightBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 6
		Tiles(LoadingTile).Rent.OneHouse = 30
		Tiles(LoadingTile).Rent.TwoHouse = 90
		Tiles(LoadingTile).Rent.ThreeHouse = 270
		Tiles(LoadingTile).Rent.FourHouse = 400
		Tiles(LoadingTile).Rent.Hotel = 550
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 50
		Tiles(LoadingTile).HotelCost = 50
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF EUSTON ROAD

		'PENTONVILLE ROAD

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "PENTONVILLE ROAD"
		Tiles(LoadingTile).Cost = 120
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.LightBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 8
		Tiles(LoadingTile).Rent.OneHouse = 40
		Tiles(LoadingTile).Rent.TwoHouse = 100
		Tiles(LoadingTile).Rent.ThreeHouse = 300
		Tiles(LoadingTile).Rent.FourHouse = 450
		Tiles(LoadingTile).Rent.Hotel = 600
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 50
		Tiles(LoadingTile).HotelCost = 50
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF PENTONVILL ROAD

		'JAIL

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = True
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "JAIL"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF JAIL

		'PALL MALL

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "PALL MALL"
		Tiles(LoadingTile).Cost = 140
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.DeepPink

		'Rent
		Tiles(LoadingTile).Rent.Base = 10
		Tiles(LoadingTile).Rent.OneHouse = 50
		Tiles(LoadingTile).Rent.TwoHouse = 150
		Tiles(LoadingTile).Rent.ThreeHouse = 450
		Tiles(LoadingTile).Rent.FourHouse = 625
		Tiles(LoadingTile).Rent.Hotel = 750
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF PALL MALL

		'ELECTRIC COMPANY

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = True
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "ELECTRIC COMPANY"
		Tiles(LoadingTile).Cost = 150
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 4
		Tiles(LoadingTile).Rent.TwoUtility = 10
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF ELECTRIC COMPANY

		'WHITEHALL

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "WHITEHALL"
		Tiles(LoadingTile).Cost = 140
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.DeepPink

		'Rent
		Tiles(LoadingTile).Rent.Base = 10
		Tiles(LoadingTile).Rent.OneHouse = 50
		Tiles(LoadingTile).Rent.TwoHouse = 150
		Tiles(LoadingTile).Rent.ThreeHouse = 450
		Tiles(LoadingTile).Rent.FourHouse = 625
		Tiles(LoadingTile).Rent.Hotel = 750
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF WHITEHALL

		'NORTHUMBERLAND AVENUE

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "NORTHUMBERLAND AVENUE"
		Tiles(LoadingTile).Cost = 160
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.DeepPink

		'Rent
		Tiles(LoadingTile).Rent.Base = 12
		Tiles(LoadingTile).Rent.OneHouse = 60
		Tiles(LoadingTile).Rent.TwoHouse = 180
		Tiles(LoadingTile).Rent.ThreeHouse = 500
		Tiles(LoadingTile).Rent.FourHouse = 700
		Tiles(LoadingTile).Rent.Hotel = 900
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF NORTHUMBERLAND AVENUE

		'MARYLEBONE STATION

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = True
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "MARYLEBONE STATION"
		Tiles(LoadingTile).Cost = 200
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 25
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 25
		Tiles(LoadingTile).Rent.TwoStation = 50
		Tiles(LoadingTile).Rent.ThreeStation = 100
		Tiles(LoadingTile).Rent.FourStation = 200
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF MARYLEBONE STATION

		'BOW STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "BOW STREET"
		Tiles(LoadingTile).Cost = 180
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Orange

		'Rent
		Tiles(LoadingTile).Rent.Base = 14
		Tiles(LoadingTile).Rent.OneHouse = 70
		Tiles(LoadingTile).Rent.TwoHouse = 200
		Tiles(LoadingTile).Rent.ThreeHouse = 550
		Tiles(LoadingTile).Rent.FourHouse = 750
		Tiles(LoadingTile).Rent.Hotel = 950
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF BOW STREET

		'SECOND COMMUNITY CHEST

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = True
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "COMMUNITY CHEST"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF SECOND COMMUNITY CHEST

		'MARLBOROUGH STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "MARLBOROUGH STREET"
		Tiles(LoadingTile).Cost = 180
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Orange

		'Rent
		Tiles(LoadingTile).Rent.Base = 14
		Tiles(LoadingTile).Rent.OneHouse = 70
		Tiles(LoadingTile).Rent.TwoHouse = 200
		Tiles(LoadingTile).Rent.ThreeHouse = 550
		Tiles(LoadingTile).Rent.FourHouse = 750
		Tiles(LoadingTile).Rent.Hotel = 950
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF MARLBOROUGH STREET

		'VINE STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "VINE STREET"
		Tiles(LoadingTile).Cost = 200
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Orange

		'Rent
		Tiles(LoadingTile).Rent.Base = 16
		Tiles(LoadingTile).Rent.OneHouse = 80
		Tiles(LoadingTile).Rent.TwoHouse = 220
		Tiles(LoadingTile).Rent.ThreeHouse = 600
		Tiles(LoadingTile).Rent.FourHouse = 800
		Tiles(LoadingTile).Rent.Hotel = 1000
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 100
		Tiles(LoadingTile).HotelCost = 100
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF VINE STREET

		'FREE PARKING

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = True
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "FREE PARKING"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF FREE PARKING

		'STRAND

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "STRAND"
		Tiles(LoadingTile).Cost = 220
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Red

		'Rent
		Tiles(LoadingTile).Rent.Base = 18
		Tiles(LoadingTile).Rent.OneHouse = 90
		Tiles(LoadingTile).Rent.TwoHouse = 250
		Tiles(LoadingTile).Rent.ThreeHouse = 700
		Tiles(LoadingTile).Rent.FourHouse = 875
		Tiles(LoadingTile).Rent.Hotel = 1050
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF STRAND

		'Second Chance
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = True
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "CHANCE"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.OrangeRed

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'End of Second chance

		'FLEET STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "FLEET STREET"
		Tiles(LoadingTile).Cost = 220
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Red

		'Rent
		Tiles(LoadingTile).Rent.Base = 18
		Tiles(LoadingTile).Rent.OneHouse = 90
		Tiles(LoadingTile).Rent.TwoHouse = 250
		Tiles(LoadingTile).Rent.ThreeHouse = 700
		Tiles(LoadingTile).Rent.FourHouse = 875
		Tiles(LoadingTile).Rent.Hotel = 1050
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF FLEET STREET

		'TRAFALGAR SQUARE

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "TRAFALGAR SQUARE"
		Tiles(LoadingTile).Cost = 240
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Red

		'Rent
		Tiles(LoadingTile).Rent.Base = 20
		Tiles(LoadingTile).Rent.OneHouse = 100
		Tiles(LoadingTile).Rent.TwoHouse = 300
		Tiles(LoadingTile).Rent.ThreeHouse = 750
		Tiles(LoadingTile).Rent.FourHouse = 925
		Tiles(LoadingTile).Rent.Hotel = 1100
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF TRAFALGAR SQUARE

		'FENCHURCH ST STATION

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = True
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "FENCHURCH ST STATION"
		Tiles(LoadingTile).Cost = 200
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 25
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 25
		Tiles(LoadingTile).Rent.TwoStation = 50
		Tiles(LoadingTile).Rent.ThreeStation = 100
		Tiles(LoadingTile).Rent.FourStation = 200
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF FENCHURCH ST STATION

		'LEICESTER SQUARE

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "LEICESTER SQUARE"
		Tiles(LoadingTile).Cost = 260
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Yellow

		'Rent
		Tiles(LoadingTile).Rent.Base = 22
		Tiles(LoadingTile).Rent.OneHouse = 110
		Tiles(LoadingTile).Rent.TwoHouse = 330
		Tiles(LoadingTile).Rent.ThreeHouse = 800
		Tiles(LoadingTile).Rent.FourHouse = 975
		Tiles(LoadingTile).Rent.Hotel = 1150
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF LEICESTER SQUARE

		'COVENTRY STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "COVENTRY STREET"
		Tiles(LoadingTile).Cost = 260
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Yellow

		'Rent
		Tiles(LoadingTile).Rent.Base = 22
		Tiles(LoadingTile).Rent.OneHouse = 110
		Tiles(LoadingTile).Rent.TwoHouse = 330
		Tiles(LoadingTile).Rent.ThreeHouse = 800
		Tiles(LoadingTile).Rent.FourHouse = 975
		Tiles(LoadingTile).Rent.Hotel = 1150
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF COVENTRY STREET

		'WATER WORKS

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = True
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "WATER WORKS"
		Tiles(LoadingTile).Cost = 150
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 4
		Tiles(LoadingTile).Rent.TwoUtility = 10
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF WATER WORKS

		'PICCADILLY

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "PICCADILLY"
		Tiles(LoadingTile).Cost = 280
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Yellow

		'Rent
		Tiles(LoadingTile).Rent.Base = 22
		Tiles(LoadingTile).Rent.OneHouse = 120
		Tiles(LoadingTile).Rent.TwoHouse = 360
		Tiles(LoadingTile).Rent.ThreeHouse = 850
		Tiles(LoadingTile).Rent.FourHouse = 1025
		Tiles(LoadingTile).Rent.Hotel = 1200
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 150
		Tiles(LoadingTile).HotelCost = 150
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF PICCADILLY

		'GO TO JAIL TILE

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = True
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "GO TO JAIL"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF GO TO JAIL TILE

		'REGENT STREET

		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "REGENT STREET"
		Tiles(LoadingTile).Cost = 300
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Green

		'Rent
		Tiles(LoadingTile).Rent.Base = 26
		Tiles(LoadingTile).Rent.OneHouse = 130
		Tiles(LoadingTile).Rent.TwoHouse = 390
		Tiles(LoadingTile).Rent.ThreeHouse = 900
		Tiles(LoadingTile).Rent.FourHouse = 1100
		Tiles(LoadingTile).Rent.Hotel = 1275
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 200
		Tiles(LoadingTile).HotelCost = 200
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF REGENT STREET

		'OXFORD STREET
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "OXFORD STREET"
		Tiles(LoadingTile).Cost = 300
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Green

		'Rent
		Tiles(LoadingTile).Rent.Base = 26
		Tiles(LoadingTile).Rent.OneHouse = 130
		Tiles(LoadingTile).Rent.TwoHouse = 390
		Tiles(LoadingTile).Rent.ThreeHouse = 900
		Tiles(LoadingTile).Rent.FourHouse = 1100
		Tiles(LoadingTile).Rent.Hotel = 1275
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 200
		Tiles(LoadingTile).HotelCost = 200
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF OXFORD STREET

		'THIRD COMMUNITY CHEST
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = True
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "COMMUNITY CHEST"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF THIRD COMMUNITY CHEST

		'BOND STREET
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "BOND STREET"
		Tiles(LoadingTile).Cost = 320
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Green

		'Rent
		Tiles(LoadingTile).Rent.Base = 28
		Tiles(LoadingTile).Rent.OneHouse = 150
		Tiles(LoadingTile).Rent.TwoHouse = 450
		Tiles(LoadingTile).Rent.ThreeHouse = 1000
		Tiles(LoadingTile).Rent.FourHouse = 1200
		Tiles(LoadingTile).Rent.Hotel = 1400
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 200
		Tiles(LoadingTile).HotelCost = 200
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF BOND STREET

		'LIVERPOOL STREET STATION
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = True
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "LIVERPOOL STREET STATION"
		Tiles(LoadingTile).Cost = 200
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 25
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 25
		Tiles(LoadingTile).Rent.TwoStation = 50
		Tiles(LoadingTile).Rent.ThreeStation = 100
		Tiles(LoadingTile).Rent.FourStation = 200
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF LIVERPOOL STREET STATION

		'THIRD CHANCE
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = True
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "CHANCE"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF THIRD CHANCE

		'PARK LANE
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "PARK LANE"
		Tiles(LoadingTile).Cost = 350
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.MediumBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 35
		Tiles(LoadingTile).Rent.OneHouse = 175
		Tiles(LoadingTile).Rent.TwoHouse = 500
		Tiles(LoadingTile).Rent.ThreeHouse = 1100
		Tiles(LoadingTile).Rent.FourHouse = 1300
		Tiles(LoadingTile).Rent.Hotel = 1500
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 200
		Tiles(LoadingTile).HotelCost = 200
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF PARK LANE

		'SUPER TAX
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = True
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = False

		Tiles(LoadingTile).Title = "SUPER TAX"
		Tiles(LoadingTile).Cost = 0
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.Transparent

		'Rent
		Tiles(LoadingTile).Rent.Base = 0
		Tiles(LoadingTile).Rent.OneHouse = 0
		Tiles(LoadingTile).Rent.TwoHouse = 0
		Tiles(LoadingTile).Rent.ThreeHouse = 0
		Tiles(LoadingTile).Rent.FourHouse = 0
		Tiles(LoadingTile).Rent.Hotel = 0
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		Tiles(LoadingTile).Rent.TaxAmount = 100
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 0
		Tiles(LoadingTile).HotelCost = 0
		Tiles(LoadingTile).Owner = "N/A"


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF SUPER TAX

		'MAYFAIR
		Tiles(LoadingTile).GO = False
		Tiles(LoadingTile).Jail = False
		Tiles(LoadingTile).FreeParking = False
		Tiles(LoadingTile).Tax = False
		Tiles(LoadingTile).Station = False
		Tiles(LoadingTile).Utility = False
		Tiles(LoadingTile).GoToJail = False
		Tiles(LoadingTile).Chance = False
		Tiles(LoadingTile).ComChest = False
		Tiles(LoadingTile).IsProperty = True

		Tiles(LoadingTile).Title = "MAYFAIR"
		Tiles(LoadingTile).Cost = 400
		Tiles(LoadingTile).MortgageValue = (Tiles(LoadingTile).Cost / 2)
		Tiles(LoadingTile).Colour = Color.MediumBlue

		'Rent
		Tiles(LoadingTile).Rent.Base = 50
		Tiles(LoadingTile).Rent.OneHouse = 200
		Tiles(LoadingTile).Rent.TwoHouse = 600
		Tiles(LoadingTile).Rent.ThreeHouse = 1400
		Tiles(LoadingTile).Rent.FourHouse = 1700
		Tiles(LoadingTile).Rent.Hotel = 2000
		Tiles(LoadingTile).Rent.OneStation = 0
		Tiles(LoadingTile).Rent.TwoStation = 0
		Tiles(LoadingTile).Rent.ThreeStation = 0
		Tiles(LoadingTile).Rent.FourStation = 0
		Tiles(LoadingTile).Rent.OneUtility = 0
		Tiles(LoadingTile).Rent.TwoUtility = 0
		'End of Rent

		Tiles(LoadingTile).HouseNo = 0
		Tiles(LoadingTile).Hotel = False
		Tiles(LoadingTile).HouseCost = 200
		Tiles(LoadingTile).HotelCost = 200
		Tiles(LoadingTile).Owner = ""


		LoadingTile += 1
		LoadingBar.Value += 1
		'END OF MAYFAIR

	End Sub
	Private Function SetUpPlayers() As Boolean
		ReDim CurrentPlayers(PlayerNumber - 1)
		Try
			Select Case PlayerNumber
				Case 2
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text

				Case 3
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text

				Case 4
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text
					'Player 4
					CurrentPlayers(3).Name = PlayerFourNametxt.Text
					CurrentPlayers(3).GetOutJailCards = 0
					If PlayerFourStateDrop.Text = "Human" Then
						CurrentPlayers(3).Human = True
					Else
						CurrentPlayers(3).Human = False
					End If
					CurrentPlayers(3).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(3).Piece = PlayerFourPieceDrop.Text

				Case 5
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text
					'Player 4
					CurrentPlayers(3).Name = PlayerFourNametxt.Text
					CurrentPlayers(3).GetOutJailCards = 0
					If PlayerFourStateDrop.Text = "Human" Then
						CurrentPlayers(3).Human = True
					Else
						CurrentPlayers(3).Human = False
					End If
					CurrentPlayers(3).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(3).Piece = PlayerFourPieceDrop.Text
					'Player 5
					CurrentPlayers(4).Name = PlayerFiveNametxt.Text
					CurrentPlayers(4).GetOutJailCards = 0
					If PlayerFiveStateDrop.Text = "Human" Then
						CurrentPlayers(4).Human = True
					Else
						CurrentPlayers(4).Human = False
					End If
					CurrentPlayers(4).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(4).Piece = PlayerFivePieceDrop.Text

				Case 6
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text
					'Player 4
					CurrentPlayers(3).Name = PlayerFourNametxt.Text
					CurrentPlayers(3).GetOutJailCards = 0
					If PlayerFourStateDrop.Text = "Human" Then
						CurrentPlayers(3).Human = True
					Else
						CurrentPlayers(3).Human = False
					End If
					CurrentPlayers(3).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(3).Piece = PlayerFourPieceDrop.Text
					'Player 5
					CurrentPlayers(4).Name = PlayerFiveNametxt.Text
					CurrentPlayers(4).GetOutJailCards = 0
					If PlayerFiveStateDrop.Text = "Human" Then
						CurrentPlayers(4).Human = True
					Else
						CurrentPlayers(4).Human = False
					End If
					CurrentPlayers(4).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(4).Piece = PlayerFivePieceDrop.Text
					'Player 6
					CurrentPlayers(5).Name = PlayerSixNametxt.Text
					CurrentPlayers(5).GetOutJailCards = 0
					If PlayerSixStateDrop.Text = "Human" Then
						CurrentPlayers(5).Human = True
					Else
						CurrentPlayers(5).Human = False
					End If
					CurrentPlayers(5).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(5).Piece = PlayerSixPieceDrop.Text
				Case 7
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text
					'Player 4
					CurrentPlayers(3).Name = PlayerFourNametxt.Text
					CurrentPlayers(3).GetOutJailCards = 0
					If PlayerFourStateDrop.Text = "Human" Then
						CurrentPlayers(3).Human = True
					Else
						CurrentPlayers(3).Human = False
					End If
					CurrentPlayers(3).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(3).Piece = PlayerFourPieceDrop.Text
					'Player 5
					CurrentPlayers(4).Name = PlayerFiveNametxt.Text
					CurrentPlayers(4).GetOutJailCards = 0
					If PlayerFiveStateDrop.Text = "Human" Then
						CurrentPlayers(4).Human = True
					Else
						CurrentPlayers(4).Human = False
					End If
					CurrentPlayers(4).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(4).Piece = PlayerFivePieceDrop.Text
					'Player 6
					CurrentPlayers(5).Name = PlayerSixNametxt.Text
					CurrentPlayers(5).GetOutJailCards = 0
					If PlayerSixStateDrop.Text = "Human" Then
						CurrentPlayers(5).Human = True
					Else
						CurrentPlayers(5).Human = False
					End If
					CurrentPlayers(5).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(5).Piece = PlayerSixPieceDrop.Text
					'Player 7
					CurrentPlayers(6).Name = PlayerSevenNametxt.Text
					CurrentPlayers(6).GetOutJailCards = 0
					If PlayerSevenStateDrop.Text = "Human" Then
						CurrentPlayers(6).Human = True
					Else
						CurrentPlayers(6).Human = False
					End If
					CurrentPlayers(6).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(6).Piece = PlayerSevenPieceDrop.Text
				Case 8
					'Set up Player 1
					CurrentPlayers(0).Name = PlayerOneNametxt.Text
					CurrentPlayers(0).GetOutJailCards = 0
					If PlayerOneStateDrop.Text = "Human" Then
						CurrentPlayers(0).Human = True
					Else
						CurrentPlayers(0).Human = False
					End If
					CurrentPlayers(0).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(0).Piece = PlayerOnePieceDrop.Text
					'Player 2
					CurrentPlayers(1).Name = PlayerTwoNametxt.Text
					CurrentPlayers(1).GetOutJailCards = 0
					If PlayerTwoStateDrop.Text = "Human" Then
						CurrentPlayers(1).Human = True
					Else
						CurrentPlayers(1).Human = False
					End If
					CurrentPlayers(1).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(1).Piece = PlayerTwoPieceDrop.Text
					'Player 3
					CurrentPlayers(2).Name = PlayerThreeNametxt.Text
					CurrentPlayers(2).GetOutJailCards = 0
					If PlayerThreeStateDrop.Text = "Human" Then
						CurrentPlayers(2).Human = True
					Else
						CurrentPlayers(2).Human = False
					End If
					CurrentPlayers(2).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(2).Piece = PlayerThreePieceDrop.Text
					'Player 4
					CurrentPlayers(3).Name = PlayerFourNametxt.Text
					CurrentPlayers(3).GetOutJailCards = 0
					If PlayerFourStateDrop.Text = "Human" Then
						CurrentPlayers(3).Human = True
					Else
						CurrentPlayers(3).Human = False
					End If
					CurrentPlayers(3).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(3).Piece = PlayerFourPieceDrop.Text
					'Player 5
					CurrentPlayers(4).Name = PlayerFiveNametxt.Text
					CurrentPlayers(4).GetOutJailCards = 0
					If PlayerFiveStateDrop.Text = "Human" Then
						CurrentPlayers(4).Human = True
					Else
						CurrentPlayers(4).Human = False
					End If
					CurrentPlayers(4).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(4).Piece = PlayerFivePieceDrop.Text
					'Player 6
					CurrentPlayers(5).Name = PlayerSixNametxt.Text
					CurrentPlayers(5).GetOutJailCards = 0
					If PlayerSixStateDrop.Text = "Human" Then
						CurrentPlayers(5).Human = True
					Else
						CurrentPlayers(5).Human = False
					End If
					CurrentPlayers(5).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(5).Piece = PlayerSixPieceDrop.Text
					'Player 7
					CurrentPlayers(6).Name = PlayerSevenNametxt.Text
					CurrentPlayers(6).GetOutJailCards = 0
					If PlayerSevenStateDrop.Text = "Human" Then
						CurrentPlayers(6).Human = True
					Else
						CurrentPlayers(6).Human = False
					End If
					CurrentPlayers(6).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(6).Piece = PlayerSevenPieceDrop.Text
					'Player 8
					CurrentPlayers(7).Name = PlayerEightNametxt.Text
					CurrentPlayers(7).GetOutJailCards = 0
					If PlayerEightStateDrop.Text = "Human" Then
						CurrentPlayers(7).Human = True
					Else
						CurrentPlayers(7).Human = False
					End If
					CurrentPlayers(7).Cash = CurrentHouseRules.StartingCash
					CurrentPlayers(7).Piece = PlayerEightPieceDrop.Text
			End Select
		Catch ex As Exception
			Return False
		End Try

		For i = 0 To PlayerNumber - 1
			If CurrentPlayers(i).Name = "" Or CurrentPlayers(i).Name = " " Then
				Return False
			End If
			If CurrentPlayers(i).Piece = "None" Or CurrentPlayers(i).Piece = "" Or CurrentPlayers(i).Piece = "Piece" Then
				Return False
			End If
			If CurrentPlayers(i).Piece <> "Wheelbarrow" And CurrentPlayers(i).Piece <> "Battleship" And CurrentPlayers(i).Piece <> "Racecar" And CurrentPlayers(i).Piece <> "Thimble" And CurrentPlayers(i).Piece <> "Boot" And CurrentPlayers(i).Piece <> "Dog" And CurrentPlayers(i).Piece <> "Top Hat" And CurrentPlayers(i).Piece <> "Iron" Then
				Return False
			End If

			Dim namesearching As String = CurrentPlayers(i).Name
			Dim foundcount As Integer = 0
			For n = 0 To PlayerNumber - 1
				If CurrentPlayers(n).Name = namesearching Then
					foundcount += 1
				End If
				If CurrentPlayers(n).Name = "N/A" Then
					Return False
				End If
			Next
			If foundcount > 1 Then
				Return False
			End If
		Next
		Return True
	End Function
	Private Function SetHouseRules() As Boolean

		If FPMoneyCheck.Checked = True Then
			CurrentHouseRules.FreeParkingMoney = True
			If FPTaxesCheck.Checked = True Then
				CurrentHouseRules.FreeParkingTaxCollect = True
			Else
				CurrentHouseRules.FreeParkingTaxCollect = False
			End If
		Else
			CurrentHouseRules.FreeParkingMoney = False
		End If

		If IsNumeric(StartingCashBox.Text) Then
			CurrentHouseRules.StartingCash = CInt(StartingCashBox.Text)
		Else
			Return False
		End If

		If JailRentCheck.Checked = True Then
			CurrentHouseRules.JailRent = True
		Else
			CurrentHouseRules.JailRent = False
		End If

		If GoMultiplierDrop.SelectedItem = "1x = 200" Then
			CurrentHouseRules.GOmulitplier = 1
		ElseIf GoMultiplierDrop.SelectedItem = "2x = 400" Then
			CurrentHouseRules.GOmulitplier = 2
		ElseIf GoMultiplierDrop.SelectedItem = "3x = 600" Then
			CurrentHouseRules.GOmulitplier = 3
		ElseIf GoMultiplierDrop.SelectedItem = "4x = 800" Then
			CurrentHouseRules.GOmulitplier = 4
		Else
			Return False
		End If

		Return True
	End Function

	Private Sub UpdatePlayerStats()

		'Pictures
		If CurrentPlayers(0).Piece = "Wheelbarrow" Then
			PlayerOneStatsPicture.Image = My.Resources.Wheelbarrow
		ElseIf CurrentPlayers(0).Piece = "Battleship" Then
			PlayerOneStatsPicture.Image = My.Resources.Ship
		ElseIf CurrentPlayers(0).Piece = "Racecar" Then
			PlayerOneStatsPicture.Image = My.Resources.Car
		ElseIf CurrentPlayers(0).Piece = "Thimble" Then
			PlayerOneStatsPicture.Image = My.Resources.Thimble
		ElseIf CurrentPlayers(0).Piece = "Boot" Then
			PlayerOneStatsPicture.Image = My.Resources.Boot
		ElseIf CurrentPlayers(0).Piece = "Dog" Then
			PlayerOneStatsPicture.Image = My.Resources.Dog
		ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
			PlayerOneStatsPicture.Image = My.Resources.Hat
		ElseIf CurrentPlayers(0).Piece = "Iron" Then
			PlayerOneStatsPicture.Image = My.Resources.Iron
		End If

		If CurrentPlayers(1).Piece = "Wheelbarrow" Then
			PlayerTwoStatsPicture.Image = My.Resources.Wheelbarrow
		ElseIf CurrentPlayers(1).Piece = "Battleship" Then
			PlayerTwoStatsPicture.Image = My.Resources.Ship
		ElseIf CurrentPlayers(1).Piece = "Racecar" Then
			PlayerTwoStatsPicture.Image = My.Resources.Car
		ElseIf CurrentPlayers(1).Piece = "Thimble" Then
			PlayerTwoStatsPicture.Image = My.Resources.Thimble
		ElseIf CurrentPlayers(1).Piece = "Boot" Then
			PlayerTwoStatsPicture.Image = My.Resources.Boot
		ElseIf CurrentPlayers(1).Piece = "Dog" Then
			PlayerTwoStatsPicture.Image = My.Resources.Dog
		ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
			PlayerTwoStatsPicture.Image = My.Resources.Hat
		ElseIf CurrentPlayers(1).Piece = "Iron" Then
			PlayerTwoStatsPicture.Image = My.Resources.Iron
		End If
		If PlayerNumber >= 3 Then
			If CurrentPlayers(2).Piece = "Wheelbarrow" Then
				PlayerThreeStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(2).Piece = "Battleship" Then
				PlayerThreeStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(2).Piece = "Racecar" Then
				PlayerThreeStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(2).Piece = "Thimble" Then
				PlayerThreeStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(2).Piece = "Boot" Then
				PlayerThreeStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(2).Piece = "Dog" Then
				PlayerThreeStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
				PlayerThreeStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(2).Piece = "Iron" Then
				PlayerThreeStatsPicture.Image = My.Resources.Iron
			End If
		End If
		If PlayerNumber >= 4 Then
			If CurrentPlayers(3).Piece = "Wheelbarrow" Then
				PlayerFourStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(3).Piece = "Battleship" Then
				PlayerFourStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(3).Piece = "Racecar" Then
				PlayerFourStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(3).Piece = "Thimble" Then
				PlayerFourStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(3).Piece = "Boot" Then
				PlayerFourStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(3).Piece = "Dog" Then
				PlayerFourStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
				PlayerFourStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(3).Piece = "Iron" Then
				PlayerFourStatsPicture.Image = My.Resources.Iron
			End If
		End If
		If PlayerNumber >= 5 Then
			If CurrentPlayers(4).Piece = "Wheelbarrow" Then
				PlayerFiveStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(4).Piece = "Battleship" Then
				PlayerFiveStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(4).Piece = "Racecar" Then
				PlayerFiveStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(4).Piece = "Thimble" Then
				PlayerFiveStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(4).Piece = "Boot" Then
				PlayerFiveStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(4).Piece = "Dog" Then
				PlayerFiveStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(4).Piece = "Top Hat" Then
				PlayerFiveStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(4).Piece = "Iron" Then
				PlayerFiveStatsPicture.Image = My.Resources.Iron
			End If
		End If
		If PlayerNumber >= 6 Then

			If CurrentPlayers(5).Piece = "Wheelbarrow" Then
				PlayerSixStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(5).Piece = "Battleship" Then
				PlayerSixStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(5).Piece = "Racecar" Then
				PlayerSixStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(5).Piece = "Thimble" Then
				PlayerSixStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(5).Piece = "Boot" Then
				PlayerSixStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(5).Piece = "Dog" Then
				PlayerSixStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(5).Piece = "Top Hat" Then
				PlayerSixStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(5).Piece = "Iron" Then
				PlayerSixStatsPicture.Image = My.Resources.Iron
			End If
		End If
		If PlayerNumber >= 7 Then
			If CurrentPlayers(6).Piece = "Wheelbarrow" Then
				PlayerSevenStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(6).Piece = "Battleship" Then
				PlayerSevenStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(6).Piece = "Racecar" Then
				PlayerSevenStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(6).Piece = "Thimble" Then
				PlayerSevenStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(6).Piece = "Boot" Then
				PlayerSevenStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(6).Piece = "Dog" Then
				PlayerSevenStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(6).Piece = "Top Hat" Then
				PlayerSevenStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(6).Piece = "Iron" Then
				PlayerSevenStatsPicture.Image = My.Resources.Iron
			End If
		End If
		If PlayerNumber >= 8 Then
			If CurrentPlayers(7).Piece = "Wheelbarrow" Then
				PlayerEightStatsPicture.Image = My.Resources.Wheelbarrow
			ElseIf CurrentPlayers(7).Piece = "Battleship" Then
				PlayerEightStatsPicture.Image = My.Resources.Ship
			ElseIf CurrentPlayers(7).Piece = "Racecar" Then
				PlayerEightStatsPicture.Image = My.Resources.Car
			ElseIf CurrentPlayers(7).Piece = "Thimble" Then
				PlayerEightStatsPicture.Image = My.Resources.Thimble
			ElseIf CurrentPlayers(7).Piece = "Boot" Then
				PlayerEightStatsPicture.Image = My.Resources.Boot
			ElseIf CurrentPlayers(7).Piece = "Dog" Then
				PlayerEightStatsPicture.Image = My.Resources.Dog
			ElseIf CurrentPlayers(7).Piece = "Top Hat" Then
				PlayerEightStatsPicture.Image = My.Resources.Hat
			ElseIf CurrentPlayers(7).Piece = "Iron" Then
				PlayerEightStatsPicture.Image = My.Resources.Iron
			End If
		End If

		'Names
		PlayerOneStatsName.Text = CurrentPlayers(0).Name
		PlayerTwoStatsName.Text = CurrentPlayers(1).Name
		If PlayerNumber >= 3 Then
			PlayerThreeStatsName.Text = CurrentPlayers(2).Name
		End If
		If PlayerNumber >= 4 Then
			PlayerFourStatsName.Text = CurrentPlayers(3).Name
		End If
		If PlayerNumber >= 5 Then
			PlayerFiveStatsName.Text = CurrentPlayers(4).Name
		End If
		If PlayerNumber >= 6 Then
			PlayerSixStatsName.Text = CurrentPlayers(5).Name
		End If
		If PlayerNumber >= 7 Then
			PlayerSevenStatsName.Text = CurrentPlayers(6).Name
		End If
		If PlayerNumber >= 8 Then
			PlayerEightStatsName.Text = CurrentPlayers(7).Name
		End If


		'Cash
		PlayerOneStatsCash.Text = "£" & CurrentPlayers(0).Cash
		PlayerTwoStatsCash.Text = "£" & CurrentPlayers(1).Cash
		If PlayerNumber >= 3 Then
			PlayerThreeStatsCash.Text = "£" & CurrentPlayers(2).Cash
		End If
		If PlayerNumber >= 4 Then
			PlayerFourStatsCash.Text = "£" & CurrentPlayers(3).Cash
		End If
		If PlayerNumber >= 5 Then
			PlayerFiveStatsCash.Text = "£" & CurrentPlayers(4).Cash
		End If
		If PlayerNumber >= 6 Then
			PlayerSixStatsCash.Text = "£" & CurrentPlayers(5).Cash
		End If
		If PlayerNumber >= 7 Then
			PlayerSevenStatsCash.Text = "£" & CurrentPlayers(6).Cash
		End If
		If PlayerNumber >= 8 Then
			PlayerEightStatsCash.Text = "£" & CurrentPlayers(7).Cash
		End If

		PlayerOneStatsPicture.BringToFront()
		PlayerTwoStatsPicture.BringToFront()
		PlayerThreeStatsPicture.BringToFront()
		PlayerFourStatsPicture.BringToFront()
		PlayerFiveStatsPicture.BringToFront()
		PlayerSixStatsPicture.BringToFront()
		PlayerSevenStatsPicture.BringToFront()
		PlayerEightStatsPicture.BringToFront()

		Select Case PlayerNumber
			Case 2
				PlayerThreeStatsCash.Hide()
				PlayerThreeStatsName.Hide()
				PlayerThreeStatsPicture.Hide()
				PlayerFourStatsCash.Hide()
				PlayerFourStatsName.Hide()
				PlayerFourStatsPicture.Hide()
				PlayerFiveStatsCash.Hide()
				PlayerFiveStatsName.Hide()
				PlayerFiveStatsPicture.Hide()
				PlayerSixStatsCash.Hide()
				PlayerSixStatsName.Hide()
				PlayerSixStatsPicture.Hide()
				PlayerSevenStatsCash.Hide()
				PlayerSevenStatsName.Hide()
				PlayerSevenStatsPicture.Hide()
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 3
				PlayerFourStatsCash.Hide()
				PlayerFourStatsName.Hide()
				PlayerFourStatsPicture.Hide()
				PlayerFiveStatsCash.Hide()
				PlayerFiveStatsName.Hide()
				PlayerFiveStatsPicture.Hide()
				PlayerSixStatsCash.Hide()
				PlayerSixStatsName.Hide()
				PlayerSixStatsPicture.Hide()
				PlayerSevenStatsCash.Hide()
				PlayerSevenStatsName.Hide()
				PlayerSevenStatsPicture.Hide()
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 4
				PlayerFiveStatsCash.Hide()
				PlayerFiveStatsName.Hide()
				PlayerFiveStatsPicture.Hide()
				PlayerSixStatsCash.Hide()
				PlayerSixStatsName.Hide()
				PlayerSixStatsPicture.Hide()
				PlayerSevenStatsCash.Hide()
				PlayerSevenStatsName.Hide()
				PlayerSevenStatsPicture.Hide()
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 5
				PlayerSixStatsCash.Hide()
				PlayerSixStatsName.Hide()
				PlayerSixStatsPicture.Hide()
				PlayerSevenStatsCash.Hide()
				PlayerSevenStatsName.Hide()
				PlayerSevenStatsPicture.Hide()
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 6
				PlayerSevenStatsCash.Hide()
				PlayerSevenStatsName.Hide()
				PlayerSevenStatsPicture.Hide()
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 7
				PlayerEightStatsCash.Hide()
				PlayerEightStatsName.Hide()
				PlayerEightStatsPicture.Hide()
			Case 8
				PlayerThreeStatsCash.Show()
				PlayerThreeStatsName.Show()
				PlayerThreeStatsPicture.Show()
				PlayerFourStatsCash.Show()
				PlayerFourStatsName.Show()
				PlayerFourStatsPicture.Show()
				PlayerFiveStatsCash.Show()
				PlayerFiveStatsName.Show()
				PlayerFiveStatsPicture.Show()
				PlayerSixStatsCash.Show()
				PlayerSixStatsName.Show()
				PlayerSixStatsPicture.Show()
				PlayerSevenStatsCash.Show()
				PlayerSevenStatsName.Show()
				PlayerSevenStatsPicture.Show()
				PlayerEightStatsCash.Show()
				PlayerEightStatsName.Show()
				PlayerEightStatsPicture.Show()



		End Select
	End Sub

	Private Sub DrawAngledText(text As String, container As Label, angle As Integer)
		Dim format = New StringFormat()
		format.Alignment = StringAlignment.Center
		format.LineAlignment = StringAlignment.Center
		format.Trimming = StringTrimming.EllipsisCharacter

		Dim img = New Bitmap(container.Height - 2, container.Width)
		Dim G = Graphics.FromImage(img)

		G.Clear(container.BackColor)

		Dim brush_text = New SolidBrush(container.ForeColor)
		G.TextRenderingHint = System.Drawing.Text.TextRenderingHint.SingleBitPerPixelGridFit
		G.DrawString(text, container.Font, brush_text, New Rectangle(0, 0, img.Width, img.Height), format)
		brush_text.Dispose()

		Select Case angle
			Case 90
				img.RotateFlip(RotateFlipType.Rotate90FlipNone)
			Case 180
				img.RotateFlip(RotateFlipType.Rotate180FlipXY)
			Case 270
				img.RotateFlip(RotateFlipType.Rotate270FlipNone)
		End Select


		container.BackgroundImage = img


	End Sub

	Private Sub LoadGameBoard(newgame As Boolean)

		'Titles for Bottom Row
		Tile1Title.Text = Tiles(1).Title
		Tile2Title.Text = Tiles(2).Title
		Tile3Title.Text = Tiles(3).Title
		Tile4Title.Text = Tiles(4).Title
		Tile5Title.Text = Tiles(5).Title
		Tile6Title.Text = Tiles(6).Title
		Tile7Title.Text = Tiles(7).Title
		Tile8Title.Text = Tiles(8).Title
		Tile9Title.Text = Tiles(9).Title
		If newgame = True Then
			LoadingBar.Value += 5
		End If
		'Costs for Bottom Row
		If Tiles(1).Owner = "" Then
			Tile1Cost.Text = "£" & Tiles(1).Cost
		Else
			Tile1Cost.Text = Tiles(1).Cost
		End If

		If Tiles(3).Owner = "" Then
			Tile3Cost.Text = "£" & Tiles(3).Cost
		Else
			Tile3Cost.Text = Tiles(3).Cost
		End If
		Tile4TaxAmount.Text = "PAY £" & Tiles(4).Rent.TaxAmount

		If Tiles(5).Owner = "" Then
			Tile5Cost.Text = "£" & Tiles(5).Cost
		Else
			Tile5Cost.Text = Tiles(5).Cost
		End If
		If Tiles(6).Owner = "" Then
			Tile6Cost.Text = "£" & Tiles(6).Cost
		Else
			Tile6Cost.Text = Tiles(6).Cost
		End If
		If Tiles(8).Owner = "" Then
			Tile8Cost.Text = "£" & Tiles(8).Cost
		Else
			Tile8Cost.Text = Tiles(8).Cost
		End If
		If Tiles(9).Owner = "" Then
			Tile9Cost.Text = "£" & Tiles(9).Cost
		Else
			Tile9Cost.Text = Tiles(9).Cost
		End If

		If newgame = True Then
			LoadingBar.Value += 5
		End If
		'Colours for Bottom Row
		Tile1Colour.BackColor = Tiles(1).Colour
		Tile3Colour.BackColor = Tiles(3).Colour
		Tile6Colour.BackColor = Tiles(6).Colour
		Tile8Colour.BackColor = Tiles(8).Colour
		Tile9Colour.BackColor = Tiles(9).Colour
		If newgame = True Then
			LoadingBar.Value += 5
		End If
		'Colours for Rest of Board
		Tile11Colour.BackColor = Tiles(11).Colour
		Tile13Colour.BackColor = Tiles(13).Colour
		Tile14Colour.BackColor = Tiles(14).Colour
		Tile16Colour.BackColor = Tiles(16).Colour
		Tile18Colour.BackColor = Tiles(18).Colour
		Tile19Colour.BackColor = Tiles(19).Colour
		Tile21Colour.BackColor = Tiles(21).Colour
		Tile23Colour.BackColor = Tiles(23).Colour
		Tile24Colour.BackColor = Tiles(24).Colour
		Tile26Colour.BackColor = Tiles(26).Colour
		Tile27Colour.BackColor = Tiles(27).Colour
		Tile29Colour.BackColor = Tiles(29).Colour
		Tile31Colour.BackColor = Tiles(31).Colour
		Tile32Colour.BackColor = Tiles(32).Colour
		Tile34Colour.BackColor = Tiles(34).Colour
		Tile37Colour.BackColor = Tiles(37).Colour
		Tile39Colour.BackColor = Tiles(39).Colour
		If newgame = True Then
			LoadingBar.Value += 5
		End If

		'Left Side Titles
		DrawAngledText(Tiles(11).Title, Tile11Title, 90)
		DrawAngledText(Tiles(12).Title, Tile12Title, 90)
		DrawAngledText(Tiles(13).Title, Tile13Title, 90)
		DrawAngledText(Tiles(14).Title, Tile14Title, 90)
		DrawAngledText(Tiles(15).Title, Tile15Title, 90)
		DrawAngledText(Tiles(16).Title, Tile16Title, 90)
		DrawAngledText(Tiles(17).Title, Tile17Title, 90)
		DrawAngledText(Tiles(18).Title, Tile18Title, 90)
		DrawAngledText(Tiles(19).Title, Tile19Title, 90)
		If newgame = True Then
			LoadingBar.Value += 5
		End If

		'Left Side Costs
		If Tiles(11).Owner = "" Then
			DrawAngledText(("£" & (Tiles(11).Cost)), Tile11Cost, 90)
		Else
			DrawAngledText(((Tiles(11).Cost)), Tile11Cost, 90)
		End If

		If Tiles(12).Owner = "" Then
			DrawAngledText(("£" & (Tiles(12).Cost)), Tile12Cost, 90)
		Else
			DrawAngledText(((Tiles(12).Cost)), Tile12Cost, 90)
		End If
		If Tiles(13).Owner = "" Then
			DrawAngledText(("£" & (Tiles(13).Cost)), Tile13Cost, 90)
		Else
			DrawAngledText(((Tiles(13).Cost)), Tile13Cost, 90)
		End If
		If Tiles(14).Owner = "" Then
			DrawAngledText(("£" & (Tiles(14).Cost)), Tile14Cost, 90)
		Else
			DrawAngledText(((Tiles(14).Cost)), Tile14Cost, 90)
		End If
		If Tiles(15).Owner = "" Then
			DrawAngledText(("£" & (Tiles(15).Cost)), Tile15Cost, 90)
		Else
			DrawAngledText(((Tiles(15).Cost)), Tile15Cost, 90)
		End If
		If Tiles(16).Owner = "" Then
			DrawAngledText(("£" & (Tiles(16).Cost)), Tile16Cost, 90)
		Else
			DrawAngledText(((Tiles(16).Cost)), Tile16Cost, 90)
		End If
		If Tiles(18).Owner = "" Then
			DrawAngledText(("£" & (Tiles(18).Cost)), Tile18Cost, 90)
		Else
			DrawAngledText(((Tiles(18).Cost)), Tile18Cost, 90)
		End If
		If Tiles(19).Owner = "" Then
			DrawAngledText(("£" & (Tiles(19).Cost)), Tile19Cost, 90)
		Else
			DrawAngledText(((Tiles(19).Cost)), Tile19Cost, 90)
		End If
		If newgame = True Then
			LoadingBar.Value += 5
		End If

		'Top Side Titles
		Tile21Title.Text = Tiles(21).Title
		Tile22Title.Text = Tiles(22).Title
		Tile23Title.Text = Tiles(23).Title
		Tile24Title.Text = Tiles(24).Title
		Tile25Title.Text = Tiles(25).Title
		Tile26Title.Text = Tiles(26).Title
		Tile27Title.Text = Tiles(27).Title
		Tile28Title.Text = Tiles(28).Title
		Tile29Title.Text = Tiles(29).Title
		If newgame = True Then
			LoadingBar.Value += 5
		End If

		'Top Side Costs
		If Tiles(21).Owner = "" Then
			Tile21Cost.Text = "£" & Tiles(21).Cost
		Else
			Tile21Cost.Text = Tiles(21).Cost
		End If

		If Tiles(23).Owner = "" Then
			Tile23Cost.Text = "£" & Tiles(23).Cost
		Else
			Tile23Cost.Text = Tiles(23).Cost
		End If
		If Tiles(24).Owner = "" Then
			Tile24Cost.Text = "£" & Tiles(24).Cost
		Else
			Tile24Cost.Text = Tiles(24).Cost
		End If
		If Tiles(25).Owner = "" Then
			Tile25Cost.Text = "£" & Tiles(25).Cost
		Else
			Tile25Cost.Text = Tiles(25).Cost
		End If
		If Tiles(26).Owner = "" Then
			Tile26Cost.Text = "£" & Tiles(26).Cost
		Else
			Tile26Cost.Text = Tiles(26).Cost
		End If
		If Tiles(27).Owner = "" Then
			Tile27Cost.Text = "£" & Tiles(27).Cost
		Else
			Tile27Cost.Text = Tiles(27).Cost
		End If
		If Tiles(28).Owner = "" Then
			Tile28Cost.Text = "£" & Tiles(28).Cost
		Else
			Tile28Cost.Text = Tiles(28).Cost
		End If
		If Tiles(29).Owner = "" Then
			Tile29Cost.Text = "£" & Tiles(29).Cost
		Else
			Tile29Cost.Text = Tiles(29).Cost
		End If
		If newgame = True Then
			LoadingBar.Value += 5
		End If

		'Right Side Titles
		DrawAngledText(Tiles(31).Title, Tile31Title, 270)
		DrawAngledText(Tiles(32).Title, Tile32Title, 270)
		DrawAngledText(Tiles(33).Title, Tile33Title, 270)
		DrawAngledText(Tiles(34).Title, Tile34Title, 270)
		DrawAngledText(Tiles(35).Title, Tile35Title, 270)
		DrawAngledText(Tiles(36).Title, Tile36Title, 270)
		DrawAngledText(Tiles(37).Title, Tile37Title, 270)
		DrawAngledText(Tiles(38).Title, Tile38Title, 270)
		DrawAngledText(Tiles(39).Title, Tile39Title, 270)
		If newgame = True Then
			LoadingBar.Value += 5
		End If
		'Right Side Costs
		If Tiles(31).Owner = "" Then
			DrawAngledText(("£" & (Tiles(31).Cost)), Tile31Cost, 270)
		Else
			DrawAngledText(((Tiles(31).Cost)), Tile31Cost, 270)
		End If

		If Tiles(32).Owner = "" Then
			DrawAngledText(("£" & (Tiles(32).Cost)), Tile32Cost, 270)
		Else
			DrawAngledText(((Tiles(32).Cost)), Tile32Cost, 270)
		End If
		If Tiles(34).Owner = "" Then
			DrawAngledText(("£" & (Tiles(34).Cost)), Tile34Cost, 270)
		Else
			DrawAngledText(((Tiles(34).Cost)), Tile34Cost, 270)
		End If
		If Tiles(35).Owner = "" Then
			DrawAngledText(("£" & (Tiles(35).Cost)), Tile35Cost, 270)
		Else
			DrawAngledText(((Tiles(35).Cost)), Tile35Cost, 270)
		End If
		If Tiles(37).Owner = "" Then
			DrawAngledText(("£" & (Tiles(37).Cost)), Tile37Cost, 270)
		Else
			DrawAngledText(((Tiles(37).Cost)), Tile37Cost, 270)
		End If
		DrawAngledText(("PAY £" & (Tiles(38).Rent.TaxAmount)), Tile38TaxAmount, 270)
		If Tiles(39).Owner = "" Then
			DrawAngledText(("£" & (Tiles(39).Cost)), Tile39Cost, 270)
		Else
			DrawAngledText(((Tiles(39).Cost)), Tile39Cost, 270)
		End If
		If newgame = True Then
			LoadingBar.Value += 5
		End If


		'Mortgaged Properties Bottom

		If Tiles(1).Mortgaged = True Then
			Dim imagelbl1 As New PictureBox
			imagelbl1.Height = Tile1Cost.Top - (Tile1Title.Top + Tile1Title.Height)
			imagelbl1.Width = 75
			imagelbl1.Image = My.Resources.MortgagedBottom
			imagelbl1.BackColor = Color.Transparent
			imagelbl1.BringToFront()
			imagelbl1.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl1)
			imagelbl1.Parent = Tile1
			imagelbl1.Top = Tile1Title.Top + Tile1Title.Height
			imagelbl1.Left = 0
			imagelbl1.Tag = "Mortgaged"
		End If

		If Tiles(3).Mortgaged = True Then
			Dim imagelbl3 As New PictureBox
			imagelbl3.Height = Tile3Cost.Top - (Tile3Title.Top + Tile3Title.Height)
			imagelbl3.Width = 75
			imagelbl3.Image = My.Resources.MortgagedBottom
			imagelbl3.BackColor = Color.Transparent
			imagelbl3.BringToFront()
			imagelbl3.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl3)
			imagelbl3.Parent = Tile3
			imagelbl3.Top = Tile3Title.Top + Tile3Title.Height
			imagelbl3.Left = 0
			imagelbl3.Tag = "Mortgaged"
		End If

		If Tiles(6).Mortgaged = True Then
			Dim imagelbl6 As New PictureBox
			imagelbl6.Height = Tile6Cost.Top - (Tile6Title.Top + Tile6Title.Height)
			imagelbl6.Width = 75
			imagelbl6.Image = My.Resources.MortgagedBottom
			imagelbl6.BackColor = Color.Transparent
			imagelbl6.BringToFront()
			imagelbl6.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl6)
			imagelbl6.Parent = Tile6
			imagelbl6.Top = Tile6Title.Top + Tile6Title.Height
			imagelbl6.Left = 0
			imagelbl6.Tag = "Mortgaged"
		End If

		If Tiles(8).Mortgaged = True Then
			Dim imagelbl8 As New PictureBox
			imagelbl8.Height = Tile8Cost.Top - (Tile8Title.Top + Tile8Title.Height)
			imagelbl8.Width = 75
			imagelbl8.Image = My.Resources.MortgagedBottom
			imagelbl8.BackColor = Color.Transparent
			imagelbl8.BringToFront()
			imagelbl8.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl8)
			imagelbl8.Parent = Tile8
			imagelbl8.Top = Tile8Title.Top + Tile8Title.Height
			imagelbl8.Left = 0
			imagelbl8.Tag = "Mortgaged"
		End If

		If Tiles(9).Mortgaged = True Then
			Dim imagelbl9 As New PictureBox
			imagelbl9.Height = Tile9Cost.Top - (Tile9Title.Top + Tile9Title.Height)
			imagelbl9.Width = 75
			imagelbl9.Image = My.Resources.MortgagedBottom
			imagelbl9.BackColor = Color.Transparent
			imagelbl9.BringToFront()
			imagelbl9.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl9)
			imagelbl9.Parent = Tile9
			imagelbl9.Top = Tile9Title.Top + Tile9Title.Height
			imagelbl9.Left = 0
			imagelbl9.Tag = "Mortgaged"
		End If








		'Mortgaged Properties Left
		If Tiles(11).Mortgaged = True Then
			Dim imagelbl11 As New PictureBox
			imagelbl11.Height = 60
			imagelbl11.Width = Tile11Title.Left - (Tile11Cost.Left + Tile11Cost.Width)
			imagelbl11.Image = My.Resources.MortgagedLeft
			imagelbl11.BackColor = Color.Transparent
			imagelbl11.BringToFront()
			imagelbl11.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl11)
			imagelbl11.Parent = Tile11
			imagelbl11.Top = 0
			imagelbl11.Left = Tile11Cost.Left + Tile11Cost.Width
			imagelbl11.Tag = "Mortgaged"
		End If

		If Tiles(13).Mortgaged = True Then
			Dim imagelbl13 As New PictureBox
			imagelbl13.Height = 60
			imagelbl13.Width = Tile13Title.Left - (Tile13Cost.Left + Tile13Cost.Width)
			imagelbl13.Image = My.Resources.MortgagedLeft
			imagelbl13.BackColor = Color.Transparent
			imagelbl13.BringToFront()
			imagelbl13.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl13)
			imagelbl13.Parent = Tile13
			imagelbl13.Top = 0
			imagelbl13.Left = Tile13Cost.Left + Tile13Cost.Width
			imagelbl13.Tag = "Mortgaged"
		End If

		If Tiles(14).Mortgaged = True Then
			Dim imagelbl14 As New PictureBox
			imagelbl14.Height = 60
			imagelbl14.Width = Tile14Title.Left - (Tile14Cost.Left + Tile14Cost.Width)
			imagelbl14.Image = My.Resources.MortgagedLeft
			imagelbl14.BackColor = Color.Transparent
			imagelbl14.BringToFront()
			imagelbl14.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl14)
			imagelbl14.Parent = Tile14
			imagelbl14.Top = 0
			imagelbl14.Left = Tile14Cost.Left + Tile14Cost.Width
			imagelbl14.Tag = "Mortgaged"
		End If

		If Tiles(16).Mortgaged = True Then
			Dim imagelbl16 As New PictureBox
			imagelbl16.Height = 60
			imagelbl16.Width = Tile16Title.Left - (Tile16Cost.Left + Tile16Cost.Width)
			imagelbl16.Image = My.Resources.MortgagedLeft
			imagelbl16.BackColor = Color.Transparent
			imagelbl16.BringToFront()
			imagelbl16.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl16)
			imagelbl16.Parent = Tile16
			imagelbl16.Top = 0
			imagelbl16.Left = Tile16Cost.Left + Tile16Cost.Width
			imagelbl16.Tag = "Mortgaged"
		End If

		If Tiles(18).Mortgaged = True Then
			Dim imagelbl18 As New PictureBox
			imagelbl18.Height = 60
			imagelbl18.Width = Tile18Title.Left - (Tile18Cost.Left + Tile18Cost.Width)
			imagelbl18.Image = My.Resources.MortgagedLeft
			imagelbl18.BackColor = Color.Transparent
			imagelbl18.BringToFront()
			imagelbl18.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl18)
			imagelbl18.Parent = Tile18
			imagelbl18.Top = 0
			imagelbl18.Left = Tile18Cost.Left + Tile18Cost.Width
			imagelbl18.Tag = "Mortgaged"
		End If

		If Tiles(19).Mortgaged = True Then
			Dim imagelbl19 As New PictureBox
			imagelbl19.Height = 60
			imagelbl19.Width = Tile19Title.Left - (Tile19Cost.Left + Tile19Cost.Width)
			imagelbl19.Image = My.Resources.MortgagedLeft
			imagelbl19.BackColor = Color.Transparent
			imagelbl19.BringToFront()
			imagelbl19.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl19)
			imagelbl19.Parent = Tile19
			imagelbl19.Top = 0
			imagelbl19.Left = Tile19Cost.Left + Tile19Cost.Width
			imagelbl19.Tag = "Mortgaged"
		End If


		'Mortgaged Properties Top
		If Tiles(21).Mortgaged = True Then
			Dim imagelbl21 As New PictureBox
			imagelbl21.Height = Tile21Title.Top - (Tile21Cost.Top + Tile21Cost.Height)
			imagelbl21.Width = 75
			imagelbl21.Image = My.Resources.MortgagedBottom
			imagelbl21.BackColor = Color.Transparent
			imagelbl21.BringToFront()
			imagelbl21.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl21)
			imagelbl21.Parent = Tile21
			imagelbl21.Top = Tile21Cost.Top + Tile21Cost.Height
			imagelbl21.Left = 0
			imagelbl21.Tag = "Mortgaged"
		End If

		If Tiles(23).Mortgaged = True Then
			Dim imagelbl23 As New PictureBox
			imagelbl23.Height = Tile23Title.Top - (Tile23Cost.Top + Tile23Cost.Height)
			imagelbl23.Width = 75
			imagelbl23.Image = My.Resources.MortgagedBottom
			imagelbl23.BackColor = Color.Transparent
			imagelbl23.BringToFront()
			imagelbl23.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl23)
			imagelbl23.Parent = Tile23
			imagelbl23.Top = Tile23Cost.Top + Tile23Cost.Height
			imagelbl23.Left = 0
			imagelbl23.Tag = "Mortgaged"
		End If

		If Tiles(24).Mortgaged = True Then
			Dim imagelbl24 As New PictureBox
			imagelbl24.Height = Tile24Title.Top - (Tile24Cost.Top + Tile24Cost.Height)
			imagelbl24.Width = 75
			imagelbl24.Image = My.Resources.MortgagedBottom
			imagelbl24.BackColor = Color.Transparent
			imagelbl24.BringToFront()
			imagelbl24.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl24)
			imagelbl24.Parent = Tile24
			imagelbl24.Top = Tile24Cost.Top + Tile24Cost.Height
			imagelbl24.Left = 0
			imagelbl24.Tag = "Mortgaged"
		End If

		If Tiles(26).Mortgaged = True Then
			Dim imagelbl26 As New PictureBox
			imagelbl26.Height = Tile26Title.Top - (Tile26Cost.Top + Tile26Cost.Height)
			imagelbl26.Width = 75
			imagelbl26.Image = My.Resources.MortgagedBottom
			imagelbl26.BackColor = Color.Transparent
			imagelbl26.BringToFront()
			imagelbl26.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl26)
			imagelbl26.Parent = Tile26
			imagelbl26.Top = Tile26Cost.Top + Tile26Cost.Height
			imagelbl26.Left = 0
			imagelbl26.Tag = "Mortgaged"
		End If

		If Tiles(27).Mortgaged = True Then
			Dim imagelbl27 As New PictureBox
			imagelbl27.Height = Tile27Title.Top - (Tile27Cost.Top + Tile27Cost.Height)
			imagelbl27.Width = 75
			imagelbl27.Image = My.Resources.MortgagedBottom
			imagelbl27.BackColor = Color.Transparent
			imagelbl27.BringToFront()
			imagelbl27.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl27)
			imagelbl27.Parent = Tile27
			imagelbl27.Top = Tile27Cost.Top + Tile27Cost.Height
			imagelbl27.Left = 0
			imagelbl27.Tag = "Mortgaged"
		End If

		If Tiles(29).Mortgaged = True Then
			Dim imagelbl29 As New PictureBox
			imagelbl29.Height = Tile29Title.Top - (Tile29Cost.Top + Tile29Cost.Height)
			imagelbl29.Width = 75
			imagelbl29.Image = My.Resources.MortgagedBottom
			imagelbl29.BackColor = Color.Transparent
			imagelbl29.BringToFront()
			imagelbl29.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl29)
			imagelbl29.Parent = Tile29
			imagelbl29.Top = Tile29Cost.Top + Tile29Cost.Height
			imagelbl29.Left = 0
			imagelbl29.Tag = "Mortgaged"
		End If

		'Mortgaged Properties Right
		If Tiles(31).Mortgaged = True Then
			Dim imagelbl31 As New PictureBox
			imagelbl31.Height = 60
			imagelbl31.Width = Tile31Cost.Left - (Tile31Title.Left + Tile31Title.Width)
			imagelbl31.Image = My.Resources.MortgagedRight
			imagelbl31.BackColor = Color.Transparent
			imagelbl31.BringToFront()
			imagelbl31.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl31)
			imagelbl31.Parent = Tile31
			imagelbl31.Top = 0
			imagelbl31.Left = Tile31Title.Left + Tile31Title.Width
			imagelbl31.Tag = "Mortgaged"
		End If

		If Tiles(32).Mortgaged = True Then
			Dim imagelbl32 As New PictureBox
			imagelbl32.Height = 60
			imagelbl32.Width = Tile32Cost.Left - (Tile32Title.Left + Tile32Title.Width)
			imagelbl32.Image = My.Resources.MortgagedRight
			imagelbl32.BackColor = Color.Transparent
			imagelbl32.BringToFront()
			imagelbl32.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl32)
			imagelbl32.Parent = Tile32
			imagelbl32.Top = 0
			imagelbl32.Left = Tile32Title.Left + Tile32Title.Width
			imagelbl32.Tag = "Mortgaged"
		End If

		If Tiles(34).Mortgaged = True Then
			Dim imagelbl34 As New PictureBox
			imagelbl34.Height = 60
			imagelbl34.Width = Tile34Cost.Left - (Tile34Title.Left + Tile34Title.Width)
			imagelbl34.Image = My.Resources.MortgagedRight
			imagelbl34.BackColor = Color.Transparent
			imagelbl34.BringToFront()
			imagelbl34.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl34)
			imagelbl34.Parent = Tile34
			imagelbl34.Top = 0
			imagelbl34.Left = Tile34Title.Left + Tile34Title.Width
			imagelbl34.Tag = "Mortgaged"
		End If

		If Tiles(37).Mortgaged = True Then
			Dim imagelbl37 As New PictureBox
			imagelbl37.Height = 60
			imagelbl37.Width = Tile37Cost.Left - (Tile37Title.Left + Tile37Title.Width)
			imagelbl37.Image = My.Resources.MortgagedRight
			imagelbl37.BackColor = Color.Transparent
			imagelbl37.BringToFront()
			imagelbl37.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl37)
			imagelbl37.Parent = Tile37
			imagelbl37.Top = 0
			imagelbl37.Left = Tile37Title.Left + Tile37Title.Width
			imagelbl37.Tag = "Mortgaged"
		End If


		If Tiles(37).Mortgaged = True Then
			Dim imagelbl37 As New PictureBox
			imagelbl37.Height = 60
			imagelbl37.Width = Tile37Cost.Left - (Tile37Title.Left + Tile37Title.Width)
			imagelbl37.Image = My.Resources.MortgagedRight
			imagelbl37.BackColor = Color.Transparent
			imagelbl37.BringToFront()
			imagelbl37.SizeMode = PictureBoxSizeMode.Zoom
			Me.Controls.Add(imagelbl37)
			imagelbl37.Parent = Tile37
			imagelbl37.Top = 0
			imagelbl37.Left = Tile37Title.Left + Tile37Title.Width
			imagelbl37.Tag = "Mortgaged"
		End If


		'RailWays Mortgaged?
		If Tiles(5).Mortgaged = True Then
			Tile5Art.Image = My.Resources.MortgagedBottom
		Else
			Tile5Art.Image = My.Resources.BottomRail
		End If

		If Tiles(15).Mortgaged = True Then
			Tile15Art.Image = My.Resources.MortgagedLeft
		Else
			Tile15Art.Image = My.Resources.LeftRail
		End If

		If Tiles(25).Mortgaged = True Then
			Tile25Art.Image = My.Resources.MortgagedBottom
		Else
			Tile25Art.Image = My.Resources.BottomRail
		End If

		If Tiles(35).Mortgaged = True Then
			Tile5Art.Image = My.Resources.MortgagedRight
		Else
			Tile35Art.Image = My.Resources.RightRail
		End If

		'Utilities Mortgaged?

		If Tiles(12).Mortgaged = True Then
			Tile12Art.Image = My.Resources.MortgagedLeft
		Else
			Tile12Art.Image = My.Resources.ElectricLeft
		End If

		If Tiles(28).Mortgaged = True Then
			Tile28Art.Image = My.Resources.MortgagedBottom
		Else
			Tile28Art.Image = My.Resources.WaterWorks
		End If



		'Load Houses
		Dim housewidth As Integer = 10
		Dim houseshifter As Integer = 0
		For i = 0 To 39
			If Tiles(i).HouseNo > 0 Or Tiles(i).Hotel = True Then
				houseshifter = 0
				Select Case i
					Case 1
						Tile1Colour.Controls.Clear()

						For n = 1 To Tiles(i).HouseNo
							Dim Housepic1 As New PictureBox
							Housepic1.Height = Tile1Colour.Height
							Housepic1.Width = housewidth
							Housepic1.Image = My.Resources.MonopolyHouse
							Housepic1.BackColor = Color.Transparent
							Housepic1.BringToFront()
							Housepic1.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic1)
							Housepic1.Parent = Tile1Colour
							Housepic1.Top = 0
							Housepic1.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic1.Tag = "HousePic1_" & n
						Next

					Case 3
						Tile3Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic3 As New PictureBox
							Housepic3.Height = Tile3Colour.Height
							Housepic3.Width = housewidth
							Housepic3.Image = My.Resources.MonopolyHouse
							Housepic3.BackColor = Color.Transparent
							Housepic3.BringToFront()
							Housepic3.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic3)
							Housepic3.Parent = Tile3Colour
							Housepic3.Top = 0
							Housepic3.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic3.Tag = "HousePic3_" & n
						Next

					Case 6
						Tile6Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic6 As New PictureBox
							Housepic6.Height = Tile6Colour.Height
							Housepic6.Width = housewidth
							Housepic6.Image = My.Resources.MonopolyHouse
							Housepic6.BackColor = Color.Transparent
							Housepic6.BringToFront()
							Housepic6.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic6)
							Housepic6.Parent = Tile6Colour
							Housepic6.Top = 0
							Housepic6.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic6.Tag = "HousePic6_" & n
						Next

					Case 8
						Tile8Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic8 As New PictureBox
							Housepic8.Height = Tile8Colour.Height
							Housepic8.Width = housewidth
							Housepic8.Image = My.Resources.MonopolyHouse
							Housepic8.BackColor = Color.Transparent
							Housepic8.BringToFront()
							Housepic8.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic8)
							Housepic8.Parent = Tile8Colour
							Housepic8.Top = 0
							Housepic8.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic8.Tag = "HousePic8_" & n
						Next

					Case 9
						Tile9Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic9 As New PictureBox
							Housepic9.Height = Tile9Colour.Height
							Housepic9.Width = housewidth
							Housepic9.Image = My.Resources.MonopolyHouse
							Housepic9.BackColor = Color.Transparent
							Housepic9.BringToFront()
							Housepic9.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic9)
							Housepic9.Parent = Tile9Colour
							Housepic9.Top = 0
							Housepic9.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic9.Tag = "HousePic9_" & n
						Next


					Case 11
						Tile11Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic11 As New PictureBox
							Housepic11.Height = housewidth
							Housepic11.Width = Tile11Colour.Width
							Housepic11.Image = My.Resources.MonopolyHouse
							Housepic11.BackColor = Color.Transparent
							Housepic11.BringToFront()
							Housepic11.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic11)
							Housepic11.Parent = Tile11Colour
							Housepic11.Top = houseshifter
							Housepic11.Left = 0
							houseshifter += housewidth + 5
							Housepic11.Tag = "HousePic11_" & n
						Next

					Case 13
						Tile13Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic13 As New PictureBox
							Housepic13.Height = housewidth
							Housepic13.Width = Tile13Colour.Width
							Housepic13.Image = My.Resources.MonopolyHouse
							Housepic13.BackColor = Color.Transparent
							Housepic13.BringToFront()
							Housepic13.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic13)
							Housepic13.Parent = Tile13Colour
							Housepic13.Top = houseshifter
							Housepic13.Left = 0
							houseshifter += housewidth + 5
							Housepic13.Tag = "HousePic13_" & n
						Next

					Case 14
						Tile14Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic14 As New PictureBox
							Housepic14.Height = housewidth
							Housepic14.Width = Tile14Colour.Width
							Housepic14.Image = My.Resources.MonopolyHouse
							Housepic14.BackColor = Color.Transparent
							Housepic14.BringToFront()
							Housepic14.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic14)
							Housepic14.Parent = Tile14Colour
							Housepic14.Top = houseshifter
							Housepic14.Left = 0
							houseshifter += housewidth + 5
							Housepic14.Tag = "HousePic14_" & n
						Next

					Case 16
						Tile16Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic16 As New PictureBox
							Housepic16.Height = housewidth
							Housepic16.Width = Tile16Colour.Width
							Housepic16.Image = My.Resources.MonopolyHouse
							Housepic16.BackColor = Color.Transparent
							Housepic16.BringToFront()
							Housepic16.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic16)
							Housepic16.Parent = Tile16Colour
							Housepic16.Top = houseshifter
							Housepic16.Left = 0
							houseshifter += housewidth + 5
							Housepic16.Tag = "HousePic16_" & n
						Next

					Case 18
						Tile18Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic18 As New PictureBox
							Housepic18.Height = housewidth
							Housepic18.Width = Tile18Colour.Width
							Housepic18.Image = My.Resources.MonopolyHouse
							Housepic18.BackColor = Color.Transparent
							Housepic18.BringToFront()
							Housepic18.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic18)
							Housepic18.Parent = Tile18Colour
							Housepic18.Top = houseshifter
							Housepic18.Left = 0
							houseshifter += housewidth + 5
							Housepic18.Tag = "HousePic18_" & n
						Next

					Case 19
						Tile19Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic19 As New PictureBox
							Housepic19.Height = housewidth
							Housepic19.Width = Tile19Colour.Width
							Housepic19.Image = My.Resources.MonopolyHouse
							Housepic19.BackColor = Color.Transparent
							Housepic19.BringToFront()
							Housepic19.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic19)
							Housepic19.Parent = Tile19Colour
							Housepic19.Top = houseshifter
							Housepic19.Left = 0
							houseshifter += housewidth + 5
							Housepic19.Tag = "HousePic19_" & n
						Next

					Case 21
						Tile21Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic21 As New PictureBox
							Housepic21.Height = Tile21Colour.Height
							Housepic21.Width = housewidth
							Housepic21.Image = My.Resources.MonopolyHouse
							Housepic21.BackColor = Color.Transparent
							Housepic21.BringToFront()
							Housepic21.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic21)
							Housepic21.Parent = Tile21Colour
							Housepic21.Top = 0
							Housepic21.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic21.Tag = "HousePic21_" & n
						Next

					Case 23
						Tile23Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic23 As New PictureBox
							Housepic23.Height = Tile23Colour.Height
							Housepic23.Width = housewidth
							Housepic23.Image = My.Resources.MonopolyHouse
							Housepic23.BackColor = Color.Transparent
							Housepic23.BringToFront()
							Housepic23.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic23)
							Housepic23.Parent = Tile23Colour
							Housepic23.Top = 0
							Housepic23.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic23.Tag = "HousePic23_" & n
						Next

					Case 24
						Tile24Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic24 As New PictureBox
							Housepic24.Height = Tile24Colour.Height
							Housepic24.Width = housewidth
							Housepic24.Image = My.Resources.MonopolyHouse
							Housepic24.BackColor = Color.Transparent
							Housepic24.BringToFront()
							Housepic24.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic24)
							Housepic24.Parent = Tile24Colour
							Housepic24.Top = 0
							Housepic24.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic24.Tag = "HousePic24_" & n
						Next

					Case 26
						Tile26Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic26 As New PictureBox
							Housepic26.Height = Tile26Colour.Height
							Housepic26.Width = housewidth
							Housepic26.Image = My.Resources.MonopolyHouse
							Housepic26.BackColor = Color.Transparent
							Housepic26.BringToFront()
							Housepic26.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic26)
							Housepic26.Parent = Tile26Colour
							Housepic26.Top = 0
							Housepic26.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic26.Tag = "HousePic26_" & n
						Next

					Case 27
						Tile27Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic27 As New PictureBox
							Housepic27.Height = Tile27Colour.Height
							Housepic27.Width = housewidth
							Housepic27.Image = My.Resources.MonopolyHouse
							Housepic27.BackColor = Color.Transparent
							Housepic27.BringToFront()
							Housepic27.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic27)
							Housepic27.Parent = Tile27Colour
							Housepic27.Top = 0
							Housepic27.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic27.Tag = "HousePic27_" & n
						Next

					Case 29
						Tile29Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic29 As New PictureBox
							Housepic29.Height = Tile29Colour.Height
							Housepic29.Width = housewidth
							Housepic29.Image = My.Resources.MonopolyHouse
							Housepic29.BackColor = Color.Transparent
							Housepic29.BringToFront()
							Housepic29.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic29)
							Housepic29.Parent = Tile29Colour
							Housepic29.Top = 0
							Housepic29.Left = houseshifter
							houseshifter += housewidth + 5
							Housepic29.Tag = "HousePic29_" & n
						Next

					Case 31
						Tile31Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic31 As New PictureBox
							Housepic31.Height = housewidth
							Housepic31.Width = Tile31Colour.Width
							Housepic31.Image = My.Resources.MonopolyHouse
							Housepic31.BackColor = Color.Transparent
							Housepic31.BringToFront()
							Housepic31.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic31)
							Housepic31.Parent = Tile31Colour
							Housepic31.Top = houseshifter
							Housepic31.Left = 0
							houseshifter += housewidth + 5
							Housepic31.Tag = "HousePic31_" & n
						Next

					Case 32
						Tile32Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic32 As New PictureBox
							Housepic32.Height = housewidth
							Housepic32.Width = Tile32Colour.Width
							Housepic32.Image = My.Resources.MonopolyHouse
							Housepic32.BackColor = Color.Transparent
							Housepic32.BringToFront()
							Housepic32.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic32)
							Housepic32.Parent = Tile32Colour
							Housepic32.Top = houseshifter
							Housepic32.Left = 0
							houseshifter += housewidth + 5
							Housepic32.Tag = "HousePic32_" & n
						Next

					Case 34
						Tile34Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic34 As New PictureBox
							Housepic34.Height = housewidth
							Housepic34.Width = Tile34Colour.Width
							Housepic34.Image = My.Resources.MonopolyHouse
							Housepic34.BackColor = Color.Transparent
							Housepic34.BringToFront()
							Housepic34.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic34)
							Housepic34.Parent = Tile34Colour
							Housepic34.Top = houseshifter
							Housepic34.Left = 0
							houseshifter += housewidth + 5
							Housepic34.Tag = "HousePic34_" & n
						Next


					Case 37
						Tile37Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic37 As New PictureBox
							Housepic37.Height = housewidth
							Housepic37.Width = Tile37Colour.Width
							Housepic37.Image = My.Resources.MonopolyHouse
							Housepic37.BackColor = Color.Transparent
							Housepic37.BringToFront()
							Housepic37.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic37)
							Housepic37.Parent = Tile37Colour
							Housepic37.Top = houseshifter
							Housepic37.Left = 0
							houseshifter += housewidth + 5
							Housepic37.Tag = "HousePic37_" & n
						Next

					Case 39
						Tile39Colour.Controls.Clear()
						For n = 1 To Tiles(i).HouseNo
							Dim Housepic39 As New PictureBox
							Housepic39.Height = housewidth
							Housepic39.Width = Tile39Colour.Width
							Housepic39.Image = My.Resources.MonopolyHouse
							Housepic39.BackColor = Color.Transparent
							Housepic39.BringToFront()
							Housepic39.SizeMode = PictureBoxSizeMode.Zoom
							Me.Controls.Add(Housepic39)
							Housepic39.Parent = Tile39Colour
							Housepic39.Top = houseshifter
							Housepic39.Left = 0
							houseshifter += housewidth + 5
							Housepic39.Tag = "HousePic39_" & n
						Next
				End Select
			End If
		Next

		'Load Hotels
		For i = 0 To 39
			If Tiles(i).Hotel = True Then
				Select Case i
					Case 1
						Dim HotelPic1 As New PictureBox
						HotelPic1.Height = Tile1Colour.Height
						HotelPic1.Width = housewidth * 2
						HotelPic1.Image = My.Resources.monopoly_hotel
						HotelPic1.BackColor = Color.Transparent
						HotelPic1.BringToFront()
						HotelPic1.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic1)
						HotelPic1.Parent = Tile1Colour
						HotelPic1.Top = 0
						HotelPic1.Left = (Tile1Colour.Width / 2) - HotelPic1.Width
						HotelPic1.Tag = "HotelPic1"
					Case 3
						Dim HotelPic3 As New PictureBox
						HotelPic3.Height = Tile3Colour.Height
						HotelPic3.Width = housewidth * 2
						HotelPic3.Image = My.Resources.monopoly_hotel
						HotelPic3.BackColor = Color.Transparent
						HotelPic3.BringToFront()
						HotelPic3.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic3)
						HotelPic3.Parent = Tile3Colour
						HotelPic3.Top = 0
						HotelPic3.Left = (Tile3Colour.Width / 2) - HotelPic3.Width
						HotelPic3.Tag = "HotelPic3"
					Case 6
						Dim HotelPic6 As New PictureBox
						HotelPic6.Height = Tile6Colour.Height
						HotelPic6.Width = housewidth * 2
						HotelPic6.Image = My.Resources.monopoly_hotel
						HotelPic6.BackColor = Color.Transparent
						HotelPic6.BringToFront()
						HotelPic6.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic6)
						HotelPic6.Parent = Tile6Colour
						HotelPic6.Top = 0
						HotelPic6.Left = (Tile6Colour.Width / 2) - HotelPic6.Width
						HotelPic6.Tag = "HotelPic6"
					Case 8
						Dim HotelPic8 As New PictureBox
						HotelPic8.Height = Tile8Colour.Height
						HotelPic8.Width = housewidth * 2
						HotelPic8.Image = My.Resources.monopoly_hotel
						HotelPic8.BackColor = Color.Transparent
						HotelPic8.BringToFront()
						HotelPic8.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic8)
						HotelPic8.Parent = Tile8Colour
						HotelPic8.Top = 0
						HotelPic8.Left = (Tile8Colour.Width / 2) - HotelPic8.Width
						HotelPic8.Tag = "HotelPic8"
					Case 9
						Dim HotelPic9 As New PictureBox
						HotelPic9.Height = Tile9Colour.Height
						HotelPic9.Width = housewidth * 2
						HotelPic9.Image = My.Resources.monopoly_hotel
						HotelPic9.BackColor = Color.Transparent
						HotelPic9.BringToFront()
						HotelPic9.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic9)
						HotelPic9.Parent = Tile9Colour
						HotelPic9.Top = 0
						HotelPic9.Left = (Tile9Colour.Width / 2) - HotelPic9.Width
						HotelPic9.Tag = "HotelPic9"
					Case 11
						Dim HotelPic11 As New PictureBox
						HotelPic11.Height = housewidth * 2
						HotelPic11.Width = Tile11Colour.Width
						HotelPic11.Image = My.Resources.monopoly_hotel
						HotelPic11.BackColor = Color.Transparent
						HotelPic11.BringToFront()
						HotelPic11.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic11)
						HotelPic11.Parent = Tile11Colour
						HotelPic11.Top = (Tile11Colour.Height / 2) - HotelPic11.Height
						HotelPic11.Left = 0
						HotelPic11.Tag = "HotelPic11"
					Case 13
						Dim HotelPic13 As New PictureBox
						HotelPic13.Height = housewidth * 2
						HotelPic13.Width = Tile13Colour.Width
						HotelPic13.Image = My.Resources.monopoly_hotel
						HotelPic13.BackColor = Color.Transparent
						HotelPic13.BringToFront()
						HotelPic13.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic13)
						HotelPic13.Parent = Tile13Colour
						HotelPic13.Top = (Tile13Colour.Height / 2) - HotelPic13.Height
						HotelPic13.Left = 0
						HotelPic13.Tag = "HotelPic13"
					Case 14
						Dim HotelPic14 As New PictureBox
						HotelPic14.Height = housewidth * 2
						HotelPic14.Width = Tile14Colour.Width
						HotelPic14.Image = My.Resources.monopoly_hotel
						HotelPic14.BackColor = Color.Transparent
						HotelPic14.BringToFront()
						HotelPic14.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic14)
						HotelPic14.Parent = Tile14Colour
						HotelPic14.Top = (Tile14Colour.Height / 2) - HotelPic14.Height
						HotelPic14.Left = 0
						HotelPic14.Tag = "HotelPic14"
					Case 16
						Dim HotelPic16 As New PictureBox
						HotelPic16.Height = housewidth * 2
						HotelPic16.Width = Tile16Colour.Width
						HotelPic16.Image = My.Resources.monopoly_hotel
						HotelPic16.BackColor = Color.Transparent
						HotelPic16.BringToFront()
						HotelPic16.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic16)
						HotelPic16.Parent = Tile16Colour
						HotelPic16.Top = (Tile16Colour.Height / 2) - HotelPic16.Height
						HotelPic16.Left = 0
						HotelPic16.Tag = "HotelPic16"
					Case 18
						Dim HotelPic18 As New PictureBox
						HotelPic18.Height = housewidth * 2
						HotelPic18.Width = Tile18Colour.Width
						HotelPic18.Image = My.Resources.monopoly_hotel
						HotelPic18.BackColor = Color.Transparent
						HotelPic18.BringToFront()
						HotelPic18.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic18)
						HotelPic18.Parent = Tile18Colour
						HotelPic18.Top = (Tile18Colour.Height / 2) - HotelPic18.Height
						HotelPic18.Left = 0
						HotelPic18.Tag = "HotelPic18"
					Case 19
						Dim HotelPic19 As New PictureBox
						HotelPic19.Height = housewidth * 2
						HotelPic19.Width = Tile19Colour.Width
						HotelPic19.Image = My.Resources.monopoly_hotel
						HotelPic19.BackColor = Color.Transparent
						HotelPic19.BringToFront()
						HotelPic19.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic19)
						HotelPic19.Parent = Tile19Colour
						HotelPic19.Top = (Tile19Colour.Height / 2) - HotelPic19.Height
						HotelPic19.Left = 0
						HotelPic19.Tag = "HotelPic19"
					Case 21
						Dim HotelPic21 As New PictureBox
						HotelPic21.Height = Tile21Colour.Height
						HotelPic21.Width = housewidth * 2
						HotelPic21.Image = My.Resources.monopoly_hotel
						HotelPic21.BackColor = Color.Transparent
						HotelPic21.BringToFront()
						HotelPic21.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic21)
						HotelPic21.Parent = Tile21Colour
						HotelPic21.Top = 0
						HotelPic21.Left = (Tile21Colour.Width / 2) - HotelPic21.Width
						HotelPic21.Tag = "HotelPic21"
					Case 23
						Dim HotelPic23 As New PictureBox
						HotelPic23.Height = Tile23Colour.Height
						HotelPic23.Width = housewidth * 2
						HotelPic23.Image = My.Resources.monopoly_hotel
						HotelPic23.BackColor = Color.Transparent
						HotelPic23.BringToFront()
						HotelPic23.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic23)
						HotelPic23.Parent = Tile23Colour
						HotelPic23.Top = 0
						HotelPic23.Left = (Tile23Colour.Width / 2) - HotelPic23.Width
						HotelPic23.Tag = "HotelPic23"
					Case 24
						Dim HotelPic24 As New PictureBox
						HotelPic24.Height = Tile24Colour.Height
						HotelPic24.Width = housewidth * 2
						HotelPic24.Image = My.Resources.monopoly_hotel
						HotelPic24.BackColor = Color.Transparent
						HotelPic24.BringToFront()
						HotelPic24.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic24)
						HotelPic24.Parent = Tile24Colour
						HotelPic24.Top = 0
						HotelPic24.Left = (Tile24Colour.Width / 2) - HotelPic24.Width
						HotelPic24.Tag = "HotelPic24"
					Case 26
						Dim HotelPic26 As New PictureBox
						HotelPic26.Height = Tile26Colour.Height
						HotelPic26.Width = housewidth * 2
						HotelPic26.Image = My.Resources.monopoly_hotel
						HotelPic26.BackColor = Color.Transparent
						HotelPic26.BringToFront()
						HotelPic26.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic26)
						HotelPic26.Parent = Tile26Colour
						HotelPic26.Top = 0
						HotelPic26.Left = (Tile26Colour.Width / 2) - HotelPic26.Width
						HotelPic26.Tag = "HotelPic26"
					Case 27
						Dim HotelPic27 As New PictureBox
						HotelPic27.Height = Tile27Colour.Height
						HotelPic27.Width = housewidth * 2
						HotelPic27.Image = My.Resources.monopoly_hotel
						HotelPic27.BackColor = Color.Transparent
						HotelPic27.BringToFront()
						HotelPic27.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic27)
						HotelPic27.Parent = Tile27Colour
						HotelPic27.Top = 0
						HotelPic27.Left = (Tile27Colour.Width / 2) - HotelPic27.Width
						HotelPic27.Tag = "HotelPic27"
					Case 29
						Dim HotelPic29 As New PictureBox
						HotelPic29.Height = Tile29Colour.Height
						HotelPic29.Width = housewidth * 2
						HotelPic29.Image = My.Resources.monopoly_hotel
						HotelPic29.BackColor = Color.Transparent
						HotelPic29.BringToFront()
						HotelPic29.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic29)
						HotelPic29.Parent = Tile29Colour
						HotelPic29.Top = 0
						HotelPic29.Left = (Tile29Colour.Width / 2) - HotelPic29.Width
						HotelPic29.Tag = "HotelPic29"
					Case 31
						Dim HotelPic31 As New PictureBox
						HotelPic31.Height = housewidth * 2
						HotelPic31.Width = Tile31Colour.Width
						HotelPic31.Image = My.Resources.monopoly_hotel
						HotelPic31.BackColor = Color.Transparent
						HotelPic31.BringToFront()
						HotelPic31.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic31)
						HotelPic31.Parent = Tile31Colour
						HotelPic31.Top = (Tile31Colour.Height / 2) - HotelPic31.Height
						HotelPic31.Left = 0
						HotelPic31.Tag = "HotelPic31"
					Case 32
						Dim HotelPic32 As New PictureBox
						HotelPic32.Height = housewidth * 2
						HotelPic32.Width = Tile32Colour.Width
						HotelPic32.Image = My.Resources.monopoly_hotel
						HotelPic32.BackColor = Color.Transparent
						HotelPic32.BringToFront()
						HotelPic32.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic32)
						HotelPic32.Parent = Tile32Colour
						HotelPic32.Top = (Tile32Colour.Height / 2) - HotelPic32.Height
						HotelPic32.Left = 0
						HotelPic32.Tag = "HotelPic32"
					Case 34
						Dim HotelPic34 As New PictureBox
						HotelPic34.Height = housewidth * 2
						HotelPic34.Width = Tile34Colour.Width
						HotelPic34.Image = My.Resources.monopoly_hotel
						HotelPic34.BackColor = Color.Transparent
						HotelPic34.BringToFront()
						HotelPic34.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic34)
						HotelPic34.Parent = Tile34Colour
						HotelPic34.Top = (Tile34Colour.Height / 2) - HotelPic34.Height
						HotelPic34.Left = 0
						HotelPic34.Tag = "HotelPic34"
					Case 37
						Dim HotelPic37 As New PictureBox
						HotelPic37.Height = housewidth * 2
						HotelPic37.Width = Tile37Colour.Width
						HotelPic37.Image = My.Resources.monopoly_hotel
						HotelPic37.BackColor = Color.Transparent
						HotelPic37.BringToFront()
						HotelPic37.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic37)
						HotelPic37.Parent = Tile37Colour
						HotelPic37.Top = (Tile37Colour.Height / 2) - HotelPic37.Height
						HotelPic37.Left = 0
						HotelPic37.Tag = "HotelPic37"
					Case 39
						Dim HotelPic39 As New PictureBox
						HotelPic39.Height = housewidth * 2
						HotelPic39.Width = Tile39Colour.Width
						HotelPic39.Image = My.Resources.monopoly_hotel
						HotelPic39.BackColor = Color.Transparent
						HotelPic39.BringToFront()
						HotelPic39.SizeMode = PictureBoxSizeMode.Zoom
						Me.Controls.Add(HotelPic39)
						HotelPic39.Parent = Tile39Colour
						HotelPic39.Top = (Tile39Colour.Height / 2) - HotelPic39.Height
						HotelPic39.Left = 0
						HotelPic39.Tag = "HotelPic39"
				End Select

			End If

		Next




	End Sub

	Private Sub LoadStartGame_Click(sender As Object, e As EventArgs) Handles LoadStartGame.Click
		TurnOrderPanel.Show()
		LoadStartGame.Hide()
		TurnOrderCurrentTurn = 0
		TurnOrdertitleLbl.Text = CurrentPlayers(0).Name & " Roll the Dice"
		rollable1 = True
		rollable2 = True

		If CurrentPlayers(0).Human = False Then
			If rollable1 = True Then
				Dice1Spinner.Enabled = True
				rollable1 = False
			ElseIf Dice1Spinner.Enabled = True Then
				Dice1Spinner.Enabled = False
				If rollable2 = False And Dice2Spinner.Enabled = False Then
					CalculateTurnTotalForPlayer()
				End If
			End If
			If rollable2 = True Then
				Dice2Spinner.Enabled = True
				rollable2 = False
			ElseIf Dice2Spinner.Enabled = True Then
				Dice2Spinner.Enabled = False
				If rollable1 = False And Dice1Spinner.Enabled = False Then
					CalculateTurnTotalForPlayer()
				End If

			End If

		End If

	End Sub

	Private Sub InitialisePlayerTokens()
		Select Case PlayerNumber
			Case 2
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
			Case 3
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
			Case 4
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(3).Piece = "Wheelbarrow" Then
					Player4Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(3).Piece = "Battleship" Then
					Player4Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(3).Piece = "Racecar" Then
					Player4Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(3).Piece = "Thimble" Then
					Player4Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(3).Piece = "Boot" Then
					Player4Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(3).Piece = "Dog" Then
					Player4Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
					Player4Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(3).Piece = "Iron" Then
					Player4Token.Image = My.Resources.Iron
				End If
			Case 5
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(3).Piece = "Wheelbarrow" Then
					Player4Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(3).Piece = "Battleship" Then
					Player4Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(3).Piece = "Racecar" Then
					Player4Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(3).Piece = "Thimble" Then
					Player4Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(3).Piece = "Boot" Then
					Player4Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(3).Piece = "Dog" Then
					Player4Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
					Player4Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(3).Piece = "Iron" Then
					Player4Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(4).Piece = "Wheelbarrow" Then
					Player5Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(4).Piece = "Battleship" Then
					Player5Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(4).Piece = "Racecar" Then
					Player5Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(4).Piece = "Thimble" Then
					Player5Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(4).Piece = "Boot" Then
					Player5Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(4).Piece = "Dog" Then
					Player5Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(4).Piece = "Top Hat" Then
					Player5Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(4).Piece = "Iron" Then
					Player5Token.Image = My.Resources.Iron
				End If
			Case 6
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(3).Piece = "Wheelbarrow" Then
					Player4Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(3).Piece = "Battleship" Then
					Player4Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(3).Piece = "Racecar" Then
					Player4Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(3).Piece = "Thimble" Then
					Player4Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(3).Piece = "Boot" Then
					Player4Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(3).Piece = "Dog" Then
					Player4Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
					Player4Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(3).Piece = "Iron" Then
					Player4Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(4).Piece = "Wheelbarrow" Then
					Player5Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(4).Piece = "Battleship" Then
					Player5Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(4).Piece = "Racecar" Then
					Player5Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(4).Piece = "Thimble" Then
					Player5Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(4).Piece = "Boot" Then
					Player5Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(4).Piece = "Dog" Then
					Player5Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(4).Piece = "Top Hat" Then
					Player5Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(4).Piece = "Iron" Then
					Player5Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(5).Piece = "Wheelbarrow" Then
					Player6Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(5).Piece = "Battleship" Then
					Player6Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(5).Piece = "Racecar" Then
					Player6Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(5).Piece = "Thimble" Then
					Player6Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(5).Piece = "Boot" Then
					Player6Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(5).Piece = "Dog" Then
					Player6Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(5).Piece = "Top Hat" Then
					Player6Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(5).Piece = "Iron" Then
					Player6Token.Image = My.Resources.Iron
				End If
			Case 7
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(3).Piece = "Wheelbarrow" Then
					Player4Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(3).Piece = "Battleship" Then
					Player4Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(3).Piece = "Racecar" Then
					Player4Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(3).Piece = "Thimble" Then
					Player4Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(3).Piece = "Boot" Then
					Player4Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(3).Piece = "Dog" Then
					Player4Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
					Player4Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(3).Piece = "Iron" Then
					Player4Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(4).Piece = "Wheelbarrow" Then
					Player5Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(4).Piece = "Battleship" Then
					Player5Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(4).Piece = "Racecar" Then
					Player5Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(4).Piece = "Thimble" Then
					Player5Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(4).Piece = "Boot" Then
					Player5Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(4).Piece = "Dog" Then
					Player5Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(4).Piece = "Top Hat" Then
					Player5Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(4).Piece = "Iron" Then
					Player5Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(5).Piece = "Wheelbarrow" Then
					Player6Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(5).Piece = "Battleship" Then
					Player6Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(5).Piece = "Racecar" Then
					Player6Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(5).Piece = "Thimble" Then
					Player6Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(5).Piece = "Boot" Then
					Player6Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(5).Piece = "Dog" Then
					Player6Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(5).Piece = "Top Hat" Then
					Player6Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(5).Piece = "Iron" Then
					Player6Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(6).Piece = "Wheelbarrow" Then
					Player7Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(6).Piece = "Battleship" Then
					Player7Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(6).Piece = "Racecar" Then
					Player7Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(6).Piece = "Thimble" Then
					Player7Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(6).Piece = "Boot" Then
					Player7Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(6).Piece = "Dog" Then
					Player7Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(6).Piece = "Top Hat" Then
					Player7Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(6).Piece = "Iron" Then
					Player7Token.Image = My.Resources.Iron
				End If
			Case 8
				If CurrentPlayers(0).Piece = "Wheelbarrow" Then
					Player1Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(0).Piece = "Battleship" Then
					Player1Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(0).Piece = "Racecar" Then
					Player1Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(0).Piece = "Thimble" Then
					Player1Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(0).Piece = "Boot" Then
					Player1Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(0).Piece = "Dog" Then
					Player1Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(0).Piece = "Top Hat" Then
					Player1Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(0).Piece = "Iron" Then
					Player1Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(1).Piece = "Wheelbarrow" Then
					Player2Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(1).Piece = "Battleship" Then
					Player2Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(1).Piece = "Racecar" Then
					Player2Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(1).Piece = "Thimble" Then
					Player2Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(1).Piece = "Boot" Then
					Player2Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(1).Piece = "Dog" Then
					Player2Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(1).Piece = "Top Hat" Then
					Player2Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(1).Piece = "Iron" Then
					Player2Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(2).Piece = "Wheelbarrow" Then
					Player3Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(2).Piece = "Battleship" Then
					Player3Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(2).Piece = "Racecar" Then
					Player3Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(2).Piece = "Thimble" Then
					Player3Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(2).Piece = "Boot" Then
					Player3Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(2).Piece = "Dog" Then
					Player3Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(2).Piece = "Top Hat" Then
					Player3Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(2).Piece = "Iron" Then
					Player3Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(3).Piece = "Wheelbarrow" Then
					Player4Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(3).Piece = "Battleship" Then
					Player4Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(3).Piece = "Racecar" Then
					Player4Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(3).Piece = "Thimble" Then
					Player4Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(3).Piece = "Boot" Then
					Player4Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(3).Piece = "Dog" Then
					Player4Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(3).Piece = "Top Hat" Then
					Player4Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(3).Piece = "Iron" Then
					Player4Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(4).Piece = "Wheelbarrow" Then
					Player5Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(4).Piece = "Battleship" Then
					Player5Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(4).Piece = "Racecar" Then
					Player5Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(4).Piece = "Thimble" Then
					Player5Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(4).Piece = "Boot" Then
					Player5Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(4).Piece = "Dog" Then
					Player5Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(4).Piece = "Top Hat" Then
					Player5Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(4).Piece = "Iron" Then
					Player5Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(5).Piece = "Wheelbarrow" Then
					Player6Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(5).Piece = "Battleship" Then
					Player6Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(5).Piece = "Racecar" Then
					Player6Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(5).Piece = "Thimble" Then
					Player6Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(5).Piece = "Boot" Then
					Player6Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(5).Piece = "Dog" Then
					Player6Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(5).Piece = "Top Hat" Then
					Player6Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(5).Piece = "Iron" Then
					Player6Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(6).Piece = "Wheelbarrow" Then
					Player7Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(6).Piece = "Battleship" Then
					Player7Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(6).Piece = "Racecar" Then
					Player7Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(6).Piece = "Thimble" Then
					Player7Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(6).Piece = "Boot" Then
					Player7Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(6).Piece = "Dog" Then
					Player7Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(6).Piece = "Top Hat" Then
					Player7Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(6).Piece = "Iron" Then
					Player7Token.Image = My.Resources.Iron
				End If
				If CurrentPlayers(7).Piece = "Wheelbarrow" Then
					Player8Token.Image = My.Resources.Wheelbarrow
				ElseIf CurrentPlayers(7).Piece = "Battleship" Then
					Player8Token.Image = My.Resources.Ship
				ElseIf CurrentPlayers(7).Piece = "Racecar" Then
					Player8Token.Image = My.Resources.Car
				ElseIf CurrentPlayers(7).Piece = "Thimble" Then
					Player8Token.Image = My.Resources.Thimble
				ElseIf CurrentPlayers(7).Piece = "Boot" Then
					Player8Token.Image = My.Resources.Boot
				ElseIf CurrentPlayers(7).Piece = "Dog" Then
					Player8Token.Image = My.Resources.Dog
				ElseIf CurrentPlayers(7).Piece = "Top Hat" Then
					Player8Token.Image = My.Resources.Hat
				ElseIf CurrentPlayers(7).Piece = "Iron" Then
					Player8Token.Image = My.Resources.Iron
				End If
		End Select

		Player1Token.Parent = Tile0
		Player2Token.Parent = Tile0
		Player3Token.Parent = Tile0
		Player4Token.Parent = Tile0
		Player5Token.Parent = Tile0
		Player6Token.Parent = Tile0
		Player7Token.Parent = Tile0
		Player8Token.Parent = Tile0

		Player1Token.Top = 20
		Player2Token.Top = 40
		Player3Token.Top = 60
		Player4Token.Top = 80
		Player1Token.Left = 0
		Player2Token.Left = 0
		Player3Token.Left = 0
		Player4Token.Left = 0
		Player5Token.Top = 20
		Player6Token.Top = 40
		Player7Token.Top = 60
		Player8Token.Top = 80
		Player5Token.Left = 25
		Player6Token.Left = 25
		Player7Token.Left = 25
		Player8Token.Left = 25

		Player1Token.BringToFront()
		Player2Token.BringToFront()
		Player3Token.BringToFront()
		Player4Token.BringToFront()
		Player5Token.BringToFront()
		Player6Token.BringToFront()
		Player7Token.BringToFront()
		Player8Token.BringToFront()

		Select Case PlayerNumber
			Case 2
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Hide()
				Player4Token.Hide()
				Player5Token.Hide()
				Player6Token.Hide()
				Player7Token.Hide()
				Player8Token.Hide()
			Case 3
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Hide()
				Player5Token.Hide()
				Player6Token.Hide()
				Player7Token.Hide()
				Player8Token.Hide()
			Case 4
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Show()
				Player5Token.Hide()
				Player6Token.Hide()
				Player7Token.Hide()
				Player8Token.Hide()
			Case 5
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Show()
				Player5Token.Show()
				Player6Token.Hide()
				Player7Token.Hide()
				Player8Token.Hide()
			Case 6
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Show()
				Player5Token.Show()
				Player6Token.Show()
				Player7Token.Hide()
				Player8Token.Hide()
			Case 7
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Show()
				Player5Token.Show()
				Player6Token.Show()
				Player7Token.Show()
				Player8Token.Hide()
			Case 8
				Player1Token.Show()
				Player2Token.Show()
				Player3Token.Show()
				Player4Token.Show()
				Player5Token.Show()
				Player6Token.Show()
				Player7Token.Show()
				Player8Token.Show()
		End Select

	End Sub

	Private Function RollDie()
		Dim randomNumber As Integer
		Dim rand As New Random()

		randomNumber = rand.Next(1, 7)

		Return randomNumber
	End Function

	Private Sub Die1_Click(sender As Object, e As EventArgs) Handles Die1.Click
		If rollable1 = True Then
			Dice1Spinner.Enabled = True
			rollable1 = False
		ElseIf Dice1Spinner.Enabled = True Then
			Dice1Spinner.Enabled = False
			If rollable2 = False And Dice2Spinner.Enabled = False Then
				CalculateTurnTotalForPlayer()
			End If
		End If



	End Sub

	Private Sub DiceSpinner_Tick(sender As Object, e As EventArgs) Handles Dice1Spinner.Tick
		Dim die1no As Integer
		die1no = RollDie()
		If die1no = 1 Then
			Die1.Image = My.Resources.Die1
			die1amt = 1
		ElseIf die1no = 2 Then
			Die1.Image = My.Resources.Die2
			die1amt = 2
		ElseIf die1no = 3 Then
			Die1.Image = My.Resources.Die3
			die1amt = 3
		ElseIf die1no = 4 Then
			Die1.Image = My.Resources.Die4
			die1amt = 4
		ElseIf die1no = 5 Then
			Die1.Image = My.Resources.Die5
			die1amt = 5
		ElseIf die1no = 6 Then
			Die1.Image = My.Resources.Die6
			die1amt = 6
		End If

		If CurrentPlayers(TurnOrderCurrentTurn).Human = False Then
			If rollable1 = True Then
				Dice1Spinner.Enabled = True
				rollable1 = False
			ElseIf Dice1Spinner.Enabled = True Then
				Dice1Spinner.Enabled = False
				If rollable2 = False And Dice2Spinner.Enabled = False Then
					CalculateTurnTotalForPlayer()
				End If
			End If
		End If
	End Sub

	Private Sub Dice2Spinner_Tick(sender As Object, e As EventArgs) Handles Dice2Spinner.Tick
		Dim die2no As Integer
		die2no = RollDie()
		If die2no = 1 Then
			Die2.Image = My.Resources.Die1
			die2amt = 1
		ElseIf die2no = 2 Then
			Die2.Image = My.Resources.Die2
			die2amt = 2
		ElseIf die2no = 3 Then
			Die2.Image = My.Resources.Die3
			die2amt = 3
		ElseIf die2no = 4 Then
			Die2.Image = My.Resources.Die4
			die2amt = 4
		ElseIf die2no = 5 Then
			Die2.Image = My.Resources.Die5
			die2amt = 5
		ElseIf die2no = 6 Then
			Die2.Image = My.Resources.Die6
			die2amt = 6
		End If

		If CurrentPlayers(TurnOrderCurrentTurn).Human = False Then
			If rollable2 = True Then
				Dice2Spinner.Enabled = True
				rollable2 = False
			ElseIf Dice2Spinner.Enabled = True Then
				Dice2Spinner.Enabled = False
				If rollable1 = False And Dice1Spinner.Enabled = False Then
					CalculateTurnTotalForPlayer()
				End If

			End If
		End If
	End Sub

	Private Sub Die2_Click(sender As Object, e As EventArgs) Handles Die2.Click
		If rollable2 = True Then
			Dice2Spinner.Enabled = True
			rollable2 = False
		ElseIf Dice2Spinner.Enabled = True Then
			Dice2Spinner.Enabled = False
			If rollable1 = False And Dice1Spinner.Enabled = False Then
				CalculateTurnTotalForPlayer()
			End If

		End If
	End Sub

	Private Sub CalculateTurnTotalForPlayer()
		Dim total As Integer
		total = die1amt + die2amt

		TurnorderRolls(TurnOrderCurrentTurn) = total
		RolledAmtlbl.Show()
		RolledAmtlbl.Text = CurrentPlayers(TurnOrderCurrentTurn).Name & " Rolled a " & total
		CurrentPlayers(TurnOrderCurrentTurn).TurnOrderRolledAmt = total

		If TurnOrderCurrentTurn < PlayerNumber - 1 Then
			TurnOrderCurrentTurn += 1

			TurnOrdertitleLbl.Text = CurrentPlayers(TurnOrderCurrentTurn).Name & " Roll the Dice"

			rollable1 = True
			rollable2 = True

			If CurrentPlayers(TurnOrderCurrentTurn).Human = False Then
				If rollable1 = True Then
					Dice1Spinner.Enabled = True
					rollable1 = False
				ElseIf Dice1Spinner.Enabled = True Then
					Dice1Spinner.Enabled = False
					If rollable2 = False And Dice2Spinner.Enabled = False Then
						CalculateTurnTotalForPlayer()
					End If
				End If
				If rollable2 = True Then
					Dice2Spinner.Enabled = True
					rollable2 = False
				ElseIf Dice2Spinner.Enabled = True Then
					Dice2Spinner.Enabled = False
					If rollable1 = False And Dice1Spinner.Enabled = False Then
						CalculateTurnTotalForPlayer()
					End If

				End If


			End If
		Else
			'bubble sort player array based on rolls
			Dim temp As Player
			If PlayerNumber = 2 Then
				If CurrentPlayers(0).TurnOrderRolledAmt < CurrentPlayers(1).TurnOrderRolledAmt Then
					temp = CurrentPlayers(0)
					CurrentPlayers(0) = CurrentPlayers(1)
					CurrentPlayers(1) = temp
				End If
			Else
				For I = PlayerNumber - 1 To 0 Step -1
					For J = 0 To I - 1
						If CurrentPlayers(J).TurnOrderRolledAmt < CurrentPlayers(J + 1).TurnOrderRolledAmt Then
							temp = CurrentPlayers(J)
							CurrentPlayers(J) = CurrentPlayers(J + 1)
							CurrentPlayers(J + 1) = temp
						End If
					Next

				Next
			End If



			'print out sorted order
			OrderofPlaylbl.Text = "ORDER OF TURNS:" & Chr(13)
			For i = 0 To PlayerNumber - 1
				OrderofPlaylbl.Text += i + 1 & ": " & CurrentPlayers(i).Name & " Rolled " & CurrentPlayers(i).TurnOrderRolledAmt & Chr(13)
			Next
			StartGameTurnorder.Show()
		End If


	End Sub

	Private Sub StartGameTurnorder_Click(sender As Object, e As EventArgs) Handles StartGameTurnorder.Click
		LoadingPanel.Hide()
		GamePanel.Show()
		StartGameTurnorder.Hide()
		For i = 0 To PlayerNumber - 1
			CurrentPlayers(i).CurrentPos = 0
		Next
		Turn = 0
		InitialisePlayerTokens()
		PlayTurn()

	End Sub

	Private Sub UpdateTurnDisplayer()
		Dim length As Integer
		length = CurrentPlayers(Turn).Name.Length

		If LCase(CurrentPlayers(Turn).Name(length - 1)) = "s" Then
			TurnIndicatorlbl.Text = CurrentPlayers(Turn).Name & "' Turn"
		Else
			TurnIndicatorlbl.Text = CurrentPlayers(Turn).Name & "'s Turn"
		End If
	End Sub

	Private Sub RollDiceBtn_Click(sender As Object, e As EventArgs) Handles RollDiceBtn.Click
		Randomize()
		GameDie1Target = Int((80 - 20 + 1) * Rnd() + 20)
		Randomize()
		GameDie2Target = Int((80 - 20 + 1) * Rnd() + 20)
		GameDie1Spinner.Enabled = True
		GameDie2Spinner.Enabled = True
		GameDie1Spinner.Interval = 25
		GameDie2Spinner.Interval = 25
		RollDiceBtn.Hide()
		If CurrentPlayers(Turn).InJail = True Then
			PayOutOfJail.Hide()
			GetOutofJailFreeBtn.Hide()
		End If
	End Sub

	Private Sub GameDie1Spinner_Tick(sender As Object, e As EventArgs) Handles GameDie1Spinner.Tick
		GameDie1Target -= 1
		GameDie1Spinner.Interval += 1
		If GameDie1Target > 0 Then

			Dim die1no As Integer
			die1no = RollDie()
			If die1no = 1 Then
				GameDie1.Image = My.Resources.Die1
				die1amt = 1
			ElseIf die1no = 2 Then
				GameDie1.Image = My.Resources.Die2
				die1amt = 2
			ElseIf die1no = 3 Then
				GameDie1.Image = My.Resources.Die3
				die1amt = 3
			ElseIf die1no = 4 Then
				GameDie1.Image = My.Resources.Die4
				die1amt = 4
			ElseIf die1no = 5 Then
				GameDie1.Image = My.Resources.Die5
				die1amt = 5
			ElseIf die1no = 6 Then
				GameDie1.Image = My.Resources.Die6
				die1amt = 6
			End If
		Else
			GameDie1Spinner.Enabled = False
			If GameDie2Spinner.Enabled = False Then
				RetrieveTotalGameDieAmount()
			End If
		End If
	End Sub

	Private Sub GameDie2Spinner_Tick(sender As Object, e As EventArgs) Handles GameDie2Spinner.Tick
		GameDie2Target -= 1
		GameDie2Spinner.Interval += 1
		If GameDie2Target > 0 Then

			Dim die2no As Integer
			die2no = RollDie()
			If die2no = 1 Then
				GameDie2.Image = My.Resources.Die1
				die2amt = 1
			ElseIf die2no = 2 Then
				GameDie2.Image = My.Resources.Die2
				die2amt = 2
			ElseIf die2no = 3 Then
				GameDie2.Image = My.Resources.Die3
				die2amt = 3
			ElseIf die2no = 4 Then
				GameDie2.Image = My.Resources.Die4
				die2amt = 4
			ElseIf die2no = 5 Then
				GameDie2.Image = My.Resources.Die5
				die2amt = 5
			ElseIf die2no = 6 Then
				GameDie2.Image = My.Resources.Die6
				die2amt = 6
			End If
		Else
			GameDie2Spinner.Enabled = False
			If GameDie1Spinner.Enabled = False Then
				RetrieveTotalGameDieAmount()
			End If
		End If
	End Sub

	Private Sub RetrieveTotalGameDieAmount()
		Dim total As Integer
		total = die1amt + die2amt
		rolledAmt = total

		If CurrentPlayers(Turn).InJail = True Then
			If die1amt = die2amt Then
				MakeMove(total, True)
				CurrentPlayers(Turn).TurnsInJail = 0
			Else
				NextTurnbtn.Show()

				GetOutofJailFreeBtn.Hide()
				PayOutOfJail.Hide()
			End If
		Else
			If die1amt = die2amt Then
				MakeMove(total, True)
			Else
				MakeMove(total, False)
			End If
		End If


	End Sub

	Private Sub PlayTurn()

		UpdateTurnDisplayer()
		UpdatePlayerStats()
		If CurrentPlayers(Turn).Bankrupt = False Then
			If CurrentPlayers(Turn).Human = True Then
				If CurrentPlayers(Turn).InJail = False Then
					RollDiceBtn.Parent = GameBoard
					RollDiceBtn.Show()
					NextTurnbtn.Hide()
					PayOutOfJail.Hide()
					GetOutofJailFreeBtn.Hide()
				Else
					CurrentPlayers(Turn).TurnsInJail += 1
					PayOutOfJail.Show()
					If CurrentPlayers(Turn).GetOutJailCards > 0 Then
						GetOutofJailFreeBtn.Show()
					End If
					If CurrentPlayers(Turn).TurnsInJail <> 3 Then
						RollDiceBtn.Parent = GameBoard
						RollDiceBtn.Show()
						NextTurnbtn.Hide()
					Else
						NextTurnbtn.Hide()
						PayOutOfJail.Show()
						GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Has spent 3 turns in jail and now must pay to get out"

					End If

				End If
				ManagePropBtn.Hide()
				For i = 0 To 39
					If Tiles(i).Owner = CurrentPlayers(Turn).Name Then
						ManagePropBtn.Show()
						Exit For
					End If
				Next


			Else
				NextTurnbtn.Hide()
				Randomize()
				GameDie1Target = Int((80 - 20 + 1) * Rnd() + 20)
				Randomize()
				GameDie2Target = Int((80 - 20 + 1) * Rnd() + 20)
				GameDie1Spinner.Enabled = True
				GameDie2Spinner.Enabled = True
				GameDie1Spinner.Interval = 25
				GameDie2Spinner.Interval = 25
				RollDiceBtn.Hide()

			End If

		Else
			Turn = Turn + 1
			If Turn > PlayerNumber - 1 Then
				Turn = 0
			End If
			PlayTurn()
		End If

	End Sub

	Private Sub MakeMove(Spaces As Integer, doubleroll As Boolean)

		movemax = Spaces
		If Spaces > 12 Then
			Mover.Interval = 250
		Else
			Mover.Interval = 500
			If doubleroll = True And CurrentPlayers(Turn).CurrentPos + Spaces = 30 Then
				doubleroll = False
				doublecount = 0
			End If
			If doubleroll = False Then
				GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Rolled " & Spaces
			ElseIf doubleroll = True And CurrentPlayers(Turn).InJail = True Then
				CurrentPlayers(Turn).InJail = False
				doublecount = 0
				doubleroll = False
				GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Rolled " & Spaces & ", Doubles! They get out of Jail"
			Else
				GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Rolled " & Spaces & ", Doubles! They get to roll again!"
			End If
		End If


		Mover.Start()

		If doubleroll = True Then
			doublecount += 1
			If doublecount = 3 Then
				'send player to jail
				Mover.Stop()
				JailCurrentplayer()
				NextTurnbtn.Show()

				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Rolled 3 doubles in a row and has been Jailed"


			End If
		Else
			doublecount = 0
		End If
		CounterForTurnsInGame += 1
		TurnCounterlbl.Text = "Turn: " & CounterForTurnsInGame

	End Sub

	Private Sub UpdatePlayerPositions()
		For plr = 0 To PlayerNumber - 1
			Select Case CurrentPlayers(plr).CurrentPos
				Case 0
					If plr = 0 Then
						Player1Token.Parent = Tile0
						Player1Token.Top = 20
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile0
						Player2Token.Top = 40
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile0
						Player3Token.Top = 60
						Player3Token.Left = 0
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile0
						Player4Token.Top = 80
						Player4Token.Left = 0
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile0
						Player5Token.Top = 20
						Player5Token.Left = 25
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile0
						Player6Token.Top = 40
						Player6Token.Left = 25
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile0
						Player7Token.Top = 60
						Player7Token.Left = 25
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile0
						Player8Token.Top = 80
						Player8Token.Left = 25
					End If
				Case 10
					If plr = 0 Then
						Player1Token.Parent = Tile10
						Player1Token.Top = 20
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile10
						Player2Token.Top = 40
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile10
						Player3Token.Top = 60
						Player3Token.Left = 0
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile10
						Player4Token.Top = 80
						Player4Token.Left = 0
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile10
						Player5Token.Top = 20
						Player5Token.Left = 25
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile10
						Player6Token.Top = 40
						Player6Token.Left = 25
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile10
						Player7Token.Top = 60
						Player7Token.Left = 25
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile10
						Player8Token.Top = 80
						Player8Token.Left = 25
					End If
				Case 20
					If plr = 0 Then
						Player1Token.Parent = Tile20
						Player1Token.Top = 20
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile20
						Player2Token.Top = 40
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile20
						Player3Token.Top = 60
						Player3Token.Left = 0
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile20
						Player4Token.Top = 80
						Player4Token.Left = 0
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile20
						Player5Token.Top = 20
						Player5Token.Left = 25
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile20
						Player6Token.Top = 40
						Player6Token.Left = 25
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile20
						Player7Token.Top = 60
						Player7Token.Left = 25
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile20
						Player8Token.Top = 80
						Player8Token.Left = 25
					End If
				Case 30
					If plr = 0 Then
						Player1Token.Parent = Tile30
						Player1Token.Top = 20
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile30
						Player2Token.Top = 40
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile30
						Player3Token.Top = 60
						Player3Token.Left = 0
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile30
						Player4Token.Top = 80
						Player4Token.Left = 0
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile30
						Player5Token.Top = 20
						Player5Token.Left = 25
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile30
						Player6Token.Top = 40
						Player6Token.Left = 25
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile30
						Player7Token.Top = 60
						Player7Token.Left = 25
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile30
						Player8Token.Top = 80
						Player8Token.Left = 25
					End If
				Case 1
					If plr = 0 Then
						Player1Token.Parent = Tile1
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile1
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile1
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile1
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile1
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile1
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile1
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile1
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 2
					If plr = 0 Then
						Player1Token.Parent = Tile2
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile2
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile2
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile2
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile2
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile2
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile2
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile2
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 3
					If plr = 0 Then
						Player1Token.Parent = Tile3
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile3
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile3
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile3
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile3
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile3
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile3
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile3
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 4
					If plr = 0 Then
						Player1Token.Parent = Tile4
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile4
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile4
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile4
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile4
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile4
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile4
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile4
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 5
					If plr = 0 Then
						Player1Token.Parent = Tile5
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile5
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile5
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile5
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile5
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile5
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile5
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile5
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 6
					If plr = 0 Then
						Player1Token.Parent = Tile6
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile6
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile6
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile6
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile6
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile6
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile6
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile6
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 7
					If plr = 0 Then
						Player1Token.Parent = Tile7
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile7
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile7
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile7
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile7
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile7
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile7
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile7
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 8
					If plr = 0 Then
						Player1Token.Parent = Tile8
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile8
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile8
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile8
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile8
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile8
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile8
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile8
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 9
					If plr = 0 Then
						Player1Token.Parent = Tile9
						Player1Token.Top = 50
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile9
						Player2Token.Top = 70
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile9
						Player3Token.Top = 50
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile9
						Player4Token.Top = 70
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile9
						Player5Token.Top = 50
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile9
						Player6Token.Top = 70
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile9
						Player7Token.Top = 55
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile9
						Player8Token.Top = 75
						Player8Token.Left = 60
					End If
				Case 11
					If plr = 0 Then
						Player1Token.Parent = Tile11
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile11
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile11
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile11
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile11
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile11
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile11
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile11
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 12
					If plr = 0 Then
						Player1Token.Parent = Tile12
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile12
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile12
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile12
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile12
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile12
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile12
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile12
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 13
					If plr = 0 Then
						Player1Token.Parent = Tile13
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile13
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile13
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile13
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile13
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile13
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile13
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile13
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 14
					If plr = 0 Then
						Player1Token.Parent = Tile14
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile14
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile14
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile14
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile14
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile14
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile14
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile14
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 15
					If plr = 0 Then
						Player1Token.Parent = Tile15
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile15
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile15
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile15
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile15
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile15
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile15
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile15
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 16
					If plr = 0 Then
						Player1Token.Parent = Tile16
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile16
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile16
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile16
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile16
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile16
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile16
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile16
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 17
					If plr = 0 Then
						Player1Token.Parent = Tile17
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile17
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile17
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile17
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile17
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile17
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile17
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile17
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 18
					If plr = 0 Then
						Player1Token.Parent = Tile18
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile18
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile18
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile18
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile18
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile18
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile18
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile18
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 19
					If plr = 0 Then
						Player1Token.Parent = Tile19
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile19
						Player2Token.Top = 0
						Player2Token.Left = 30
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile19
						Player3Token.Top = 0
						Player3Token.Left = 10
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile19
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile19
						Player5Token.Top = 20
						Player5Token.Left = 30
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile19
						Player6Token.Top = 20
						Player6Token.Left = 10
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile19
						Player7Token.Top = 40
						Player7Token.Left = 40
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile19
						Player8Token.Top = 40
						Player8Token.Left = 20
					End If
				Case 21
					If plr = 0 Then
						Player1Token.Parent = Tile21
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile21
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile21
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile21
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile21
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile21
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile21
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile21
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 22
					If plr = 0 Then
						Player1Token.Parent = Tile22
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile22
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile22
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile22
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile22
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile22
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile22
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile22
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 23
					If plr = 0 Then
						Player1Token.Parent = Tile23
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile23
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile23
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile23
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile23
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile23
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile23
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile23
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 24
					If plr = 0 Then
						Player1Token.Parent = Tile24
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile24
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile24
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile24
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile24
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile24
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile24
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile24
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 25
					If plr = 0 Then
						Player1Token.Parent = Tile25
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile25
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile25
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile25
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile25
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile25
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile25
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile25
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 26
					If plr = 0 Then
						Player1Token.Parent = Tile26
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile26
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile26
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile26
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile26
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile26
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile26
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile26
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 27
					If plr = 0 Then
						Player1Token.Parent = Tile27
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile27
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile27
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile27
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile27
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile27
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile27
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile27
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 28
					If plr = 0 Then
						Player1Token.Parent = Tile28
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile28
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile28
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile28
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile28
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile28
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile28
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile28
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 29
					If plr = 0 Then
						Player1Token.Parent = Tile29
						Player1Token.Top = 30
						Player1Token.Left = 0
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile29
						Player2Token.Top = 50
						Player2Token.Left = 0
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile29
						Player3Token.Top = 30
						Player3Token.Left = 20
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile29
						Player4Token.Top = 50
						Player4Token.Left = 20
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile29
						Player5Token.Top = 30
						Player5Token.Left = 40
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile29
						Player6Token.Top = 50
						Player6Token.Left = 40
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile29
						Player7Token.Top = 35
						Player7Token.Left = 60
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile29
						Player8Token.Top = 55
						Player8Token.Left = 60
					End If
				Case 31
					If plr = 0 Then
						Player1Token.Parent = Tile31
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile31
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile31
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile31
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile31
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile31
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile31
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile31
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 32
					If plr = 0 Then
						Player1Token.Parent = Tile32
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile32
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile32
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile32
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile32
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile32
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile32
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile32
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 33
					If plr = 0 Then
						Player1Token.Parent = Tile33
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile33
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile33
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile33
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile33
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile33
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile33
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile33
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 34
					If plr = 0 Then
						Player1Token.Parent = Tile34
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile34
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile34
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile34
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile34
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile34
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile34
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile34
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 35
					If plr = 0 Then
						Player1Token.Parent = Tile35
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile35
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile35
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile35
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile35
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile35
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile35
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile35
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 36
					If plr = 0 Then
						Player1Token.Parent = Tile36
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile36
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile36
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile36
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile36
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile36
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile36
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile36
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 37
					If plr = 0 Then
						Player1Token.Parent = Tile37
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile37
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile37
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile37
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile37
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile37
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile37
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile37
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 38
					If plr = 0 Then
						Player1Token.Parent = Tile38
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile38
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile38
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile38
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile38
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile38
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile38
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile38
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
				Case 39
					If plr = 0 Then
						Player1Token.Parent = Tile39
						Player1Token.Top = 0
						Player1Token.Left = 50
					ElseIf plr = 1 Then
						Player2Token.Parent = Tile39
						Player2Token.Top = 0
						Player2Token.Left = 70
					ElseIf plr = 2 Then
						Player3Token.Parent = Tile39
						Player3Token.Top = 0
						Player3Token.Left = 90
					ElseIf plr = 3 Then
						Player4Token.Parent = Tile39
						Player4Token.Top = 20
						Player4Token.Left = 50
					ElseIf plr = 4 Then
						Player5Token.Parent = Tile39
						Player5Token.Top = 20
						Player5Token.Left = 70
					ElseIf plr = 5 Then
						Player6Token.Parent = Tile39
						Player6Token.Top = 20
						Player6Token.Left = 90
					ElseIf plr = 6 Then
						Player7Token.Parent = Tile39
						Player7Token.Top = 40
						Player7Token.Left = 45
					ElseIf plr = 7 Then
						Player8Token.Parent = Tile39
						Player8Token.Top = 40
						Player8Token.Left = 75
					End If
			End Select
		Next


		Player1Token.BringToFront()
		Player2Token.BringToFront()
		Player3Token.BringToFront()
		Player4Token.BringToFront()
		Player5Token.BringToFront()
		Player6Token.BringToFront()
		Player7Token.BringToFront()
		Player8Token.BringToFront()

	End Sub

	Private Sub JailCurrentplayer()
		CurrentPlayers(Turn).InJail = True
		CurrentPlayers(Turn).CurrentPos = 10
		GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Was Sent Directly To Jail"
		UpdatePlayerPositions()
		doublecount = 0

	End Sub

	Private Sub FlashTile(Tile As Panel, inital As Boolean)
		Dim targetR As Integer
		Dim targetG As Integer
		Dim targetB As Integer

		If inital = True Then
			Flasher.Start()
			FRed = 255
			FGreen = 255
			FBlue = 0
		Else
			targetR = Me.BackColor.R
			targetG = Me.BackColor.G
			targetB = Me.BackColor.B

			If targetR = FRed And targetB = FBlue And targetG = FGreen Then
				Flasher.Stop()
				FRed = 255
				FGreen = 255
				FBlue = 0
				Tile.BackColor = Color.Transparent
			Else
				Tile.BackColor = Color.FromArgb(FRed, FGreen, FBlue)
			End If
			If targetR < FRed Then
				FRed -= 2
			Else
				FRed += 2
			End If
			If targetG < FGreen Then
				FGreen -= 2
			Else
				FGreen += 2
			End If
			If targetB < FBlue Then
				FBlue -= 2
			Else
				FBlue += 2
			End If



		End If



	End Sub

	Private Sub Flasher_Tick(sender As Object, e As EventArgs) Handles Flasher.Tick

		Select Case CurrentPlayers(Turn).CurrentPos
			Case 0
				FlashTile(Tile0, False)
			Case 1
				FlashTile(Tile1, False)
			Case 2
				FlashTile(Tile2, False)
			Case 3
				FlashTile(Tile3, False)
			Case 4
				FlashTile(Tile4, False)
			Case 5
				FlashTile(Tile5, False)
			Case 6
				FlashTile(Tile6, False)
			Case 7
				FlashTile(Tile7, False)
			Case 8
				FlashTile(Tile8, False)
			Case 9
				FlashTile(Tile9, False)
			Case 10
				FlashTile(Tile10, False)
			Case 11
				FlashTile(Tile11, False)
			Case 12
				FlashTile(Tile12, False)
			Case 13
				FlashTile(Tile13, False)
			Case 14
				FlashTile(Tile14, False)
			Case 15
				FlashTile(Tile15, False)
			Case 16
				FlashTile(Tile16, False)
			Case 17
				FlashTile(Tile17, False)
			Case 18
				FlashTile(Tile18, False)
			Case 19
				FlashTile(Tile19, False)
			Case 20
				FlashTile(Tile20, False)
			Case 21
				FlashTile(Tile21, False)
			Case 22
				FlashTile(Tile22, False)
			Case 23
				FlashTile(Tile23, False)
			Case 24
				FlashTile(Tile24, False)
			Case 25
				FlashTile(Tile25, False)
			Case 26
				FlashTile(Tile26, False)
			Case 27
				FlashTile(Tile27, False)
			Case 28
				FlashTile(Tile28, False)
			Case 29
				FlashTile(Tile29, False)
			Case 30
				FlashTile(Tile30, False)
			Case 31
				FlashTile(Tile31, False)
			Case 32
				FlashTile(Tile32, False)
			Case 33
				FlashTile(Tile33, False)
			Case 34
				FlashTile(Tile34, False)
			Case 35
				FlashTile(Tile35, False)
			Case 36
				FlashTile(Tile36, False)
			Case 37
				FlashTile(Tile37, False)
			Case 38
				FlashTile(Tile38, False)
			Case 39
				FlashTile(Tile39, False)



		End Select
	End Sub

	Private Sub Mover_Tick(sender As Object, e As EventArgs) Handles Mover.Tick
		If movemax <> 0 Then
			ClearTileBackground()
			CurrentPlayers(Turn).CurrentPos += 1
			If CurrentPlayers(Turn).CurrentPos > 39 Then
				CurrentPlayers(Turn).CurrentPos -= 40
				CurrentPlayers(Turn).Cash += 200
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Recieved £200 for Passing GO"
				UpdatePlayerStats()
			End If
			UpdatePlayerPositions()
			FlashTile(Tile0, True)
			movemax -= 1
		Else
			Mover.Stop()
			ClearTileBackground()
			If doublecount = 0 Then
				CheckWhichTilePlayerLandedOn()
				If PropertyPopup.Visible = False And cardpopuppanel.visible = False Then
					NextTurnbtn.Show()
				End If
				If CurrentPlayers(Turn).Human = False Then
					Turn = Turn + 1
					If Turn > PlayerNumber - 1 Then
						Turn = 0
					End If
					PlayTurn()
				End If
			Else
				CheckWhichTilePlayerLandedOn()
				PlayTurn()
			End If


		End If

	End Sub
	Private Sub CheckWhichTilePlayerLandedOn()
		Dim CurrentTile As Integer
		CurrentTile = CurrentPlayers(Turn).CurrentPos

		If CurrentPlayers(Turn).Human = True Then
			If Tiles(CurrentTile).IsProperty = True Then
				'Property
				If Tiles(CurrentTile).Owner = "" Then
					'Not Owned
					NextTurnbtn.Hide()
					PropertyPopup.Show()
					RailRoadPopupDisplay.Hide()
					UtilityPopUpDisplay.Hide()
					PropertyPopupDisplay.BringToFront()
					PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
					PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
					PropertyPopupDisplay.Show()
					PropertyPopupDisplay.BorderStyle = BorderStyle.FixedSingle
					PopupColour.BackColor = Tiles(CurrentTile).Colour
					PopupTitle.Text = Tiles(CurrentTile).Title
					BaseRentlbl.Text = "Rent £" & Tiles(CurrentTile).Rent.Base
					OneHouseRent.Text = "£" & Tiles(CurrentTile).Rent.OneHouse
					TwoHouseRent.Text = "£" & Tiles(CurrentTile).Rent.TwoHouse
					ThreeHouseRent.Text = "£" & Tiles(CurrentTile).Rent.ThreeHouse
					HouseFourRent.Text = "£" & Tiles(CurrentTile).Rent.FourHouse
					HotelRent.Text = "£" & Tiles(CurrentTile).Rent.Hotel
					PropMortgageValue.Text = "Mortgage Value £" & Tiles(CurrentTile).MortgageValue
					PropHouseCost.Text = "Houses Cost £" & Tiles(CurrentTile).HouseCost & " Each"
					PropHotelCost.Text = "Hotel, £" & Tiles(CurrentTile).HotelCost & " Plus 4 Houses"
					BuyPropBtn.Show()
					ClosePropPopupbtn.Show()

				Else
					'Owned
					If Tiles(CurrentTile).Owner <> CurrentPlayers(Turn).Name Then
						PayRentForProperty(CurrentTile)
					End If
				End If

			ElseIf Tiles(CurrentTile).Station = True Then
				'station
				If Tiles(CurrentTile).Owner = "" Then
					'not owned
					PropertyPopup.Show()
					NextTurnbtn.Hide()
					PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
					PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
					RailRoadPopupDisplay.Show()
					PropertyPopupDisplay.Hide()
					UtilityPopUpDisplay.Hide()
					RailRoadPopupDisplay.BringToFront()
					RailRoadPopupDisplay.BorderStyle = BorderStyle.FixedSingle
					RailRoadPopupArt.Image = My.Resources.BottomRail
					RailRoadPopupArt.SizeMode = PictureBoxSizeMode.Zoom
					RailRoadPopupTitle.Text = Tiles(CurrentTile).Title
					BuyPropBtn.Show()
					ClosePropPopupbtn.Show()
				Else
					'owned
					If Tiles(CurrentTile).Owner <> CurrentPlayers(Turn).Name Then
						PayRentForRailRoad(CurrentTile)
					End If
				End If

			ElseIf Tiles(CurrentTile).Utility = True Then
				'utility
				If Tiles(CurrentTile).Owner = "" Then
					'not owned
					PropertyPopup.Show()
					NextTurnbtn.Hide()
					PropertyPopupDisplay.Hide()
					RailRoadPopupDisplay.Hide()
					PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
					PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
					UtilityPopUpDisplay.Show()
					UtilityPopUpDisplay.BringToFront()
					UtilityPopUpDisplay.BorderStyle = BorderStyle.FixedSingle
					If CurrentTile = 12 Then
						UtilityPopupArt.Image = My.Resources.Electric
					ElseIf CurrentTile = 28 Then
						UtilityPopupArt.Image = My.Resources.WaterWorks
					End If
					UtilityPopupArt.SizeMode = PictureBoxSizeMode.Zoom
					UtilityPopupTitle.Text = Tiles(CurrentTile).Title
					BuyPropBtn.Show()
					ClosePropPopupbtn.Show()
				Else
					'owned
					If Tiles(CurrentTile).Owner <> CurrentPlayers(Turn).Name Then
						PayRentForUtility(CurrentTile)
					End If
				End If
			ElseIf Tiles(CurrentTile).Tax = True Then
				'Tax
				PayTax(CurrentTile)
				UpdatePlayerStats()

			ElseIf Tiles(CurrentTile).GoToJail = True Then
				'Go To Jail
				JailCurrentplayer()
				If doublecount > 0 Then
					doublecount = 0
					NextTurnbtn.Show()
					RollDiceBtn.Hide()
				End If
			ElseIf Tiles(CurrentTile).GO = True Then
				'GO
				CurrentPlayers(Turn).Cash += (200 * CurrentHouseRules.GOmulitplier) - 200
				If CurrentHouseRules.GOmulitplier >= 2 Then
					GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Recieved an extra £" & (200 * CurrentHouseRules.GOmulitplier) - 200 & " For Landing on GO"
				End If
				UpdatePlayerStats()

			ElseIf Tiles(CurrentTile).FreeParking = True Then
				'FreeParking
				If CurrentHouseRules.FreeParkingMoney = True Then
					CurrentPlayers(Turn).Cash += MoneyOnFreeParking
					GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Recieved £" & MoneyOnFreeParking & " For Landing on Free Parking"
					MoneyOnFreeParking = 100
					UpdatePlayerStats()

				End If
			ElseIf Tiles(CurrentTile).Chance = True Then
				DrawChanceCard()
			ElseIf Tiles(CurrentTile).ComChest = True Then
				DrawComChestCard()
			End If

		End If



	End Sub

	Private Sub DrawComChestCard()
		CardPopupPanel.BackColor = Color.DeepSkyBlue
		CardPopupPanel.borderstyle = BorderStyle.fixedsingle
		GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Landed On Community Chest And Drew A Card"
		CardPopupPanel.Show()
		CardPopupBtn.Show()
		CardPopupTitle.Text = "COMMUNITY CHEST"
		Dim carddrawn As String
		carddrawn = ComChestCards(0)
		ComChestCards.RemoveAt(0)
		ComChestCards.Add(carddrawn)
		CardPopupText.Text = carddrawn
	End Sub
	Private Sub DrawChanceCard()
		CardPopupPanel.BackColor = Color.Orange
		CardPopupPanel.borderstyle = BorderStyle.fixedsingle
		GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Landed On Chance And Drew A Chance Card"
		CardPopupPanel.Show()
		CardPopupBtn.Show()
		CardPopupTitle.Text = "CHANCE"
		Dim carddrawn As String
		carddrawn = ChanceCards(0)
		ChanceCards.RemoveAt(0)
		ChanceCards.Add(carddrawn)
		CardPopupText.Text = carddrawn
	End Sub
	Private Sub PayTax(tile As Integer)
		If CurrentPlayers(Turn).Cash - Tiles(tile).Rent.TaxAmount >= 0 Then
			CurrentPlayers(Turn).Cash -= Tiles(tile).Rent.TaxAmount
			GameActivityLogLbl.Text = CurrentPlayers(Turn).Name & " Paid £" & Tiles(tile).Rent.TaxAmount & " In " & Tiles(tile).Title
			If CurrentHouseRules.FreeParkingMoney = True And CurrentHouseRules.FreeParkingTaxCollect = True Then
				MoneyOnFreeParking += Tiles(tile).Rent.TaxAmount
			End If
		Else
			AutoMortgagePropertiesToPayDebts(Turn, Tiles(tile).Rent.TaxAmount)
			CurrentPlayers(Turn).Cash -= Tiles(tile).Rent.TaxAmount
		End If
	End Sub
	Private Sub PayRentForUtility(tileno As Integer)
		Dim utilitiesOwned As Byte
		If Tiles(tileno).Mortgaged = True Then
			GameActivityLogLbl.Text += Chr(13) & Tiles(tileno).Title & " Is Mortgaged so no rent is paid"
			Exit Sub
		End If
		If Tiles(12).Owner = Tiles(CurrentPlayers(Turn).CurrentPos).Owner Then
			utilitiesOwned += 1
		End If
		If Tiles(28).Owner = Tiles(CurrentPlayers(Turn).CurrentPos).Owner Then
			utilitiesOwned += 1
		End If

		If utilitiesOwned = 1 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(tileno).Owner Then
					PayMoneyToPlayer(Tiles(tileno).Rent.OneUtility * rolledAmt, Turn, i)
				End If
			Next
		ElseIf utilitiesOwned = 2 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(tileno).Owner Then
					PayMoneyToPlayer(Tiles(tileno).Rent.TwoUtility * rolledAmt, Turn, i)
				End If
			Next
		End If
	End Sub
	Private Sub PayRentForRailRoad(Tileno As Integer)
		Dim stationsowned As Integer
		If Tiles(Tileno).Mortgaged = True Then
			GameActivityLogLbl.Text += Chr(13) & Tiles(Tileno).Title & " Is Mortgaged so no rent is paid"
			Exit Sub
		End If
		If Tiles(5).Owner = Tiles(Tileno).Owner Then
			stationsowned += 1
		End If
		If Tiles(15).Owner = Tiles(Tileno).Owner Then
			stationsowned += 1
		End If
		If Tiles(25).Owner = Tiles(Tileno).Owner Then
			stationsowned += 1
		End If
		If Tiles(35).Owner = Tiles(Tileno).Owner Then
			stationsowned += 1
		End If

		If stationsowned = 1 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(Tileno).Owner Then
					PayMoneyToPlayer(Tiles(Tileno).Rent.OneStation, Turn, i)
				End If
			Next
		ElseIf stationsowned = 2 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(Tileno).Owner Then
					PayMoneyToPlayer(Tiles(Tileno).Rent.TwoStation, Turn, i)
				End If
			Next
		ElseIf stationsowned = 3 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(Tileno).Owner Then
					PayMoneyToPlayer(Tiles(Tileno).Rent.ThreeStation, Turn, i)
				End If
			Next
		ElseIf stationsowned = 4 Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(Tileno).Owner Then
					PayMoneyToPlayer(Tiles(Tileno).Rent.FourStation, Turn, i)
				End If
			Next
		End If
	End Sub
	Private Sub ClearTileBackground()
		Select Case CurrentPlayers(Turn).CurrentPos
			Case 0
				Tile0.BackColor = Color.Transparent
			Case 1
				Tile1.BackColor = Color.Transparent
			Case 2
				Tile2.BackColor = Color.Transparent
			Case 3
				Tile3.BackColor = Color.Transparent
			Case 4
				Tile4.BackColor = Color.Transparent
			Case 5
				Tile5.BackColor = Color.Transparent
			Case 6
				Tile6.BackColor = Color.Transparent
			Case 7
				Tile7.BackColor = Color.Transparent
			Case 8
				Tile8.BackColor = Color.Transparent
			Case 9
				Tile9.BackColor = Color.Transparent
			Case 10
				Tile10.BackColor = Color.Transparent
			Case 11
				Tile11.BackColor = Color.Transparent
			Case 12
				Tile12.BackColor = Color.Transparent
			Case 13
				Tile13.BackColor = Color.Transparent
			Case 14
				Tile14.BackColor = Color.Transparent
			Case 15
				Tile15.BackColor = Color.Transparent
			Case 16
				Tile16.BackColor = Color.Transparent
			Case 17
				Tile17.BackColor = Color.Transparent
			Case 18
				Tile18.BackColor = Color.Transparent
			Case 19
				Tile19.BackColor = Color.Transparent
			Case 20
				Tile20.BackColor = Color.Transparent
			Case 21
				Tile21.BackColor = Color.Transparent
			Case 22
				Tile22.BackColor = Color.Transparent
			Case 23
				Tile23.BackColor = Color.Transparent
			Case 24
				Tile24.BackColor = Color.Transparent
			Case 25
				Tile25.BackColor = Color.Transparent
			Case 26
				Tile26.BackColor = Color.Transparent
			Case 27
				Tile27.BackColor = Color.Transparent
			Case 28
				Tile28.BackColor = Color.Transparent
			Case 29
				Tile29.BackColor = Color.Transparent
			Case 30
				Tile30.BackColor = Color.Transparent
			Case 31
				Tile31.BackColor = Color.Transparent
			Case 32
				Tile32.BackColor = Color.Transparent
			Case 33
				Tile33.BackColor = Color.Transparent
			Case 34
				Tile34.BackColor = Color.Transparent
			Case 35
				Tile35.BackColor = Color.Transparent
			Case 36
				Tile36.BackColor = Color.Transparent
			Case 37
				Tile37.BackColor = Color.Transparent
			Case 38
				Tile38.BackColor = Color.Transparent
			Case 39
				Tile39.BackColor = Color.Transparent
		End Select
	End Sub

	Private Sub PayRentForProperty(propertyNo As Integer)
		Dim housenumber As Integer
		Dim hotel As Boolean

		Dim SetTotal As Byte
		Dim OwnedinSet As Byte
		Dim searchColour As System.Drawing.Color

		If Tiles(propertyNo).Mortgaged = True Then
			GameActivityLogLbl.Text += Chr(13) & Tiles(propertyNo).Title & " Is Mortgaged so no rent is paid"
			Exit Sub
		End If

		housenumber = Tiles(propertyNo).HouseNo
		hotel = Tiles(propertyNo).Hotel

		If housenumber = 0 And hotel = False Then
			'check set
			searchColour = Tiles(propertyNo).Colour
			For i = 0 To 39
				If Tiles(i).Colour = searchColour Then
					SetTotal += 1
					If Tiles(i).Owner = Tiles(propertyNo).Owner Then
						OwnedinSet += 1
					End If
				End If
			Next
			If SetTotal = OwnedinSet Then
				For i = 0 To PlayerNumber - 1
					If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
						PayMoneyToPlayer(Tiles(propertyNo).Rent.Base * 2, Turn, i)
					End If
				Next
			Else
				For i = 0 To PlayerNumber - 1
					If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
						PayMoneyToPlayer(Tiles(propertyNo).Rent.Base, Turn, i)
					End If
				Next
			End If
		End If

		If hotel = True Then
			For i = 0 To PlayerNumber - 1
				If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
					PayMoneyToPlayer(Tiles(propertyNo).Rent.Hotel, Turn, i)
				End If
			Next
		Else
			Select Case housenumber
				Case 1
					For i = 0 To PlayerNumber - 1
						If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
							PayMoneyToPlayer(Tiles(propertyNo).Rent.OneHouse, Turn, i)
						End If
					Next
				Case 2
					For i = 0 To PlayerNumber - 1
						If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
							PayMoneyToPlayer(Tiles(propertyNo).Rent.TwoHouse, Turn, i)
						End If
					Next
				Case 3
					For i = 0 To PlayerNumber - 1
						If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
							PayMoneyToPlayer(Tiles(propertyNo).Rent.ThreeHouse, Turn, i)
						End If
					Next
				Case 4
					For i = 0 To PlayerNumber - 1
						If CurrentPlayers(i).Name = Tiles(propertyNo).Owner Then
							PayMoneyToPlayer(Tiles(propertyNo).Rent.FourHouse, Turn, i)
						End If
					Next
			End Select
		End If

	End Sub

	Private Sub Nxtturnbtn_Click(sender As Object, e As EventArgs) Handles NextTurnbtn.Click
		Turn = Turn + 1
		If Turn > PlayerNumber - 1 Then
			Turn = 0
		End If
		PlayTurn()
	End Sub

	Private Sub BuyPropBtn_Click(sender As Object, e As EventArgs) Handles BuyPropBtn.Click

		If ((CurrentPlayers(Turn).Cash) - (Tiles(CurrentPlayers(Turn).CurrentPos).Cost)) >= 0 Then

			GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Bought " & Tiles(CurrentPlayers(Turn).CurrentPos).Title & " For £" & Tiles(CurrentPlayers(Turn).CurrentPos).Cost

			Tiles(CurrentPlayers(Turn).CurrentPos).Owner = CurrentPlayers(Turn).Name 'Change Owner
			CurrentPlayers(Turn).Cash -= Tiles(CurrentPlayers(Turn).CurrentPos).Cost 'Subtract cost
			Tiles(CurrentPlayers(Turn).CurrentPos).Cost = CurrentPlayers(Turn).Name	'Display Owner in cost lbl			solve "£" later??


			UpdatePlayerStats()
			LoadGameBoard(False)
			ClosePropPopupbtn.Hide()
			BuyPropBtn.Hide()
			PropertyPopup.Hide()
			StartAuctionBtn.Hide()

			If doublecount = 0 Then
				NextTurnbtn.Show()

			End If
		Else
			MsgBox("You Cannot Afford This!", MsgBoxStyle.Exclamation, "Insufficient Funds")
		End If


	End Sub

	Private Sub ClosePropPopupbtn_Click(sender As Object, e As EventArgs) Handles ClosePropPopupbtn.Click
		ClosePropPopupbtn.Hide()
		BuyPropBtn.Hide()
		PropertyPopup.Hide()
		StartAuctionBtn.Hide()

		If doublecount = 0 Then
			NextTurnbtn.Show()

		End If

		If PropertyManagerPanel.Visible = True Then
			NextTurnbtn.Hide()
		End If

	End Sub

	Private Sub PayMoneyToPlayer(AmountToPay As Integer, Payer As Integer, Payee As Integer)
		If CurrentPlayers(Payee).InJail = True Then
			If CurrentHouseRules.JailRent = True Then
				If CurrentPlayers(Payer).Cash - AmountToPay >= 0 Then
					CurrentPlayers(Payer).Cash -= AmountToPay
					CurrentPlayers(Payee).Cash += AmountToPay
					GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Payer).Name & " Paid £" & AmountToPay & " In Rent To " & CurrentPlayers(Payee).Name
				Else
					'cant afford it
					AutoMortgagePropertiesToPayDebts(Payer, Abs(CurrentPlayers(Payer).Cash - AmountToPay))
					CurrentPlayers(Payer).Cash -= AmountToPay
					CurrentPlayers(Payee).Cash += AmountToPay
				End If

				UpdatePlayerStats()
			Else
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Payee).Name & " is in Jail Therefore " & CurrentPlayers(Payer).Name & " Does not pay rent"
			End If
		Else
			If CurrentPlayers(Payer).Cash - AmountToPay >= 0 Then
				CurrentPlayers(Payer).Cash -= AmountToPay
				CurrentPlayers(Payee).Cash += AmountToPay
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Payer).Name & " Paid £" & AmountToPay & " In Rent To " & CurrentPlayers(Payee).Name
			Else
				'cant afford it
				AutoMortgagePropertiesToPayDebts(Payer, Abs(CurrentPlayers(Payer).Cash - AmountToPay))
				CurrentPlayers(Payer).Cash -= AmountToPay
				CurrentPlayers(Payee).Cash += AmountToPay
			End If

			UpdatePlayerStats()
		End If


	End Sub

	Private Sub AutoMortgagePropertiesToPayDebts(Player As Integer, AmountNeeded As Integer)
		'Run through Properties and mortgage some to pay debts algorithm

		AmountNeeded -= CurrentPlayers(Player).Cash

		'Sell Hotels
		For i = 0 To 39
			If Tiles(i).Owner = CurrentPlayers(Player).Name Then
				If Tiles(i).Hotel = True Then
					Tiles(i).Hotel = False
					AmountNeeded -= (Tiles(i).HotelCost / 2)
					CurrentPlayers(Player).Cash += (Tiles(i).HotelCost / 2)
					Tiles(i).HouseNo = 4
					LoadGameBoard(False)
					GameActivityLogLbl.Text += Chr(13) & "The Hotel on " & Tiles(i).Title & "Was sold to pay Debts"
				End If
			End If
			If AmountNeeded <= 0 Then
				Exit For
			End If
		Next

		'Sell Houses
		If AmountNeeded > 0 Then
			For i = 0 To 39
				If Tiles(i).Owner = CurrentPlayers(Player).Name Then
					If Tiles(i).HouseNo > 0 Then
						Tiles(i).HouseNo -= 1
						LoadGameBoard(False)
						CurrentPlayers(Player).Cash += (Tiles(i).HouseCost / 2)
						AmountNeeded -= (Tiles(i).HouseCost / 2)
						GameActivityLogLbl.Text += Chr(13) & "A house on " & Tiles(i).Title & "Was sold to pay Debts"
					End If
				End If
				If AmountNeeded <= 0 Then
					Exit For
				End If
			Next
		End If

		'Mortgage Properties
		If AmountNeeded > 0 Then
			For i = 0 To 39
				If Tiles(i).Owner = CurrentPlayers(Player).Name Then
					If Tiles(i).Mortgaged = False Then
						Tiles(i).Mortgaged = True
						LoadGameBoard(False)
						AmountNeeded -= Tiles(i).MortgageValue
						CurrentPlayers(Player).Cash += Tiles(i).MortgageValue
						GameActivityLogLbl.Text += Chr(13) & Tiles(i).Title & " Was Automatically Mortgaged to Pay Debts"
					End If
				End If
				If AmountNeeded <= 0 Then
					Exit For
				End If
			Next
		End If


		If AmountNeeded > 0 Then
			MsgBox(CurrentPlayers(Player).Name & " Is Bankrupt")
			PlayerBankrupted(Player)
		End If

	End Sub

	Private Sub ShuffleChanceCards()
		ChanceCards.Sort(New Randomizer(Of String)())
	End Sub

	Private Sub CardPopupBtn_Click(sender As Object, e As EventArgs) Handles CardPopupBtn.Click
		CardPopupPanel.Hide()
		CardPopupBtn.Hide()
		If CardPopupTitle.Text = "CHANCE" Then
			InitiateEffectOfChanceCard(CardPopupText.Text)
			If doublecount = 0 And Mover.Enabled = False Then
				NextTurnbtn.Show()
			End If
		Else
			InitiateEffectOfComChestCard(CardPopupText.Text)
			If doublecount = 0 And Mover.Enabled = False Then
				NextTurnbtn.Show()
			End If
		End If

	End Sub

	Private Sub InitiateEffectOfComChestCard(CardText As String)
		Select Case CardText
			Case "Income Tax refund Collect £20"
				CurrentPlayers(Turn).Cash += 20
				UpdatePlayerStats()

			Case "From Sale of Stock you get £50"
				CurrentPlayers(Turn).Cash += 50
				UpdatePlayerStats()
			Case "It is Your Birthday Collect $10 from each Player"
				For i = 0 To PlayerNumber - 1
					If i = Turn Then
						Continue For
					Else
						If CurrentPlayers(i).Cash - 10 >= 0 Then
							CurrentPlayers(i).Cash -= 10
							CurrentPlayers(Turn).Cash += 10
						Else
							AutoMortgagePropertiesToPayDebts(i, 10)
							CurrentPlayers(i).Cash -= 10
							CurrentPlayers(Turn).Cash += 10
						End If
					End If
				Next
				UpdatePlayerStats()

			Case "Receive Interest on Shares £25"
				CurrentPlayers(Turn).Cash += 25
				UpdatePlayerStats()
			Case "Get out of Jail Free"
				CurrentPlayers(Turn).GetOutJailCards += 1
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Got a Get out of jail free card"
			Case "Advance to 'Go'"
				NextTurnbtn.Hide()
				MakeMove(40 - CurrentPlayers(Turn).CurrentPos, False)


			Case "Pay Hospital £100"
				If CurrentPlayers(Turn).Cash - 100 >= 0 Then
					CurrentPlayers(Turn).Cash -= 100
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 100
					End If
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 100)
					CurrentPlayers(Turn).Cash -= 100
				End If
				UpdatePlayerStats()

			Case "You have Won Second Prize in a Beauty Contest Collect £10"
				CurrentPlayers(Turn).Cash += 10
				UpdatePlayerStats()

			Case "Bank Error in your Favour Collect £200"
				CurrentPlayers(Turn).Cash += 200
				UpdatePlayerStats()
			Case "You Inherit £100"
				CurrentPlayers(Turn).Cash += 100
				UpdatePlayerStats()
			Case "Go to Jail. Move Directly to Jail. Do not Pass 'Go'. Do not Collect £200"
				JailCurrentplayer()

			Case "Pay your Insurance Premium £50"
				If CurrentPlayers(Turn).Cash - 50 >= 0 Then
					CurrentPlayers(Turn).Cash -= 50
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 50
					End If
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 50)
					CurrentPlayers(Turn).Cash -= 50
				End If
				UpdatePlayerStats()
			Case "Doctor's Fee Pay £50"
				If CurrentPlayers(Turn).Cash - 50 >= 0 Then
					CurrentPlayers(Turn).Cash -= 50
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 50
					End If
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 50)
					CurrentPlayers(Turn).Cash -= 50
				End If
				UpdatePlayerStats()
			Case "Go Back to Old Kent Road"
				CurrentPlayers(Turn).CurrentPos = 1
				UpdatePlayerPositions()
				CheckWhichTilePlayerLandedOn()
			Case "Annuity Matures Collect £100"
				CurrentPlayers(Turn).Cash += 100
				UpdatePlayerStats()

		End Select
	End Sub
	Private Sub InitiateEffectOfChanceCard(CardText As String)
		Select Case CardText
			Case "Advance to Mayfair"
				NextTurnbtn.Hide()
				MakeMove(39 - CurrentPlayers(Turn).CurrentPos, False)

			Case "Advance to Go"
				NextTurnbtn.Hide()
				MakeMove(40 - CurrentPlayers(Turn).CurrentPos, False)

			Case "You are Assessed for Street Repairs £40 per House £115 per Hotel"

				Dim housescount As Integer
				Dim hotelscount As Integer
				For prop = 0 To 39
					If Tiles(prop).Owner = CurrentPlayers(Turn).Name Then
						housescount += Tiles(prop).HouseNo
						If Tiles(prop).Hotel = True Then
							hotelscount += 1
						End If
					End If
				Next
				Dim amountToPay As Integer
				amountToPay = (40 * housescount) + 115 * hotelscount

				If amountToPay > 0 Then
					If CurrentPlayers(Turn).Cash - amountToPay >= 0 Then
						CurrentPlayers(Turn).Cash -= amountToPay
						If CurrentHouseRules.freeparkingtaxcollect = True Then
							MoneyOnFreeParking += amounttopay
						End If
						GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £" & amountToPay & " for their " & housescount & " Houses and " & hotelscount & " Hotels"
					Else
						AutoMortgagePropertiesToPayDebts(Turn, amountToPay)
						CurrentPlayers(Turn).Cash -= amountToPay
						If CurrentHouseRules.freeparkingtaxcollect = True Then
							MoneyOnFreeParking += amounttopay
						End If
						GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £" & amountToPay & " for their " & housescount & " Houses and " & hotelscount & " Hotels"
					End If
				End If

			Case "Go to Jail. Move Directly to Jail. Do not pass 'Go' Do not Collect £200"
				JailCurrentplayer()
			Case "Bank pays you Dividend of £50"
				CurrentPlayers(Turn).Cash += 50
				UpdatePlayerStats()

			Case "Go back 3 Spaces"
				CurrentPlayers(Turn).CurrentPos -= 3
				UpdatePlayerPositions()
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Was Sent back 3 spaces to " & Tiles(CurrentPlayers(Turn).CurrentPos).Title
				CheckWhichTilePlayerLandedOn()
			Case "Pay School Fees of £150"
				If CurrentPlayers(Turn).Cash - 150 >= 0 Then
					CurrentPlayers(Turn).Cash -= 150
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 150
					End If
					UpdatePlayerStats()
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 150)
					CurrentPlayers(Turn).Cash -= 150
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 150
					End If
					UpdatePlayerStats()
				End If
			Case "Make General Repairs on all of Your Houses " & Chr(13) & "For each House pay £25" & Chr(13) & "For each Hotel pay £100"
				Dim housescount As Integer
				Dim hotelscount As Integer
				For prop = 0 To 39
					If Tiles(prop).Owner = CurrentPlayers(Turn).Name Then
						housescount += Tiles(prop).HouseNo
						If Tiles(prop).Hotel = True Then
							hotelscount += 1
						End If
					End If
				Next
				Dim amountToPay As Integer
				amountToPay = (25 * housescount) + 100 * hotelscount

				If amountToPay > 0 Then
					If CurrentPlayers(Turn).Cash - amountToPay >= 0 Then
						CurrentPlayers(Turn).Cash -= amountToPay
						If CurrentHouseRules.freeparkingtaxcollect = True Then
							MoneyOnFreeParking += amounttopay
						End If
						GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £" & amountToPay & " for their " & housescount & " Houses and " & hotelscount & " Hotels"
					Else
						AutoMortgagePropertiesToPayDebts(Turn, amountToPay)
						CurrentPlayers(Turn).Cash -= amountToPay
						If CurrentHouseRules.freeparkingtaxcollect = True Then
							MoneyOnFreeParking += amounttopay
						End If
						GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £" & amountToPay & " for their " & housescount & " Houses and " & hotelscount & " Hotels"
					End If
				End If

			Case "Speeding Fine £15"
				If CurrentPlayers(Turn).Cash - 15 >= 0 Then
					CurrentPlayers(Turn).Cash -= 15
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 15
					End If
					UpdatePlayerStats()
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 15)
					CurrentPlayers(Turn).Cash -= 15
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 15
					End If
					UpdatePlayerStats()
				End If
			Case "You have won a Crossword Competition Collect £100"
				CurrentPlayers(Turn).Cash += 100
				UpdatePlayerPositions()

			Case "Your Building and Loan Matures Collect £150"
				CurrentPlayers(Turn).Cash += 150
				UpdatePlayerStats()

			Case "Get out of Jail Free"
				CurrentPlayers(Turn).GetOutJailCards += 1
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Got a Get out of jail free card"
			Case "Avance to Trafalgar Square If you Pass 'Go' Collect £200"
				NextTurnbtn.Hide()
				If CurrentPlayers(Turn).CurrentPos = 36 Then
					MakeMove(28, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 7 Then
					MakeMove(17, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 22 Then
					MakeMove(2, False)
				End If

			Case "Take a Trip to Marylebone Station and if you Pass 'Go' Collect £200"
				NextTurnbtn.Hide()
				If CurrentPlayers(Turn).CurrentPos = 36 Then
					MakeMove(19, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 7 Then
					MakeMove(8, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 22 Then
					MakeMove(33, False)
				End If

			Case "Advance to Pall Mall If you Pass 'Go' Collect £200"
				NextTurnbtn.Hide()
				If CurrentPlayers(Turn).CurrentPos = 36 Then
					MakeMove(15, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 7 Then
					MakeMove(4, False)
				ElseIf CurrentPlayers(Turn).CurrentPos = 22 Then
					MakeMove(29, False)
				End If


			Case "'Drunk in Charge' Fine £20"
				If CurrentPlayers(Turn).Cash - 20 >= 0 Then
					CurrentPlayers(Turn).Cash -= 20
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 20
					End If
					UpdatePlayerStats()
				Else
					AutoMortgagePropertiesToPayDebts(Turn, 20)
					CurrentPlayers(Turn).Cash -= 20
					If CurrentHouseRules.freeparkingtaxcollect = True Then
						MoneyOnFreeParking += 20
					End If
					UpdatePlayerStats()
				End If

		End Select
	End Sub

	Private Sub GetofJailFreeBtn_Click(sender As Object, e As EventArgs) Handles GetOutofJailFreeBtn.Click
		CurrentPlayers(Turn).InJail = False
		CurrentPlayers(Turn).GetOutJailCards -= 1
		GetOutofJailFreeBtn.hide()
		PayOutOfJail.hide()
		GameActivityLogLbl.text += Chr(13) & CurrentPlayers(turn).name & " Used a 'Get Out Of Jail Free' Card"
		CurrentPlayers(turn).turnsinjail = 0
	End Sub

	Private Sub PayOutOfJail_Click(sender As Object, e As EventArgs) Handles PayOutOfJail.Click
		PayOutOfJail.hide()
		GetOutofJailFreeBtn.hide()
		If CurrentPlayers(Turn).Cash - 50 >= 0 Then
			CurrentPlayers(Turn).InJail = False
			CurrentPlayers(Turn).Cash -= 50
			If CurrentHouseRules.freeparkingtaxcollect = True Then
				MoneyOnFreeParking += 50
			End If
			UpdatePlayerStats()

		Else
			AutoMortgagePropertiesToPayDebts(Turn, 50)
			CurrentPlayers(Turn).InJail = False
			CurrentPlayers(Turn).Cash -= 50
			If CurrentHouseRules.freeparkingtaxcollect = True Then
				MoneyOnFreeParking += 50
			End If
			UpdatePlayerStats()
		End If

		If GameActivityLogLbl.Text.Contains("Has spent 3 turns in jail and now must pay to get out") Then
			RollDiceBtn.Show()
		End If
		CurrentPlayers(Turn).TurnsInJail = 0
		GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £50 to get out of Jail"
	End Sub

	Private Sub ManagePropBtn_Click(sender As Object, e As EventArgs) Handles ManagePropBtn.Click
		NextTurnbtn.Hide()
		ManagePropBtn.Hide()
		PopulatePropertyListBox()
		PropertyManagerPanel.Show()
		PropertyManagerPanel.BringToFront()
	End Sub

	Private Sub PopulatePropertyListBox()
		PropList.BeginUpdate()
		PropList.Items.Clear()
		For i = 0 To 39
			If Tiles(i).Owner = CurrentPlayers(Turn).Name Then
				Dim LI As New ListViewItem

				LI = PropList.Items.Add(Tiles(i).Title)
				If Tiles(i).IsProperty = True Then
					LI.SubItems.Add("Property")
				ElseIf Tiles(i).Station = True Then
					LI.SubItems.Add("Station")
				ElseIf Tiles(i).Utility = True Then
					LI.SubItems.Add("Utility")
				End If

				If CheckIfPlayerOwnsFullSet(Turn, i) = True Then
					LI.SubItems.Add("Yes")
				Else
					LI.SubItems.Add("No")
				End If


				LI.SubItems.Add(Tiles(i).HouseNo)

				If Tiles(i).Hotel = True Then
					LI.SubItems.Add("Yes")
				Else
					LI.SubItems.Add("No")
				End If

				If Tiles(i).Mortgaged = True Then
					LI.SubItems.Add("Yes")
				Else
					LI.SubItems.Add("No")
				End If

				LI.SubItems.Add("£" & Tiles(i).MortgageValue)
				LI.BackColor = Tiles(i).Colour

			End If
		Next

		PropList.Update()
		PropList.EndUpdate()

	End Sub

	Private Function CheckIfPlayerOwnsFullSet(Player As Integer, CheckTile As Integer) As Boolean

		If Tiles(CheckTile).Utility = True Then
			If Tiles(12).Owner = CurrentPlayers(Player).Name And Tiles(28).Owner = CurrentPlayers(Player).Name Then
				Return True
			End If
		ElseIf Tiles(CheckTile).Station = True Then
			If Tiles(5).Owner = CurrentPlayers(Player).Name And Tiles(15).Owner = CurrentPlayers(Player).Name And Tiles(25).Owner = CurrentPlayers(Player).Name And Tiles(35).Owner = CurrentPlayers(Player).Name Then
				Return True
			End If
		Else
			Dim TotalInSet As Integer
			Dim TotalOwned As Integer
			Dim SearchColour As System.Drawing.Color
			SearchColour = Tiles(CheckTile).Colour
			For tile = 0 To 39
				If Tiles(tile).Colour = SearchColour Then
					TotalInSet += 1
					If Tiles(tile).Owner = CurrentPlayers(Player).Name Then
						TotalOwned += 1
					End If
				End If
			Next
			If TotalInSet = TotalOwned Then
				Return True
			End If
		End If
		Return False
	End Function

	Private Sub ManageCloseBtn_Click(sender As Object, e As EventArgs) Handles ManageCloseBtn.Click
		PropertyManagerPanel.Hide()
		ManagePropBtn.Show()
		If Mover.Enabled = False And GameDie1Spinner.Enabled = False And GameDie2Spinner.Enabled = False And RollDiceBtn.Visible = False Then
			NextTurnbtn.Show()
		End If


	End Sub

	Private Sub PropItemSelected(sender As Object, e As EventArgs) Handles PropList.SelectedIndexChanged
		Dim itemName As String

		If PropList.SelectedItems.Count > 0 Then
			itemName = PropList.SelectedItems.Item(0).Text

			For i = 0 To 39
				If itemName = Tiles(i).Title Then
					If Tiles(i).Mortgaged = False And Tiles(i).HouseNo = 0 And Tiles(i).Hotel = False Then
						ManageMortgagebtn.Show()
						ManageUnMortgageBtn.Hide()
					ElseIf Tiles(i).Mortgaged = True Then
						ManageUnMortgageBtn.Show()
						ManageMortgagebtn.Hide()
					End If
					If CheckIfPlayerOwnsFullSet(Turn, i) = True And Tiles(i).Hotel = False Then
						ManageBuyHouseBtn.Show()
					Else
						ManageBuyHouseBtn.Hide()
					End If
					If Tiles(i).HouseNo > 0 Then
						ManageSellHousebtn.Show()
						ManageBuyHotelBtn.Hide()

						If Tiles(i).HouseNo = 4 Then
							ManageBuyHouseBtn.Hide()
							ManageBuyHotelBtn.Show()
						Else
							ManageBuyHotelBtn.Hide()
						End If
					Else
						ManageSellHousebtn.Hide()
					End If
					If Tiles(i).Hotel = True Then
						ManageBuyHotelBtn.Hide()
						ManageSellHotelBtn.Show()
					Else
						ManageSellHotelBtn.Hide()
					End If


				End If
			Next
		End If
	End Sub

	Private Sub ManageMortgagebtn_Click(sender As Object, e As EventArgs) Handles ManageMortgagebtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next




			If Tiles(selected).Mortgaged = True Then
				MsgBox(Tiles(selected).Title & " Is Already Mortgaged")

			Else
				Tiles(selected).Mortgaged = True
				LoadGameBoard(False)
				CurrentPlayers(Turn).Cash += Tiles(selected).MortgageValue
				GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Mortgaged " & Tiles(selected).Title & " For £" & Tiles(selected).MortgageValue
				UpdatePlayerStats()
				Dim itemName As String

				If PropList.SelectedItems.Count > 0 Then
					itemName = PropList.SelectedItems.Item(0).Text

					For i = 0 To 39
						If itemName = Tiles(i).Title Then
							If Tiles(i).Mortgaged = False Then
								ManageMortgagebtn.Show()
								ManageUnMortgageBtn.Hide()
							Else
								ManageUnMortgageBtn.Show()
								ManageMortgagebtn.Hide()
							End If
							If CheckIfPlayerOwnsFullSet(Turn, i) = True Then
								ManageBuyHouseBtn.Show()
							Else
								ManageBuyHouseBtn.Hide()
							End If
							If Tiles(i).HouseNo > 0 Then
								ManageSellHousebtn.Show()
								ManageBuyHotelBtn.Hide()

								If Tiles(i).HouseNo = 4 Then
									ManageBuyHouseBtn.Hide()
									ManageBuyHotelBtn.Show()
								Else
									ManageBuyHotelBtn.Hide()
								End If
							End If


						End If
					Next
				End If
			End If
		End If

			PopulatePropertyListBox()
			

	End Sub

	Private Sub ManageViewPropertyBtn_Click(sender As Object, e As EventArgs) Handles ManageViewPropertyBtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).IsProperty = True Then
				'Property
				NextTurnbtn.Hide()
				PropertyPopup.Show()
				RailRoadPopupDisplay.Hide()
				UtilityPopUpDisplay.Hide()
				PropertyPopupDisplay.BringToFront()
				PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
				PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
				PropertyPopupDisplay.Show()
				PropertyPopupDisplay.BorderStyle = BorderStyle.FixedSingle
				PopupColour.BackColor = Tiles(selected).Colour
				PopupTitle.Text = Tiles(selected).Title
				BaseRentlbl.Text = "Rent £" & Tiles(selected).Rent.Base
				OneHouseRent.Text = "£" & Tiles(selected).Rent.OneHouse
				TwoHouseRent.Text = "£" & Tiles(selected).Rent.TwoHouse
				ThreeHouseRent.Text = "£" & Tiles(selected).Rent.ThreeHouse
				HouseFourRent.Text = "£" & Tiles(selected).Rent.FourHouse
				HotelRent.Text = "£" & Tiles(selected).Rent.Hotel
				PropMortgageValue.Text = "Mortgage Value £" & Tiles(selected).MortgageValue
				PropHouseCost.Text = "Houses Cost £" & Tiles(selected).HouseCost & " Each"
				PropHotelCost.Text = "Hotel, £" & Tiles(selected).HotelCost & " Plus 4 Houses"
				BuyPropBtn.Hide()
				ClosePropPopupbtn.Show()

			ElseIf Tiles(selected).Station = True Then
				'station
				PropertyPopup.Show()
				NextTurnbtn.Hide()
				PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
				PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
				RailRoadPopupDisplay.Show()
				PropertyPopupDisplay.Hide()
				UtilityPopUpDisplay.Hide()
				RailRoadPopupDisplay.BringToFront()
				RailRoadPopupDisplay.BorderStyle = BorderStyle.FixedSingle
				RailRoadPopupArt.Image = My.Resources.BottomRail
				RailRoadPopupArt.SizeMode = PictureBoxSizeMode.Zoom
				RailRoadPopupTitle.Text = Tiles(selected).Title
				BuyPropBtn.Hide()
				ClosePropPopupbtn.Show()

			ElseIf Tiles(selected).Utility = True Then
				'utility
				PropertyPopup.Show()
				NextTurnbtn.Hide()
				PropertyPopupDisplay.Hide()
				RailRoadPopupDisplay.Hide()
				PropertyPopup.Left = Me.Width / 2 - (PropertyPopup.Width / 2)
				PropertyPopup.Top = Me.Height / 2 - (PropertyPopup.Height / 2)
				UtilityPopUpDisplay.Show()
				UtilityPopUpDisplay.BringToFront()
				UtilityPopUpDisplay.BorderStyle = BorderStyle.FixedSingle
				If selected = 12 Then
					UtilityPopupArt.Image = My.Resources.Electric
				ElseIf selected = 28 Then
					UtilityPopupArt.Image = My.Resources.WaterWorks
				End If
				UtilityPopupArt.SizeMode = PictureBoxSizeMode.Zoom
				UtilityPopupTitle.Text = Tiles(selected).Title
				BuyPropBtn.Hide()
				ClosePropPopupbtn.Show()

			End If

		End If

	End Sub

	Private Sub ManageUnMortgageBtn_Click(sender As Object, e As EventArgs) Handles ManageUnMortgageBtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).Mortgaged = False Then
				MsgBox("This Property is not Mortgaged")
			Else
				If CurrentPlayers(Turn).Cash - Tiles(selected).MortgageValue >= 0 Then
					CurrentPlayers(Turn).Cash -= Tiles(selected).MortgageValue
					Tiles(selected).Mortgaged = False
                    GameActivityLogLbl.Text += Chr(13) & CurrentPlayers(Turn).Name & " Paid £" & Tiles(selected).MortgageValue & " to Unmortgage " & Tiles(selected).Title
					'Clear Mortgaged sign
					Select Case selected
						Case 1
							For i = Tile1.Controls.Count - 1 To 0 Step -1
								If Tile1.Controls(i).Tag = "Mortgaged" Then
									Tile1.Controls(i).Dispose()
								End If
							Next
						Case 3
							For i = Tile3.Controls.Count - 1 To 0 Step -1
								If Tile3.Controls(i).Tag = "Mortgaged" Then
									Tile3.Controls(i).Dispose()
								End If
							Next
						Case 6
							For i = Tile6.Controls.Count - 1 To 0 Step -1
								If Tile6.Controls(i).Tag = "Mortgaged" Then
									Tile6.Controls(i).Dispose()
								End If
							Next
						Case 8
							For i = Tile8.Controls.Count - 1 To 0 Step -1
								If Tile8.Controls(i).Tag = "Mortgaged" Then
									Tile8.Controls(i).Dispose()
								End If
							Next
						Case 9
							For i = Tile9.Controls.Count - 1 To 0 Step -1
								If Tile9.Controls(i).Tag = "Mortgaged" Then
									Tile9.Controls(i).Dispose()
								End If
							Next
						Case 11
							For i = Tile11.Controls.Count - 1 To 0 Step -1
								If Tile11.Controls(i).Tag = "Mortgaged" Then
									Tile11.Controls(i).Dispose()
								End If
							Next
						Case 13
							For i = Tile13.Controls.Count - 1 To 0 Step -1
								If Tile13.Controls(i).Tag = "Mortgaged" Then
									Tile13.Controls(i).Dispose()
								End If
							Next
						Case 14
							For i = Tile14.Controls.Count - 1 To 0 Step -1
								If Tile14.Controls(i).Tag = "Mortgaged" Then
									Tile14.Controls(i).Dispose()
								End If
							Next
						Case 16
							For i = Tile16.Controls.Count - 1 To 0 Step -1
								If Tile16.Controls(i).Tag = "Mortgaged" Then
									Tile16.Controls(i).Dispose()
								End If
							Next
						Case 18
							For i = Tile18.Controls.Count - 1 To 0 Step -1
								If Tile18.Controls(i).Tag = "Mortgaged" Then
									Tile18.Controls(i).Dispose()
								End If
							Next
						Case 19
							For i = Tile19.Controls.Count - 1 To 0 Step -1
								If Tile19.Controls(i).Tag = "Mortgaged" Then
									Tile19.Controls(i).Dispose()
								End If
							Next
						Case 21
							For i = Tile21.Controls.Count - 1 To 0 Step -1
								If Tile21.Controls(i).Tag = "Mortgaged" Then
									Tile21.Controls(i).Dispose()
								End If
							Next
						Case 23
							For i = Tile23.Controls.Count - 1 To 0 Step -1
								If Tile23.Controls(i).Tag = "Mortgaged" Then
									Tile23.Controls(i).Dispose()
								End If
							Next
						Case 24
							For i = Tile24.Controls.Count - 1 To 0 Step -1
								If Tile24.Controls(i).Tag = "Mortgaged" Then
									Tile24.Controls(i).Dispose()
								End If
							Next
						Case 26
							For i = Tile26.Controls.Count - 1 To 0 Step -1
								If Tile26.Controls(i).Tag = "Mortgaged" Then
									Tile26.Controls(i).Dispose()
								End If
							Next
						Case 27
							For i = Tile27.Controls.Count - 1 To 0 Step -1
								If Tile27.Controls(i).Tag = "Mortgaged" Then
									Tile27.Controls(i).Dispose()
								End If
							Next
						Case 29
							For i = Tile29.Controls.Count - 1 To 0 Step -1
								If Tile29.Controls(i).Tag = "Mortgaged" Then
									Tile29.Controls(i).Dispose()
								End If
							Next
						Case 31
							For i = Tile31.Controls.Count - 1 To 0 Step -1
								If Tile31.Controls(i).Tag = "Mortgaged" Then
									Tile31.Controls(i).Dispose()
								End If
							Next
						Case 32
							For i = Tile32.Controls.Count - 1 To 0 Step -1
								If Tile32.Controls(i).Tag = "Mortgaged" Then
									Tile32.Controls(i).Dispose()
								End If
							Next
						Case 34
							For i = Tile34.Controls.Count - 1 To 0 Step -1
								If Tile34.Controls(i).Tag = "Mortgaged" Then
									Tile34.Controls(i).Dispose()
								End If
							Next
						Case 37
							For i = Tile37.Controls.Count - 1 To 0 Step -1
								If Tile37.Controls(i).Tag = "Mortgaged" Then
									Tile37.Controls(i).Dispose()
								End If
							Next
						Case 39
							For i = Tile39.Controls.Count - 1 To 0 Step -1
								If Tile39.Controls(i).Tag = "Mortgaged" Then
									Tile39.Controls(i).Dispose()
								End If
							Next
					End Select
					LoadGameBoard(False)
					PopulatePropertyListBox()
					UpdatePlayerStats()
					ManageMortgagebtn.Show()
					ManageUnMortgageBtn.Hide()


				End If
			End If


		End If

	End Sub

	Private Sub ManageBuyHouseBtn_Click(sender As Object, e As EventArgs) Handles ManageBuyHouseBtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).HouseNo < 4 Then
				If CurrentPlayers(Turn).Cash - Tiles(selected).HouseCost >= 0 Then
					Tiles(selected).HouseNo += 1
					CurrentPlayers(Turn).Cash -= Tiles(selected).HouseCost
					UpdatePlayerStats()
					LoadGameBoard(False)
					PopulatePropertyListBox()
					ManageMortgagebtn.Hide()
				Else
					MsgBox("You Cannot Afford This", MsgBoxStyle.Exclamation, "Insufficient Funds")
				End If
			Else
				MsgBox("You already have maximum houses on this property!")
			End If


		End If
	End Sub

	Private Sub ManageSellHousebtn_Click(sender As Object, e As EventArgs) Handles ManageSellHousebtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).HouseNo > 0 Then
				Tiles(selected).HouseNo -= 1
				PopulatePropertyListBox()
				LoadGameBoard(False)
				If Tiles(selected).HouseNo = 0 Then
					ManageSellHotelBtn.Hide()
				End If
				CurrentPlayers(Turn).Cash += (Tiles(selected).HouseCost / 2)
				UpdatePlayerStats()

			End If
		End If

	End Sub

	Private Sub ManageBuyHotelBtn_Click(sender As Object, e As EventArgs) Handles ManageBuyHotelBtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).HouseNo = 4 Then
				If Tiles(selected).Hotel = False Then
					If CurrentPlayers(Turn).Cash - Tiles(selected).HotelCost >= 0 Then
						Tiles(selected).HouseNo = 0
						Tiles(selected).Hotel = True
						CurrentPlayers(Turn).Cash -= Tiles(selected).HotelCost
						UpdatePlayerStats()
						LoadGameBoard(False)
						PopulatePropertyListBox()
						ManageBuyHotelBtn.Hide()
					Else
						MsgBox("You Cannot Afford This", MsgBoxStyle.Exclamation, "Insufficient Funds")
					End If
				Else
					MsgBox("This property Already has a hotel")
				End If
			End If

		End If

	End Sub

	Private Sub ManageSellHotelBtn_Click(sender As Object, e As EventArgs) Handles ManageSellHotelBtn.Click
		Dim selected As Integer

		If PropList.SelectedItems.Count > 0 Then
			For n = 0 To 39
				If PropList.SelectedItems.Item(0).Text = Tiles(n).Title Then
					selected = n
					Exit For
				End If
			Next

			If Tiles(selected).Hotel = True Then
				Tiles(selected).Hotel = False
				CurrentPlayers(Turn).Cash += (Tiles(selected).HotelCost / 2)
				Tiles(selected).HouseNo = 4

				UpdatePlayerStats()
				PopulatePropertyListBox()
				LoadGameBoard(False)
				ManageBuyHotelBtn.Show()
				ManageSellHotelBtn.Hide()
				ManageSellHousebtn.Show()
				ManageBuyHouseBtn.Hide()


			Else
				MsgBox("There is no Hotel here")
			End If



		End If


	End Sub

	Private Sub Playerbankrupted(Player As Integer)
		If Tiles(CurrentPlayers(Player).CurrentPos).Owner <> "N/A" Or Tiles(CurrentPlayers(Player).CurrentPos).Owner <> "" Then
			'Player Bankrupted by other player
			For i = 0 To 39
				'Transfer Properties
				If Tiles(i).Owner = CurrentPlayers(Player).Name Then
					Tiles(i).Cost = Tiles(CurrentPlayers(Player).CurrentPos).Owner
					Tiles(i).Owner = Tiles(CurrentPlayers(Player).CurrentPos).Owner
					LoadGameBoard(False)
				End If
			Next
		Else
			'Player Bankrupted by Game
			For i = 0 To 39
				'Free up Properties
				If Tiles(i).Owner = CurrentPlayers(Player).Name Then
					Tiles(i).Owner = ""
					Select Case i
						Case 1
							Tiles(i).Cost = 60
						Case 3
							Tiles(i).Cost = 60
						Case 5
							Tiles(i).Cost = 200
						Case 6
							Tiles(i).Cost = 100
						Case 8
							Tiles(i).Cost = 100
						Case 9
							Tiles(i).Cost = 120
						Case 11
							Tiles(i).Cost = 140
						Case 12
							Tiles(i).Cost = 150
						Case 13
							Tiles(i).Cost = 140
						Case 14
							Tiles(i).Cost = 160
						Case 15
							Tiles(i).Cost = 200
						Case 16
							Tiles(i).Cost = 180
						Case 18
							Tiles(i).Cost = 180
						Case 19
							Tiles(i).Cost = 200
						Case 21
							Tiles(i).Cost = 220
						Case 23
							Tiles(i).Cost = 220
						Case 24
							Tiles(i).Cost = 240
						Case 25
							Tiles(i).Cost = 200
						Case 26
							Tiles(i).Cost = 260
						Case 27
							Tiles(i).Cost = 260
						Case 28
							Tiles(i).Cost = 150
						Case 29
							Tiles(i).Cost = 280
						Case 31
							Tiles(i).Cost = 300
						Case 32
							Tiles(i).Cost = 300
						Case 34
							Tiles(i).Cost = 32
						Case 35
							Tiles(i).Cost = 200
						Case 37
							Tiles(i).Cost = 350
						Case 39
							Tiles(i).Cost = 400
					End Select
					LoadGameBoard(False)
				End If
			Next
		End If
		Select Case Player
			Case 1
				Player1Token.Hide()
			Case 2
				Player2Token.Hide()
			Case 3
				Player3Token.Hide()
			Case 4
				Player4Token.Hide()
			Case 5
				Player5Token.Hide()
			Case 6
				Player6Token.Hide()
			Case 7
				Player7Token.Hide()
			Case 8
				Player8Token.Hide()
		End Select
		CurrentPlayers(Player).Bankrupt = True
		Dim PlayersNotBankrupt As Integer
		For i = 0 To PlayerNumber - 1
			If CurrentPlayers(i).Bankrupt = False Then
				PlayersNotBankrupt += 1
			End If
		Next
		If PlayersNotBankrupt = 1 Then
			GameOver()
		End If
	End Sub

	Private Sub GameOver()
		For i = 0 To PlayerNumber - 1
			If CurrentPlayers(i).Bankrupt = False Then
				MsgBox("Results" & Chr(13) & CurrentPlayers(i).Name & " Wins With a Net Worth of £" & CalculateNetWorth(i) & Chr(13) & "All Other Players Are Bankrupt", MsgBoxStyle.OkOnly, "RESULTS")
				Exit For
			End If
		Next
		GameBoard.Hide()
		LoadingPanel.Hide()
		GamePanel.Hide()
		MainMenuPanel.Show()
	End Sub

	Function CalculateNetWorth(Plr As Integer) As Integer
		Dim networth As Integer

		For i = 0 To 39
			If Tiles(i).Owner = CurrentPlayers(Plr).Name Then
				If Tiles(i).Hotel = True Then
					networth += ((Tiles(i).HouseCost / 2) * 5)
				End If
				If Tiles(i).HouseNo > 0 Then
					networth += ((Tiles(i).HouseCost / 2) * Tiles(i).HouseNo)
				End If
				If Tiles(i).Mortgaged = False Then
					networth += Tiles(i).MortgageValue
				End If
				networth += CurrentPlayers(Plr).Cash
				If CurrentPlayers(Plr).GetOutJailCards > 0 Then
					networth += CurrentPlayers(Plr).GetOutJailCards * 50
				End If
			End If
		Next

		Return networth
	End Function

	Private Sub KeyPressed(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
		If GameBoard.Visible = True And InGameMenuPanel.Visible = False And PropertyPopup.Visible = False And CardPopupPanel.Visible = False Then
			If e.KeyData = Keys.Escape Then
				Pausecolour = Me.BackColor
				InGameMenuPanel.Show()
				InGameMenuPanel.Location = TurnIndicatorlbl.Location
				InGameMenuPanel.Left -= 30
				InGameMenuPanel.Top -= 30
				InGameMenuPanel.BorderStyle = BorderStyle.FixedSingle
				InGameMenuPanel.BackColor = Pausecolour
				Me.BackColor = ColorTranslator.FromHtml("#5B5B5B")
				NextTurnbtn.Hide()
				ManagePropBtn.Hide()

				If GameDie1Spinner.Enabled = True Then
					gamedie1active = True
					GameDie1Spinner.Stop()
				Else
					gamedie1active = False
				End If
				If GameDie2Spinner.Enabled = True Then
					gamedie2active = True
					GameDie2Spinner.Stop()
				Else
					gamedie2active = False
				End If
				If Flasher.Enabled = True Then
					flasheractive = True
					Flasher.Stop()
				Else
					flasheractive = False
				End If
				If Mover.Enabled = True Then
					moveractive = True
					Mover.Stop()
				Else
					moveractive = False
				End If
			End If
		End If

	End Sub

	Private Sub MenuResumeBtn_Click(sender As Object, e As EventArgs) Handles MenuResumeBtn.Click
		If moveractive = False And gamedie1active = False And gamedie2active = False And RollDiceBtn.Visible = False Then
			NextTurnbtn.Show()
		End If
		For i = 0 To 39
			If Tiles(i).Owner = CurrentPlayers(Turn).Name Then
				ManagePropBtn.Show()
				Exit For
			End If
		Next
		InGameMenuPanel.Hide()

		If gamedie1active = True Then
			GameDie1Spinner.Start()
		End If
		If gamedie2active = True Then
			GameDie2Spinner.Start()
		End If
		If flasheractive = True Then
			Flasher.Start()
		End If
		If moveractive = True Then
			Mover.Start()
		End If
		Me.BackColor = Pausecolour
	End Sub

	Private Sub MenuQuitBtn_Click(sender As Object, e As EventArgs) Handles MenuQuitBtn.Click
		If MsgBox("All Unsaved Progress Will Be Lost" & Chr(13) & "Are You Sure?", MsgBoxStyle.YesNo, "Quit Game?") = MsgBoxResult.Yes Then
			End
		End If
	End Sub

	Private Sub MenuSaveBtn_Click(sender As Object, e As EventArgs) Handles MenuSaveBtn.Click
		Dim gamename As String
		gamename = InputBox("Name This Game", "Save Game", "")
		gamename = Directory.GetCurrentDirectory & "\Saved Games\" & gamename

		If My.Computer.FileSystem.DirectoryExists(Directory.GetCurrentDirectory & "\Saved Games") = False Then
			My.Computer.FileSystem.CreateDirectory(Directory.GetCurrentDirectory & "\Saved Games")
		End If

		Try
			FileOpen(1, gamename, OpenMode.Binary)
			FilePut(1, CurrentPlayers)
			'FilePut(1, Tiles)
			'FilePut(1, CurrentHouseRules)
			FileClose(1)
		Catch ex As Exception
			MsgBox("Invalid Name, Try again")
		End Try

		Dim count As Integer
		count = 0
		Const searchChar = "\"c
		For Each foundFile As String In My.Computer.FileSystem.GetFiles(Directory.GetCurrentDirectory & "\Saved Games\")
			SavedFileNames(count) = foundFile.Substring(foundFile.LastIndexOf(searchChar, foundFile.LastIndexOf(searchChar)) + 1)
			count += 1
		Next
	End Sub


	Private Sub PopulateLoadGameListBox()
		Dim count As Integer
		Const searchChar = "\"c

		Try
			count = 0
			For Each foundFile As String In My.Computer.FileSystem.GetFiles(Directory.GetCurrentDirectory & "\Saved Games\")
				SavedFileNames(count) = foundFile.Substring(foundFile.LastIndexOf(searchChar, foundFile.LastIndexOf(searchChar)) + 1)
				count += 1
			Next
		Catch ex As Exception
			MsgBox("No Saved Games Exist")
		End Try


		If count = 0 Then
			MsgBox("There are not any saved games yet", MsgBoxStyle.Information, "No Games")
		End If

		LoadListBox.Items.Clear()

		For i = 0 To 50
			If SavedFileNames(i) <> "" Or SavedFileNames(i) <> Nothing Then
				LoadListBox.Items.Add(SavedFileNames(i))
			End If

		Next
	End Sub
	Private Sub MainLoadGame_Click(sender As Object, e As EventArgs) Handles MainLoadGame.Click
		LoadGamePanel.Show()
		LoadGamePanel.BringToFront()
		LoadGamePanel.Size = Me.Size
		LoadGamePanel.Top = 0
		LoadGamePanel.Left = 0
		PopulateLoadGameListBox()
	End Sub

	Private Sub LoadMainMenubtn_Click(sender As Object, e As EventArgs) Handles LoadMainMenubtn.Click
		LoadGamePanel.Hide()
	End Sub

	Private Sub LoadDeleteBtn_Click(sender As Object, e As EventArgs) Handles LoadDeleteBtn.Click
		Dim selected As Integer

		selected = LoadListBox.SelectedIndex

		If IsNothing(selected) Then
			MsgBox("Select a game first")
		Else
			DeleteGameNumber(selected)
		End If
	End Sub

	Private Sub DeleteGameNumber(Game As Integer)
		My.Computer.FileSystem.DeleteFile(Directory.GetCurrentDirectory & "\Saved Games\" & SavedFileNames(Game))
		For i = 0 To SavedFileNames.Length - 1
			SavedFileNames(i) = Nothing
		Next
		PopulateLoadGameListBox()
	End Sub

	Private Sub LoadLoadbtn_Click(sender As Object, e As EventArgs) Handles LoadLoadbtn.Click
		'Load Game
		Dim selected As Integer
		ReDim Preserve CurrentPlayers(7)
		selected = LoadListBox.SelectedIndex

		If IsNothing(selected) Then
			MsgBox("Select a game first")
		Else
			'Load Game
			Dim gameSave As String = Directory.GetCurrentDirectory & "\Saved Games\" & SavedFileNames(selected)
			FileOpen(1, gameSave, OpenMode.Binary)
			FileGet(1, CurrentPlayers, 1)
			'FileGet(1, Tiles)
			'FileGet(1, CurrentHouseRules)
			FileClose(1)
			MsgBox(CurrentPlayers(0).Name)
		End If

	End Sub

	Private Sub OpenTradeBtn_Click(sender As Object, e As EventArgs) Handles OpenTradeBtn.Click
		TradePanel.Show()
		Trader1Name.Text = CurrentPlayers(Turn).Name
		TradePanel.BringToFront()
		Trader1Cards = 0
		Trader1Cash = 0
		Trader2Cash = 0
		Trader1Cards = 0

		Trader2Ownedlbl.Hide()
		Trader2OwnedList.Hide()
		Trader2OfferList.Hide()
		Trader2RemoveBtn.Hide()
		Trader2RemoveJailCardbtn.Hide()
		Trader2AcceptDealBtn.Hide()
		Trader2AddBtn.Hide()
		Trader2AddJailCardbtn.Hide()
		Trader2CashBtn.Hide()
		Trader2Cashlbl.Hide()
		Trader2JailCardslbl.Hide()
		Trader1AcceptDealbtn.Text = "Accept Deal"
		Trader1AcceptDealbtn.Hide()
		Trader2AcceptDealBtn.Text = "Accept Deal"
		Trader2AcceptDealBtn.Hide()
		Trader1Cashlbl.Hide()
		Trader1JailCardlbl.Hide()


		'Populate Trader 1 Owned List
		Trader1OwnedList.Items.Clear()
		Trader1OfferList.Items.Clear()
		For i = 0 To 39
			If Tiles(i).Owner = CurrentPlayers(Turn).Name Then
				Trader1OwnedList.Items.Add(Tiles(i).Title)
			End If
		Next

		If CurrentPlayers(Turn).Cash > 0 Then
			Trader1Cashbtn.Show()
		Else
			Trader1Cashbtn.Hide()
		End If

		If CurrentPlayers(Turn).GetOutJailCards > 0 Then
			Trader1AddJailCardbtn.Show()
		Else
			Trader1AddJailCardbtn.Hide()
		End If

		Trader2NameSelect.Items.Clear()

		For i = 0 To PlayerNumber - 1
			If i = Turn Then
				Continue For
			Else
				Trader2NameSelect.Items.Add(CurrentPlayers(i).Name)
			End If
		Next

	End Sub

	Private Sub Trader1ListClicked(sender As Object, e As EventArgs) Handles Trader1OwnedList.Click
		Trader1Addbtn.Show()
		Trader1Removebtn.Hide()
	End Sub

	Private Sub TradeCancelBtn_Click(sender As Object, e As EventArgs) Handles TradeCancelBtn.Click
		TradePanel.Hide()
	End Sub

	Private Sub Trader1OfferListClicked(sender As Object, e As EventArgs) Handles Trader1OfferList.Click
		Trader1Addbtn.Hide()
		Trader1Removebtn.Show()
	End Sub

	Private Sub Trader1Addbtn_Click(sender As Object, e As EventArgs) Handles Trader1Addbtn.Click
		If Trader1AcceptDealbtn.Text <> "Decline Deal" Then
			If Trader1OwnedList.SelectedItem <> "" Or Trader1OwnedList.SelectedItem <> Nothing Then
				Trader1OfferList.Items.Add(Trader1OwnedList.SelectedItem)
				Trader1OwnedList.Items.Remove(Trader1OwnedList.SelectedItem)
			End If
			Trader1AcceptDealbtn.Show()
		Else
			MsgBox("Deal Locked In")
		End If

	End Sub

	Private Sub Trader1Removebtn_Click(sender As Object, e As EventArgs) Handles Trader1Removebtn.Click
		If Trader1AcceptDealbtn.Text <> "Decline Deal" Then
			If Trader1OfferList.SelectedItem <> "" Or Trader1OfferList.SelectedItem <> Nothing Then
				Trader1OwnedList.Items.Add(Trader1OfferList.SelectedItem)
				Trader1OfferList.Items.Remove(Trader1OfferList.SelectedItem)
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader1Cashbtn_Click(sender As Object, e As EventArgs) Handles Trader1Cashbtn.Click
		If Trader1AcceptDealbtn.Text <> "Decline Deal" Then
			Dim CashToAdd As Integer
			Dim UserInput As String

			UserInput = InputBox("Add Cash:" & Chr(13) & "Cash Available: £" & CurrentPlayers(Turn).Cash - Trader1Cash & Chr(13) & "(Entering a Negative Will Subtract Cash)")

			If IsNumeric(UserInput) Then
				CashToAdd = CInt(UserInput)
				If Abs(CashToAdd) <= CurrentPlayers(Turn).Cash - Trader1Cash Then
					Trader1Cash += CashToAdd
					Trader1Cashlbl.Text = "+ £" & Trader1Cash
					If Trader1Cash > 0 Then
						Trader1Cashlbl.Show()
					Else
						Trader1Cashlbl.Hide()
					End If
				Else
					MsgBox("Invalid Entry")
				End If
			Else
				MsgBox("Invalid Entry")
			End If

			If Trader1Cash < 0 Then
				Trader1Cash = 0
			End If

			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader1AddJailCardbtn_Click(sender As Object, e As EventArgs) Handles Trader1AddJailCardbtn.Click
		If Trader1AcceptDealbtn.Text <> "Decline Deal" Then
			If CurrentPlayers(Turn).GetOutJailCards - Trader1Cards > 0 Then
				Trader1Cards += 1
				Trader1RemoveJailCardsbtn.Show()
				If Trader1Cards > 1 Then
					Trader1JailCardlbl.Text = "+ " & Trader1Cards & "Get Out Of Jail Free Cards"
					Trader1JailCardlbl.Show()
				Else
					Trader1JailCardlbl.Text = "+ " & Trader1Cards & "Get Out Of Jail Free Card"
					Trader1JailCardlbl.Show()
				End If

			Else
				MsgBox("You Have No More Cards To Add!")
			End If
			If Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0 Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader1RemoveJailCardsbtn_Click(sender As Object, e As EventArgs) Handles Trader1RemoveJailCardsbtn.Click
		If Trader1AcceptDealbtn.Text <> "Decline Deal" Then
			If Trader1Cards > 0 Then
				Trader1Cards -= 1
				If Trader1Cards > 1 Then
					Trader1JailCardlbl.Text = "+ " & Trader1Cards & "Get Out Of Jail Free Cards"
					Trader1JailCardlbl.Show()
				Else
					Trader1JailCardlbl.Text = "+ " & Trader1Cards & "Get Out Of Jail Free Card"
					Trader1JailCardlbl.Show()
				End If
				If Trader1Cards = 0 Then
					Trader1JailCardlbl.Hide()
					Trader1RemoveJailCardsbtn.Hide()
				End If
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader2NameSelected(sender As Object, e As EventArgs) Handles Trader2NameSelect.SelectedIndexChanged

		Dim selected As Integer


		For i = 0 To PlayerNumber - 1
			If Trader2NameSelect.SelectedItem = CurrentPlayers(i).Name Then
				selected = i
				Exit For
			End If
		Next

		If selected >= 0 And selected <= 7 Then
			'Populate and Display Boxes
			Trader2OfferList.Show()
			Trader2OwnedList.Show()
			Trader2OfferList.Items.Clear()
			Trader2OwnedList.Items.Clear()
			Trader2Ownedlbl.Show()

			For i = 0 To 39
				If Tiles(i).Owner = CurrentPlayers(selected).Name Then
					Trader2OwnedList.Items.Add(Tiles(i).Title)
				End If
			Next

			If CurrentPlayers(selected).Cash > 0 Then
				Trader2CashBtn.Show()
			Else
				Trader2CashBtn.Hide()
			End If

			If CurrentPlayers(selected).GetOutJailCards > 0 Then
				Trader2AddJailCardbtn.Show()
			Else
				Trader2AddJailCardbtn.Hide()
			End If

			Trader2ID = selected

		Else
			MsgBox("Invalid Selection")
		End If
	End Sub

	Private Sub Trader2OfferClicked(sender As Object, e As EventArgs) Handles Trader2OfferList.Click
		Trader2AddBtn.Hide()
		Trader2RemoveBtn.Show()
	End Sub

	Private Sub Trader2OwnedClicked(sender As Object, e As EventArgs) Handles Trader2OwnedList.Click
		Trader2AddBtn.Show()
		Trader2RemoveBtn.Hide()
	End Sub

	Private Sub Trader2AddBtn_Click(sender As Object, e As EventArgs) Handles Trader2AddBtn.Click
		If Trader2AcceptDealBtn.Text <> "Decline Deal" Then
			If Trader2OwnedList.SelectedItem <> "" Or Trader2OwnedList.SelectedItem <> Nothing Then
				Trader2OfferList.Items.Add(Trader2OwnedList.SelectedItem)
				Trader2OwnedList.Items.Remove(Trader2OwnedList.SelectedItem)
			End If
			Trader2AcceptDealBtn.Show()
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader2RemoveBtn_Click(sender As Object, e As EventArgs) Handles Trader2RemoveBtn.Click
		If Trader2AcceptDealBtn.Text <> "Decline Deal" Then
			If Trader2OfferList.SelectedItem <> "" Or Trader2OfferList.SelectedItem <> Nothing Then
				Trader2OwnedList.Items.Add(Trader2OfferList.SelectedItem)
				Trader2OfferList.Items.Remove(Trader2OfferList.SelectedItem)
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader2CashBtn_Click(sender As Object, e As EventArgs) Handles Trader2CashBtn.Click
		If Trader2AcceptDealBtn.Text <> "Decline Deal" Then
			Dim CashToAdd As Integer
			Dim UserInput As String

			UserInput = InputBox("Add Cash:" & Chr(13) & "Cash Available: £" & CurrentPlayers(Trader2ID).Cash - Trader2Cash & Chr(13) & "(Entering a Negative Will Subtract Cash)")

			If IsNumeric(UserInput) Then
				CashToAdd = CInt(UserInput)
				If Abs(CashToAdd) <= CurrentPlayers(Trader2ID).Cash - Trader2Cash Then
					Trader2Cash += CashToAdd
					Trader2Cashlbl.Text = "+ £" & Trader2Cash
					If Trader2Cash > 0 Then
						Trader2Cashlbl.Show()
					Else
						Trader2Cashlbl.Hide()
					End If
				Else
					MsgBox("Invalid Entry")
				End If
			Else
				MsgBox("Invalid Entry")
			End If

			If Trader2Cash < 0 Then
				Trader2Cash = 0
			End If

			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader2AddJailCardbtn_Click(sender As Object, e As EventArgs) Handles Trader2AddJailCardbtn.Click
		If Trader2AcceptDealBtn.Text <> "Decline Deal" Then
			If CurrentPlayers(Trader2ID).GetOutJailCards - Trader2Cards > 0 Then
				Trader2Cards += 1
				Trader2RemoveJailCardbtn.Show()
				If Trader2Cards > 1 Then
					Trader2JailCardslbl.Text = "+ " & Trader2Cards & "Get Out Of Jail Free Cards"
					Trader2JailCardslbl.Show()
				Else
					Trader2JailCardslbl.Text = "+ " & Trader2Cards & "Get Out Of Jail Free Card"
					Trader2JailCardslbl.Show()
				End If

			Else
				MsgBox("You Have No More Cards To Add!")
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader2RemoveJailCardbtn_Click(sender As Object, e As EventArgs) Handles Trader2RemoveJailCardbtn.Click
		If Trader2AcceptDealBtn.Text <> "Decline Deal" Then
			If Trader2Cards > 0 Then
				Trader2Cards -= 1
				If Trader2Cards > 1 Then
					Trader2JailCardslbl.Text = "+ " & Trader2Cards & "Get Out Of Jail Free Cards"
					Trader2JailCardslbl.Show()
				Else
					Trader2JailCardslbl.Text = "+ " & Trader2Cards & "Get Out Of Jail Free Card"
					Trader2JailCardslbl.Show()
				End If
				If Trader2Cards = 0 Then
					Trader2JailCardslbl.Hide()
					Trader2RemoveJailCardbtn.Hide()
				End If
			End If
			If (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) And (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) Then
				Trader2AcceptDealBtn.Hide()
			Else
				Trader2AcceptDealBtn.Show()
			End If
			If (Trader1OfferList.Items.Count = 0 And Trader1Cash = 0 And Trader1Cards = 0) And (Trader2OfferList.Items.Count = 0 And Trader2Cash = 0 And Trader2Cards = 0) Then
				Trader1AcceptDealbtn.Hide()
			Else
				Trader1AcceptDealbtn.Show()
			End If
		Else
			MsgBox("Deal Locked In")
		End If
	End Sub

	Private Sub Trader1AcceptDealbtn_Click(sender As Object, e As EventArgs) Handles Trader1AcceptDealbtn.Click
		If Trader1AcceptDealbtn.Text = "Accept Deal" Then
			Trader1AcceptDealbtn.Text = "Decline Deal"
			If Trader2AcceptDealBtn.Text = "Decline Deal" Then
				MakeDeal()
			End If
		Else
			Trader1AcceptDealbtn.Text = "Accept Deal"
		End If
	End Sub

	Private Sub Trader2AcceptDealBtn_Click(sender As Object, e As EventArgs) Handles Trader2AcceptDealBtn.Click
		If Trader2AcceptDealBtn.Text = "Accept Deal" Then
			Trader2AcceptDealBtn.Text = "Decline Deal"
			If Trader1AcceptDealbtn.Text = "Decline Deal" Then
				MakeDeal()
			End If
		Else
			Trader2AcceptDealBtn.Text = "Accept Deal"
		End If
	End Sub

	Private Sub MakeDeal()
		'Make Deal With current offers

		'Transfer Cash
		CurrentPlayers(Trader2ID).Cash += Trader1Cash
		CurrentPlayers(Turn).Cash -= Trader1Cash
		CurrentPlayers(Turn).Cash += Trader2Cash
		CurrentPlayers(Trader2ID).Cash -= Trader2Cash
		'Reset
		Trader1Cash = 0
		Trader2Cash = 0

		'Transfer Cards
		CurrentPlayers(Turn).GetOutJailCards -= Trader1Cards
		CurrentPlayers(Trader2ID).GetOutJailCards -= Trader2Cards
		CurrentPlayers(Turn).GetOutJailCards += Trader2Cards
		CurrentPlayers(Trader2ID).GetOutJailCards += Trader1Cards
		'Reset
		Trader1Cards = 0
		Trader2Cards = 0

		'Transfer Properties

		'Trader 1 --> Trader 2
		If Trader1OfferList.Items.Count > 0 Then
			For i = 0 To Trader1OfferList.Items.Count - 1
				'Find it
				For n = 0 To 39
					If Tiles(n).Title = Trader1OfferList.Items(i) Then
						'Transfer
						Tiles(n).Owner = CurrentPlayers(Trader2ID).Name
						Tiles(n).Cost = CurrentPlayers(Trader2ID).Name
					End If
				Next
			Next
		End If

		'Trader 2 --> Trader 1
		If Trader2OfferList.Items.Count > 0 Then
			For i = 0 To Trader2OfferList.Items.Count - 1
				'Find it
				For n = 0 To 39
					If Tiles(n).Title = Trader2OfferList.Items(i) Then
						'Transfer
						Tiles(n).Owner = CurrentPlayers(Turn).Name
						Tiles(n).Cost = CurrentPlayers(Turn).Name
					End If
				Next
			Next
		End If

		'Reload
		TradePanel.Hide()

		UpdatePlayerStats()
		LoadGameBoard(False)
		PopulatePropertyListBox()
		MsgBox("Trade Complete")

	End Sub
End Class


Public Class Randomizer(Of T)
	Implements IComparer(Of T)

	''// Ensures different instances are sorted in different orders
	Private Shared Salter As New Random() ''// only as random as your seed
	Private Salt As Integer
	Public Sub New()
		Salt = Salter.Next(Integer.MinValue, Integer.MaxValue)
	End Sub

	Private Shared sha As New SHA1CryptoServiceProvider()
	Private Function HashNSalt(ByVal x As Integer) As Integer
		Dim b() As Byte = sha.ComputeHash(BitConverter.GetBytes(x))
		Dim r As Integer = 0
		For i As Integer = 0 To b.Length - 1 Step 4
			r = r Xor BitConverter.ToInt32(b, i)
		Next

		Return r Xor Salt
	End Function

	Public Function Compare(x As T, y As T) As Integer _
		Implements IComparer(Of T).Compare

		Return HashNSalt(x.GetHashCode()).CompareTo(HashNSalt(y.GetHashCode()))
	End Function
End Class


