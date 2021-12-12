using GoogleSheetsHelper;

namespace StandingsGoogleSheetsHelper
{
	/// <summary>
	/// Class to generate formulas for the standings table
	/// </summary>
	public class FormulaGenerator
    {
		private readonly StandingsSheetHelper _sheetHelper;
		public StandingsSheetHelper SheetHelper { get => _sheetHelper; }

		public FormulaGenerator(StandingsSheetHelper helper)
		{
			_sheetHelper = helper;
		}

		/// <summary>
		/// Gets the formula for determining who won the game
		/// </summary>stat
		/// <param name="rowNum">Row number</param>
		/// <returns></returns>
		/// <remarks>=IFS(OR(ISBLANK(B3),ISBLANK(C3)),"",B3=C3,"D",B3&gt;C3,"H",B3&lt;C3,"A")
		/// = if home or away goals scored is blank, then return blank
		/// otherwise if goals are equal return D, if home &gt; away return H, else return A
		/// </remarks>
		public string GetGameWinnerFormula(int rowNum)
		{
			string homeGoalsCell = $"{_sheetHelper.HomeGoalsColumnName}{rowNum}";
			string awayGoalsCell = $"{_sheetHelper.AwayGoalsColumnName}{rowNum}";
			return string.Format("=IFS(OR(ISBLANK({0}), ISBLANK({1})), \"\", {0}>{1}, \"{2}\", {0}<{1}, \"{3}\", {0}={1}, \"{4}\")",
				homeGoalsCell,
				awayGoalsCell,
				Constants.HOME_TEAM_INDICATOR,
				Constants.AWAY_TEAM_INDICATOR,
				Constants.DRAW_INDICATOR);
		}

		/// <summary>
		/// Gets the formula for determining the number of games played
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,B$21:B$28,"&lt;&gt;")+COUNTIFS(D$21:D$28,"="&Teams!A2,C$21:C$28,"&lt;&gt;")
		/// = count of number of times team name appears in Home column + same for Away column but only when a score has also been entered (that's the "&lt;&gt;" part)</remarks>
		public string GetGamesPlayedFormula(int startRowNum, int endRowNum, string firstTeamCell)
		{
			string homeTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.HomeTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string awayTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.AwayTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string homeScoreFormula = GetFormulaForIgnoringBlankScores(_sheetHelper.HomeGoalsColumnName, startRowNum, endRowNum);
			string awayScoreFormula = GetFormulaForIgnoringBlankScores(_sheetHelper.AwayGoalsColumnName, startRowNum, endRowNum);
			return string.Format("COUNTIFS({0},{1})+COUNTIFS({2},{3})",
				homeTeamFormula,
				homeScoreFormula,
				awayTeamFormula,
				awayScoreFormula);
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"H")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"A")
		/// = count number of times team name appears in Home column AND winning team = H + same for away column AND winning team = A</remarks>
		public string GetGamesWonFormula(int startRowNum, int endRowNum, string firstTeamCell)
		{
			string homeTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.HomeTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string awayTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.AwayTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string whoWonCellRange = Utilities.CreateCellRangeString(_sheetHelper.WinnerColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			return $"COUNTIFS({homeTeamFormula},{whoWonCellRange},\"{Constants.HOME_TEAM_INDICATOR}\")+COUNTIFS({awayTeamFormula},{whoWonCellRange},\"{Constants.AWAY_TEAM_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"A")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"H")
		/// = count number of times team name appears in Home column AND winning team = A + same for away column AND winning team = H</remarks>
		public string GetGamesLostFormula(int startRowNum, int endRowNum, string firstTeamCell)
		{
			string homeTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.HomeTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string awayTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.AwayTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string whoWonCellRange = Utilities.CreateCellRangeString(_sheetHelper.WinnerColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			return $"COUNTIFS({homeTeamFormula},{whoWonCellRange},\"{Constants.AWAY_TEAM_INDICATOR}\")+COUNTIFS({awayTeamFormula},{whoWonCellRange},\"{Constants.HOME_TEAM_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"D")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"D")
		/// = count number of times team name appears in Home column AND winning team = D + same for away column</remarks>
		public string GetGamesDrawnFormula(int startRowNum, int endRowNum, string firstTeamCell)
		{
			string homeTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.HomeTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string awayTeamFormula = GetFormulaForGameRangePerTeam(_sheetHelper.AwayTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string whoWonCellRange = Utilities.CreateCellRangeString(_sheetHelper.WinnerColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			return $"COUNTIFS({homeTeamFormula},{whoWonCellRange},\"{Constants.DRAW_INDICATOR}\")+COUNTIFS({awayTeamFormula},{whoWonCellRange},\"{Constants.DRAW_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the team rank
		/// </summary>
		/// <param name="startRowNum">Row number of the first team in the standings table</param>
		/// <param name="endRowNum">Row number of the last team in the standings table</param>
		/// <returns></returns>
		/// <remarks>RANK(M3,M$3:M$18)</remarks>
		public string GetTeamRankFormula(int startRowNum, int endRowNum)
		{
			string columnName = _sheetHelper.RankColumnName;
			string cellRange = Utilities.CreateCellRangeString(columnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			return $"RANK({columnName}{startRowNum},{cellRange})";
		}

		/// <summary>
		/// Gets the formula for determining the number of goals scored
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>SUMIFS(B$21:B$28, A$21:A$28,"="&Teams!A2)+SUMIFS(C$21:C$28, D$21:D$28,"="&Teams!A2)
		/// = sum of home goals column where home team = team name + sum of away goals column where away team = team name
		/// When doing goals against, swap the home and away goal columns</remarks>
		public string GetGoalsScoredFormula(int startRowNum, int endRowNum, string firstTeamCell) => GetGoalsFormula(startRowNum, endRowNum, firstTeamCell, true);

		/// <summary>
		/// Gets the formula for determining the number of goals conceded
		/// </summary>
		/// <param name="startRowNum">The starting row number of the game scores that need to be considered</param>
		/// <param name="endRowNum">The end row number of the game scores that need to be considered</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>SUMIFS(C$21:C$28, A$21:A$28,"="&Teams!A2)+SUMIFS(B$21:B$28, D$21:D$28,"="&Teams!A2)
		/// = sum of away goals column where home team = team name + sum of home goals column where away team = team name
		/// </remarks>
		public string GetGoalsAgainstFormula(int startRowNum, int endRowNum, string firstTeamCell) => GetGoalsFormula(startRowNum, endRowNum, firstTeamCell, false);

		private string GetGoalsFormula(int startRowNum, int endRowNum, string firstTeamCell, bool goalsFor)
		{
			string homeGoalsCellRange = Utilities.CreateCellRangeString(_sheetHelper.HomeGoalsColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string awayGoalsCellRange = Utilities.CreateCellRangeString(_sheetHelper.AwayGoalsColumnName, startRowNum, endRowNum, CellRangeOptions.FixRow);
			string homeTeamsCellRange = GetFormulaForGameRangePerTeam(_sheetHelper.HomeTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			string awayTeamsCellRange = GetFormulaForGameRangePerTeam(_sheetHelper.AwayTeamColumnName, startRowNum, endRowNum, firstTeamCell);
			return $"SUMIFS({(goalsFor ? homeGoalsCellRange : awayGoalsCellRange)}, {homeTeamsCellRange})+SUMIFS({(goalsFor ? awayGoalsCellRange : homeGoalsCellRange)}, {awayTeamsCellRange})";
		}

		/// <summary>
		/// Gets the formula for goal differential
		/// </summary>
		/// <param name="goalsForStartCell"></param>
		/// <param name="goalsAgainstStartCell"></param>
		/// <returns></returns>
		/// <remarks>P3 - Q3</remarks>
		public string GetGoalDifferentialFormula(string goalsForStartCell, string goalsAgainstStartCell)
		{
			return $"{goalsForStartCell} - {goalsAgainstStartCell}";
		}

		private string GetFormulaForGameRangePerTeam(string columnName, int startRow, int endRow, string firstTeamCell)
		{
			string cellRange = Utilities.CreateCellRangeString(columnName, startRow, endRow, CellRangeOptions.FixRow);
			return $"{cellRange},\"=\"&{firstTeamCell}"; // A$21:A$28,"="&Teams!A2
		}

		private string GetFormulaForIgnoringBlankScores(string columnName, int startRow, int endRow)
		{
			string cellRange = Utilities.CreateCellRangeString(columnName, startRow, endRow, CellRangeOptions.FixRow);
			return $"{cellRange},\"<>\""; // B$21:B$28,"<>"
		}
	}
}
