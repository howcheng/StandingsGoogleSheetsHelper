using Xunit;

namespace StandingsGoogleSheetsHelper.Tests
{
	public class FormulaGeneratorTests
	{
		// Expected formulas were all copied directly from the 2021 season sheet
		private const int START_GAMES_ROW = 3;
		private const int END_GAMES_ROW = 8;
		private const int END_TEAMS_ROW = 18;
		private const string FIRST_TEAM_CELL = "Teams!A2";

		private FormulaGenerator GetFormulaGenerator() => new FormulaGenerator(new TestSheetHelper());

		[Fact]
		public void TestGameWinnerFormula()
		{
			const string expected = "=IFS(OR(ISBLANK(B3), ISBLANK(C3)), \"\", B3>C3, \"H\", B3<C3, \"A\", B3=C3, \"D\")";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGameWinnerFormula(START_GAMES_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGamesPlayedFormula()
		{
			const string expected = "COUNTIFS(A$3:A$8,\"=\"&Teams!A2,B$3:B$8,\"<>\")+COUNTIFS(D$3:D$8,\"=\"&Teams!A2,C$3:C$8,\"<>\")";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGamesPlayedFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGamesWonFormula()
		{
			const string expected = "COUNTIFS(A$3:A$8,\"=\"&Teams!A2,E$3:E$8,\"H\")+COUNTIFS(D$3:D$8,\"=\"&Teams!A2,E$3:E$8,\"A\")";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGamesWonFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGamesLostFormula()
		{
			const string expected = "COUNTIFS(A$3:A$8,\"=\"&Teams!A2,E$3:E$8,\"A\")+COUNTIFS(D$3:D$8,\"=\"&Teams!A2,E$3:E$8,\"H\")";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGamesLostFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGamesDrawnFormula()
		{
			const string expected = "COUNTIFS(A$3:A$8,\"=\"&Teams!A2,E$3:E$8,\"D\")+COUNTIFS(D$3:D$8,\"=\"&Teams!A2,E$3:E$8,\"D\")";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGamesDrawnFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGamePointsFormula()
		{
			const string expected = "=(H3*3) + J3";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGamePointsFormula(START_GAMES_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestTeamRankFormula()
		{
			const string expected = "RANK(M3,M$3:M$18)";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetTeamRankFormula(START_GAMES_ROW, END_TEAMS_ROW);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGoalsScoredFormula()
		{
			const string expected = "SUMIFS(B$3:B$8, A$3:A$8,\"=\"&Teams!A2)+SUMIFS(C$3:C$8, D$3:D$8,\"=\"&Teams!A2)";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGoalsScoredFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGoalsAgainstFormula()
		{
			const string expected = "SUMIFS(C$3:C$8, A$3:A$8,\"=\"&Teams!A2)+SUMIFS(B$3:B$8, D$3:D$8,\"=\"&Teams!A2)";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGoalsAgainstFormula(START_GAMES_ROW, END_GAMES_ROW, FIRST_TEAM_CELL);
			Assert.Equal(expected, formula);
		}

		[Fact]
		public void TestGoalDifferentialFormula()
		{
			const string expected = "=O3 - P3";

			FormulaGenerator fg = GetFormulaGenerator();
			string formula = fg.GetGoalDifferentialFormula(START_GAMES_ROW);
			Assert.Equal(expected, formula);
		}
	}
}