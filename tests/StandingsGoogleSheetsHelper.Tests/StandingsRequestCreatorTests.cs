using AutoFixture;
using Google.Apis.Sheets.v4.Data;
using Xunit;

namespace StandingsGoogleSheetsHelper.Tests
{
	// NOTE: not valdating the actual formulas here; that's done in the FormulaGeneratorTests
	public class StandingsRequestCreatorTests
	{
		private StandingsRequestCreatorConfig GetConfig() => new Fixture().Create<StandingsRequestCreatorConfig>();
		private ScoreBasedStandingsRequestCreatorConfig GetScoreBasedConfig() => new Fixture().Create<ScoreBasedStandingsRequestCreatorConfig>();

		private TestSheetHelper _helper = new TestSheetHelper();
		private FormulaGenerator GetFormulaGenerator() => new FormulaGenerator(_helper);

		private void ValidateRequest(Request request, StandingsRequestCreatorConfig config, string columnHeader)
		{
			Assert.Equal(config.SheetId, request.RepeatCell.Range.SheetId);
			Assert.Equal(config.SheetStartRowIndex, request.RepeatCell.Range.StartRowIndex);
			Assert.Equal(config.SheetStartRowIndex + config.NumTeams, request.RepeatCell.Range.EndRowIndex);
			int columnIndex = _helper.GetColumnIndexByHeader(columnHeader);
			Assert.Equal(columnIndex, request.RepeatCell.Range.StartColumnIndex);
		}

		private void ValidateFormula(Request request, StandingsRequestCreatorConfig config, string formula, string columnHeader)
		{
			ValidateRequest(request, config, columnHeader);
			Assert.NotNull(request.RepeatCell.Cell.UserEnteredValue);
			Assert.Contains(formula, request.RepeatCell.Cell.UserEnteredValue.FormulaValue); // Contains because some of the formulas may not include the = sign
		}

		private void ValidateZeroValue(Request request, StandingsRequestCreatorConfig config, string columnHeader)
		{
			ValidateRequest(request, config, columnHeader);
			Assert.NotNull(request.RepeatCell.Cell.UserEnteredValue);
			Assert.Equal(0, request.RepeatCell.Cell.UserEnteredValue.NumberValue);
		}

		[Theory]
		[InlineData(true)]
		[InlineData(false)]
		public void TestGamesPlayedRequestCreator(bool roundCountsForStandings)
		{
			ScoreBasedStandingsRequestCreatorConfig config = GetScoreBasedConfig();
			config.RoundCountsForStandings = roundCountsForStandings;

			FormulaGenerator fg = GetFormulaGenerator();
			GamesPlayedRequestCreator creator = new GamesPlayedRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			if (roundCountsForStandings)
			{
				string formula = fg.GetGamesPlayedFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
				ValidateFormula(request, config, formula, Constants.HDR_GAMES_PLAYED);
			}
			else
				ValidateZeroValue(request, config, Constants.HDR_GAMES_PLAYED);
		}

		[Theory]
		[InlineData(true)]
		[InlineData(false)]
		public void TestGamesWonRequestCreator(bool roundCountsForStandings)
		{
			ScoreBasedStandingsRequestCreatorConfig config = GetScoreBasedConfig();
			config.RoundCountsForStandings = roundCountsForStandings;

			FormulaGenerator fg = GetFormulaGenerator();
			GamesWonRequestCreator creator = new GamesWonRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			if (roundCountsForStandings)
			{
				string formula = fg.GetGamesWonFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
				ValidateFormula(request, config, formula, Constants.HDR_NUM_WINS);
			}
			else
				ValidateZeroValue(request, config, Constants.HDR_NUM_WINS);
		}

		[Theory]
		[InlineData(true)]
		[InlineData(false)]
		public void TestGamesLostRequestCreator(bool roundCountsForStandings)
		{
			ScoreBasedStandingsRequestCreatorConfig config = GetScoreBasedConfig();
			config.RoundCountsForStandings = roundCountsForStandings;

			FormulaGenerator fg = GetFormulaGenerator();
			GamesLostRequestCreator creator = new GamesLostRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			if (roundCountsForStandings)
			{
				string formula = fg.GetGamesLostFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
				ValidateFormula(request, config, formula, Constants.HDR_NUM_LOSSES);
			}
			else
				ValidateZeroValue(request, config, Constants.HDR_NUM_LOSSES);
		}

		[Theory]
		[InlineData(true)]
		[InlineData(false)]
		public void TestGamesDrawnRequestCreator(bool roundCountsForStandings)
		{
			ScoreBasedStandingsRequestCreatorConfig config = GetScoreBasedConfig();
			config.RoundCountsForStandings = roundCountsForStandings;

			FormulaGenerator fg = GetFormulaGenerator();
			GamesDrawnRequestCreator creator = new GamesDrawnRequestCreator(fg);
			Request request = creator.CreateRequest(config);

			if (roundCountsForStandings)
			{
				string formula = fg.GetGamesDrawnFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
				ValidateFormula(request, config, formula, Constants.HDR_NUM_DRAWS);
			}
			else
				ValidateZeroValue(request, config, Constants.HDR_NUM_DRAWS);
		}

		[Fact]
		public void TestGamePointsRequestCreator()
		{
			StandingsRequestCreatorConfig config = GetConfig();

			FormulaGenerator fg = GetFormulaGenerator();
			GamePointsRequestCreator creator = new GamePointsRequestCreator(fg);
			Request request = creator.CreateRequest(config);
			string formula = fg.GetGamePointsFormula(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, Constants.HDR_GAME_PTS);
		}

		[Fact]
		public void TestTeamRankRequestCreator()
		{
			StandingsRequestCreatorConfig config = GetConfig();

			FormulaGenerator fg = GetFormulaGenerator();
			TeamRankRequestCreator creator = new TeamRankRequestCreator(fg);
			Request request = creator.CreateRequest(config);
			string formula = fg.GetTeamRankFormula(config.SheetStartRowIndex + 1, config.SheetStartRowIndex + config.NumTeams);
			ValidateFormula(request, config, formula, Constants.HDR_RANK);
		}

		[Fact]
		public void TestGoalsScoredRequestCreator()
		{
			ScoreBasedStandingsRequestCreatorConfig? config = GetScoreBasedConfig();

			FormulaGenerator fg = GetFormulaGenerator();
			GoalsScoredRequestCreator creator = new GoalsScoredRequestCreator(fg);
			Request request = creator.CreateRequest(config);
			string formula = fg.GetGoalsScoredFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
			ValidateFormula(request, config, formula, Constants.HDR_GOALS_FOR);
		}

		[Fact]
		public void TestGoalsAgainstRequestCreator()
		{
			ScoreBasedStandingsRequestCreatorConfig? config = GetScoreBasedConfig();

			FormulaGenerator fg = GetFormulaGenerator();
			GoalsAgainstRequestCreator creator = new GoalsAgainstRequestCreator(fg);
			Request request = creator.CreateRequest(config);
			string formula = fg.GetGoalsAgainstFormula(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell);
			ValidateFormula(request, config, formula, Constants.HDR_GOALS_AGAINST);
		}

		[Fact]
		public void TestGoalDifferentialRequestCreator()
		{
			StandingsRequestCreatorConfig config = GetConfig();

			FormulaGenerator fg = GetFormulaGenerator();
			GoalDifferentialRequestCreator creator = new GoalDifferentialRequestCreator(fg);
			Request request = creator.CreateRequest(config);
			string formula = fg.GetGoalDifferentialFormula(config.StartGamesRowNum);
			ValidateFormula(request, config, formula, Constants.HDR_GOAL_DIFF);
		}
	}
}
