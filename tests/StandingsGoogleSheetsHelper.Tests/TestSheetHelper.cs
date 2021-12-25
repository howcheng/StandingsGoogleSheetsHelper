using System.Collections.Generic;

namespace StandingsGoogleSheetsHelper.Tests
{
	public class TestSheetHelper : StandingsSheetHelper
	{
		private static List<string> headerColumns;
		private static List<string> standingsTableColumns;

		static TestSheetHelper()
		{
			// NOTE: These columns are the same as in the 2021 Region 42 score sheet; if you change these, the FormulaGeneratorTests will fail
			headerColumns = new List<string>
				{
					Constants.HDR_HOME_TEAM, Constants.HDR_HOME_GOALS, Constants.HDR_AWAY_GOALS, Constants.HDR_AWAY_TEAM, Constants.HDR_WINNING_TEAM
				};
			standingsTableColumns = new List<string>
				{
					Constants.HDR_TEAM_NAME, Constants.HDR_GAMES_PLAYED, Constants.HDR_NUM_WINS, Constants.HDR_NUM_LOSSES, Constants.HDR_NUM_DRAWS,
					Constants.HDR_GAME_PTS, Constants.HDR_REF_PTS, Constants.HDR_TOTAL_PTS, Constants.HDR_RANK,
					Constants.HDR_GOALS_FOR, Constants.HDR_GOALS_AGAINST, Constants.HDR_GOAL_DIFF
				};
			headerColumns.AddRange(standingsTableColumns);
		}

		public TestSheetHelper() : base(headerColumns, standingsTableColumns)
		{
		}
	}
}