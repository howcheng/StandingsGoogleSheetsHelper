using System;
using System.Collections.Generic;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace StandingsGoogleSheetsHelper
{
	/// <summary>
	/// Base class for helper methods to make sheets that show team standings
	/// </summary>
	public abstract class StandingsSheetHelper
	{
		/// <summary>
		/// A collection of all the header columns (including the ones in <see cref="StandingsTableColumns"/>)
		/// </summary>
		public List<string> HeaderRowColumns { get; private set; }
		/// <summary>
		/// A collection of the header columns used to make up the standings table
		/// </summary>
		public List<string> StandingsTableColumns { get; private set; }

		protected StandingsSheetHelper(IEnumerable<string> headerColumns, IEnumerable<string> standingsTableColumns)
		{
			HeaderRowColumns = new List<string>(headerColumns);
			StandingsTableColumns = new List<string>(standingsTableColumns);
		}

		/// <summary>
		/// Gets the column index (zero-based) by header value (from <see cref="Constants"/>).
		/// </summary>
		/// <param name="colHeader"></param>
		/// <returns>For example, <see cref="Constants.HDR_HOME_TEAM"/> returns 0 because it's the first column. If the column isn't being used, then returns -1.</returns>
		public virtual int GetColumnIndexByHeader(string colHeader)
		{
			int idx = HeaderRowColumns.IndexOf(colHeader);
			return idx;
		}

		/// <summary>
		/// Gets the column name by header value (from <see cref="Constants"/>).
		/// </summary>
		/// <param name="colHeader"></param>
		/// <returns>For example, <see cref="Constants.HDR_HOME_TEAM"/> returns "A" because it's the first column. If the column isn't being used, then returns null.</returns>
		public string GetColumnNameByHeader(string colHeader)
		{
			int idx = GetColumnIndexByHeader(colHeader);
			if (idx == -1)
				return null;
			return Utilities.ConvertIndexToColumnName(idx);
		}

		public string HomeTeamColumnName => GetColumnNameByHeader(Constants.HDR_HOME_TEAM);
		public string HomeGoalsColumnName => GetColumnNameByHeader(Constants.HDR_HOME_GOALS);
		public string AwayGoalsColumnName => GetColumnNameByHeader(Constants.HDR_AWAY_GOALS);
		public string AwayTeamColumnName => GetColumnNameByHeader(Constants.HDR_AWAY_TEAM);
		public string TeamNameColumnName => GetColumnNameByHeader(Constants.HDR_TEAM_NAME);
		public string GamesPlayedColumnName => GetColumnNameByHeader(Constants.HDR_GAMES_PLAYED);
		public string NumWinsColumnName => GetColumnNameByHeader(Constants.HDR_NUM_WINS);
		public string NumLossesColumnName => GetColumnNameByHeader(Constants.HDR_NUM_LOSSES);
		public string NumDrawsColumnName => GetColumnNameByHeader(Constants.HDR_NUM_DRAWS);
		public string GamePointsColumnName => GetColumnNameByHeader(Constants.HDR_GAME_PTS);
		public string TotalPointsColumnName => GetColumnNameByHeader(Constants.HDR_TOTAL_PTS);
		public string RankColumnName => GetColumnNameByHeader(Constants.HDR_RANK);
		public string WinnerColumnName => GetColumnNameByHeader(Constants.HDR_WINNING_TEAM);
		public string GoalsForColumnName => GetColumnNameByHeader(Constants.HDR_GOALS_FOR);
		public string GoalsAgainstColumnName => GetColumnNameByHeader(Constants.HDR_GOALS_AGAINST);
		public string GoalDifferentialColumnName => GetColumnNameByHeader(Constants.HDR_GOAL_DIFF);

		/// <summary>
		/// Creates <see cref="Request"/>s to resize the columns of the scores and standings sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="teamNameColumnWidth">The width of the team name column (after having used <see cref="SheetsClient.AutoResizeColumn(string, int)"/>, as this depends on the longest team name)</param>
		/// <returns></returns>
		public IEnumerable<Request> CreateCellWidthRequests(Sheet sheet, int teamNameColumnWidth)
		{
			List<Request> requests = new List<Request>();
			foreach (string header in HeaderRowColumns)
			{
				int? colWidth = null;
				switch (header)
				{
					case Constants.HDR_HOME_TEAM:
					case Constants.HDR_AWAY_TEAM:
					case Constants.HDR_TEAM_NAME:
						colWidth = teamNameColumnWidth;
						break;
					case Constants.HDR_WINNING_TEAM:
						// slightly wider than below because the column name is longer
						colWidth = Constants.WIDTH_WINNING_TEAM_COL;
						break;
					case Constants.HDR_TOTAL_PTS:
					case Constants.HDR_RANK:
					case Constants.HDR_CALC_RANK:
					case Constants.HDR_HOME_PTS:
					case Constants.HDR_AWAY_PTS:
						// slightly wider than below because the column name is longer
						colWidth = Constants.WIDTH_WIDE_NUM_COL;
						break;
					default:
						// columns that only hold numbers
						colWidth = Constants.WIDTH_NUM_COL;
						break;
				}
				if (!colWidth.HasValue)
					continue;

				int colIndex = HeaderRowColumns.IndexOf(header);
				Request request = RequestCreator.CreateCellWidthRequest(sheet.Properties.SheetId, colWidth.Value, colIndex);
				requests.Add(request);
			}
			return requests;
		}
	}
}
