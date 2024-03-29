﻿using System.Collections.Generic;
using System.Linq;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace StandingsGoogleSheetsHelper
{
	/// <summary>
	/// Base class for helper methods to make sheets that show team standings
	/// </summary>
	public class StandingsSheetHelper : SheetHelper
	{
		/// <summary>
		/// A collection of the header columns used to make up the standings table
		/// </summary>
		public List<string> StandingsTableColumns { get; } = new List<string>();

		public StandingsSheetHelper(IEnumerable<string> headerColumns)
			: this(headerColumns, null)
		{
		}

		public StandingsSheetHelper(IEnumerable<string> headerColumns, IEnumerable<string> standingsTableColumns)
			: base(headerColumns)
		{
			if ((standingsTableColumns?.Count() ?? 0) > 0)
				StandingsTableColumns.AddRange(standingsTableColumns);		
		}

		public override int GetColumnIndexByHeader(string colHeader)
		{
			int idx = base.GetColumnIndexByHeader(colHeader);
			if (idx > -1)
				return idx;

			if ((StandingsTableColumns?.Count() ?? 0) > 0)
			{
				idx = StandingsTableColumns.IndexOf(colHeader);
				if (idx > -1)
					idx += HeaderRowColumns.Count;
			}
			return idx;
		}

		public string HomeTeamColumnName => GetColumnNameByHeader(Constants.HDR_HOME_TEAM);
		public string HomeGoalsColumnName => GetColumnNameByHeader(Constants.HDR_HOME_GOALS);
		public string AwayGoalsColumnName => GetColumnNameByHeader(Constants.HDR_AWAY_GOALS);
		public string AwayTeamColumnName => GetColumnNameByHeader(Constants.HDR_AWAY_TEAM);
		public string ForfeitColumnName => GetColumnNameByHeader(Constants.HDR_FORFEIT);
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
		public IEnumerable<Request> CreateCellWidthRequests(int? sheetId, int teamNameColumnWidth) => CreateCellWidthRequests(sheetId, HeaderRowColumns, teamNameColumnWidth);

		public IEnumerable<Request> CreateCellWidthRequests(int? sheetId, IEnumerable<string> columnHeaders, int teamNameColumnWidth)
		{
			List<Request> requests = new List<Request>();
			foreach (string header in columnHeaders)
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
					case Constants.HDR_TIEBREAKER_H2H:
					case Constants.HDR_FORFEIT:
					case Constants.HDR_TIEBREAKER_GOALS_AGAINST_HOME:
					case Constants.HDR_TIEBREAKER_GOALS_AGAINST_AWAY:
					case Constants.HDR_TIEBREAKER_GOALS_FOR_HOME:
					case Constants.HDR_TIEBREAKER_GOALS_FOR_AWAY:
					case Constants.HDR_TIEBREAKER_WINS:
					case Constants.HDR_TIEBREAKER_CARDS:
					case Constants.HDR_TIEBREAKER_GOALS_AGAINST:
					case Constants.HDR_TIEBREAKER_GOAL_DIFF:
					case Constants.HDR_TIEBREAKER_KFTM_WINNER:
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

				int colIndex = GetColumnIndexByHeader(header);
				Request request = RequestCreator.CreateCellWidthRequest(sheetId, colWidth.Value, colIndex);
				requests.Add(request);
			}
			return requests;
		}
	}
}
