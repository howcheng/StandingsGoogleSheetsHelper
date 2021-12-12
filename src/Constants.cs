using System.Collections.Generic;

namespace StandingsGoogleSheetsHelper
{
	public static class Constants
	{
		public const char HOME_TEAM_INDICATOR = 'H';
		public const char AWAY_TEAM_INDICATOR = 'A';
		public const char WIN_INDICATOR = 'W';
		public const char LOSS_INDICATOR = 'L';
		public const char DRAW_INDICATOR = 'D';

		public const string HDR_HOME_TEAM = "HOME";
		public const string HDR_HOME_GOALS = "HG";
		public const string HDR_AWAY_GOALS = "AG";
		public const string HDR_AWAY_TEAM = "AWAY";
		public const string HDR_WINNING_TEAM = "WINNER";
		public const string HDR_TEAM_NAME = "TEAM";
		public const string HDR_TOTAL_PTS = "TOTAL";
		public const string HDR_RANK = "RANK";
		public const string HDR_GAMES_PLAYED = "GP";
		public const string HDR_NUM_WINS = "W";
		public const string HDR_NUM_LOSSES = "L";
		public const string HDR_NUM_DRAWS = "D";
		public const string HDR_GAME_PTS = "PTS";
		public const string HDR_GOALS_FOR = "GF";
		public const string HDR_GOALS_AGAINST = "GA";
		public const string HDR_GOAL_DIFF = "GD";
		// Core season specific
		public const string HDR_REF_PTS = "REF";
		public const string HDR_VOL_PTS = "VOL";
		public const string HDR_SPORTSMANSHIP_PTS = "SPT";
		public const string HDR_PTS_DEDUCTION = "DED";
		// Tournament specific
		public const string HDR_YELLOW_CARDS = "YC";
		public const string HDR_RED_CARDS = "RC";
		public const string HDR_HOME_PTS = "Pts (H)";
		public const string HDR_AWAY_PTS = "Pts (A)";
		public const string HDR_CALC_RANK = "C-RANK"; // "Calculated rank" -- rank based on the formula but doesn't take into account a manual tiebreaker
		public const string HDR_TIEBREAKER = "TB";

		/// <summary>
		/// The standard headers for the columns to enter game scores (in this order): home team, home score, away score, away team, game winner
		/// </summary>
		public static IEnumerable<string> GameScoreColumnHeaders = new[]
		{
			HDR_HOME_TEAM, HDR_HOME_GOALS, HDR_AWAY_GOALS, HDR_AWAY_TEAM, HDR_WINNING_TEAM
		};

		public const int WIDTH_WINNING_TEAM_COL = 65;
		public const int WIDTH_NUM_COL = 30;
		public const int WIDTH_WIDE_NUM_COL = 50;

		public const int ROUND_OFFSET_STANDINGS_TABLE = 2; // # of cells separating each round's scoring cells
	}
}
