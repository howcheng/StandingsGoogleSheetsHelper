﻿using System;
using System.Collections.Generic;
using System.Text;
using Google.Apis.Sheets.v4.Data;
using GoogleSheetsHelper;

namespace StandingsGoogleSheetsHelper
{
	/// <summary>
	/// Interface for classes that create Google Sheets 
	/// </summary>
	public interface IStandingsRequestCreator
	{
		/// <summary>
		/// Creates a <see cref="Request"/> to update a sheet with a repeated formula to create the standings table
		/// </summary>
		/// <returns></returns>
		Request CreateRequest(StandingsRequestCreatorConfig config);

		/// <summary>
		/// Determines if the instance of the class can be applied for the given column
		/// </summary>
		/// <param name="columnHeader"></param>
		/// <returns></returns>
		bool IsApplicableToColumn(string columnHeader);

		string ColumnHeader { get; }
	}

	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column in the standings table
	/// </summary>
	public abstract class StandingsRequestCreator
	{
		protected readonly FormulaGenerator _formulaGenerator;
		protected readonly string _columnHeader;
		protected readonly string _columnName;
		protected readonly int _columnIndex;

		protected StandingsRequestCreator(FormulaGenerator formGen, string columnHeader)
		{
			_formulaGenerator = formGen;
			_columnHeader = columnHeader;
			_columnName = _formulaGenerator.SheetHelper.GetColumnNameByHeader(columnHeader);
			if (columnHeader == null)
				throw new ArgumentException($"Can't find column '{columnHeader}' the collection. Did you forget to add it to {nameof(StandingsSheetHelper)}.{nameof(StandingsSheetHelper.HeaderRowColumns)}?", nameof(columnHeader));
			_columnIndex = _formulaGenerator.SheetHelper.GetColumnIndexByHeader(columnHeader);
		}

		public bool IsApplicableToColumn(string columnHeader) => columnHeader == _columnHeader;

		public string ColumnHeader => _columnHeader;

		/// <summary>
		/// After round 1, we must add the previous week's total to the current round's. Applies to games played, wins, losses, draws, goals for/against, ref/volunteer/etc points (if not stored cumulatively)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns></returns>
		protected string GetAddLastRoundValueFormula(int rowNum) => rowNum == 0 ? string.Empty : $"+{_columnName}{rowNum}";
	}

	/// <summary>
	/// Common configuration for all <see cref="StandingsRequestCreator"/> classes
	/// </summary>
	public class StandingsRequestCreatorConfig
	{
		/// <summary>
		/// The sheet ID (from <see cref="Sheet.Properties.SheetId"/>)
		/// </summary>
		public int? SheetId { get; set; }
		/// <summary>
		/// The index of the first row in the standings table
		/// </summary>
		public int SheetStartRowIndex { get; set; }
		/// <summary>
		/// The row number of the first row where game scores are entered (usually will be <see cref="SheetStartRowIndex"/> + 1)
		/// </summary>
		public int StartGamesRowNum { get; set; }
		/// <summary>
		/// The number of teams in the current division
		/// </summary>
		public int NumTeams { get; set; }
	}

	/// <summary>
	/// Common configuration for all <see cref="ScoreBasedStandingsRequestCreator"/> classes
	/// </summary>
	public class ScoreBasedStandingsRequestCreatorConfig : StandingsRequestCreatorConfig
	{
		/// <summary>
		/// The row number of the last row where game scores are entered
		/// </summary>
		public int EndGamesRowNum { get; set; }
		/// <summary>
		/// The cell reference for the cell where the name of the first team in the division is stored (e.g., "Teams!A2")
		/// </summary>
		public string FirstTeamsSheetCell { get; set; }
		/// <summary>
		/// The row number of the first row in the standings table of the previous round (0 when it's round 1)
		/// </summary>
		public int LastRoundStartRowNum { get; set; }
		/// <summary>
		/// Indicates that games from the current round are used to calculate standings (when <c>false</c>, all games are treated as scrimmages)
		/// </summary>
		public bool RoundCountsForStandings { get; set; } = true;
	}

	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column based on the game results in the standings table (e.g., goals scored/conceded)
	/// </summary>
	public abstract class ScoreBasedStandingsRequestCreator : StandingsRequestCreator
	{
		private ScoreBasedStandingsRequestCreatorConfig _config;
		private Func<int, int, string, string> _formulaGeneratorMethod;

		protected ScoreBasedStandingsRequestCreator(FormulaGenerator formGen, string columnHeader, Func<int, int, string, string> formulaGeneratorMethod)
			: base(formGen, columnHeader)
		{
			_formulaGeneratorMethod = formulaGeneratorMethod;
		}

		public virtual Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			_config = (ScoreBasedStandingsRequestCreatorConfig)config;
			string addLastRoundValueFormula = GetAddLastRoundValueFormula(_config.LastRoundStartRowNum);
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(_config.SheetId, _config.SheetStartRowIndex, _columnIndex, _config.NumTeams,
				$"={_formulaGeneratorMethod(_config.StartGamesRowNum, _config.EndGamesRowNum, _config.FirstTeamsSheetCell)}{addLastRoundValueFormula}");
			return request;
		}
	}

	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column based on the game results in the standings table
	/// where the game may or may not be a scrimmage (e.g., games played, wins)
	/// </summary>
	public abstract class ScrimmageBasedStandingsRequestCreator : ScoreBasedStandingsRequestCreator
	{
		private ScoreBasedStandingsRequestCreatorConfig _config;

		protected ScrimmageBasedStandingsRequestCreator(FormulaGenerator formGen, string columnHeader, Func<int, int, string, string> formulaGeneratorMethod)
			: base(formGen, columnHeader, formulaGeneratorMethod)
		{
		}

		public override Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			_config = (ScoreBasedStandingsRequestCreatorConfig)config;
			if (_config.RoundCountsForStandings)
				return base.CreateRequest(config);

			// when it's the game doesn't count for standings, enter a zero
			return new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = _config.SheetId,
						StartRowIndex = _config.SheetStartRowIndex,
						StartColumnIndex = _columnIndex,
						EndRowIndex = _config.SheetStartRowIndex + _config.NumTeams,
						EndColumnIndex = _columnIndex + 1,
					},
					Cell = new CellData
					{
						 UserEnteredValue = new ExtendedValue
						 {
							 NumberValue = 0,
						 }
					},
					Fields = nameof(CellData.UserEnteredValue).ToCamelCase(),
				},
			};
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for determining which team won
	/// </summary>
	public class GameWinnerRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public GameWinnerRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_WINNING_TEAM)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.NumTeams,
				_formulaGenerator.GetGameWinnerFormula(config.StartGamesRowNum));
			return request;
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for number of games played
	/// </summary>
	public class GamesPlayedRequestCreator : ScrimmageBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GamesPlayedRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_GAMES_PLAYED, formGen.GetGamesPlayedFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for number of games won
	/// </summary>
	public class GamesWonRequestCreator : ScrimmageBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GamesWonRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_NUM_WINS, formGen.GetGamesWonFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for number of games lost
	/// </summary>
	public class GamesLostRequestCreator : ScrimmageBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GamesLostRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_NUM_LOSSES, formGen.GetGamesLostFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for number of games drawn (tied)
	/// </summary>
	public class GamesDrawnRequestCreator : ScrimmageBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GamesDrawnRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_NUM_DRAWS, formGen.GetGamesDrawnFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for game points, using the standard formula of 3 points for a win and 1 point for a draw
	/// </summary>
	public class GamePointsRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public GamePointsRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_GAME_PTS)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.NumTeams,
				_formulaGenerator.GetGamePointsFormula(config.StartGamesRowNum));
			return request;
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for team rank
	/// </summary>
	public class TeamRankRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public TeamRankRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_RANK)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.NumTeams,
				_formulaGenerator.GetTeamRankFormula(config.SheetStartRowIndex + 1, config.SheetStartRowIndex + config.NumTeams));
			return request;
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for goals scored
	/// </summary>
	public class GoalsScoredRequestCreator : ScoreBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GoalsScoredRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_GOALS_FOR, formGen.GetGoalsScoredFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for goals conceded
	/// </summary>
	public class GoalsAgainstRequestCreator : ScoreBasedStandingsRequestCreator, IStandingsRequestCreator
	{
		public GoalsAgainstRequestCreator(FormulaGenerator formGen)
			: base(formGen, Constants.HDR_GOALS_AGAINST, formGen.GetGoalsAgainstFormula)
		{
		}
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for goal differential
	/// </summary>
	public class GoalDifferentialRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		public GoalDifferentialRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_GOAL_DIFF)
		{
		}

		public Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.NumTeams,
				_formulaGenerator.GetGoalDifferentialFormula(config.StartGamesRowNum));
			return request;
		}
	}
}
