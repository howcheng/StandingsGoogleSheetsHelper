using System;
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

		/// <summary>
		/// Gets the column header value that is applicable to this instance
		/// </summary>
		string ColumnHeader { get; }
	}

	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column in the standings table
	/// </summary>
	public abstract class StandingsRequestCreator : IStandingsRequestCreator
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

		public virtual bool IsApplicableToColumn(string columnHeader) => columnHeader == _columnHeader;

		public string ColumnHeader => _columnHeader;

		/// <summary>
		/// After round 1, we must add the previous week's total to the current round's. Applies to games played, wins, losses, draws, goals for/against, ref/volunteer/etc points (if not stored cumulatively)
		/// </summary>
		/// <param name="rowNum"></param>
		/// <returns></returns>
		protected string GetAddLastRoundValueFormula(int rowNum) => rowNum == 0 ? string.Empty : $"+{_columnName}{rowNum}";

		protected abstract string GenerateFormula(StandingsRequestCreatorConfig config);

		public virtual Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = RequestCreator.CreateRepeatedSheetFormulaRequest(config.SheetId, config.SheetStartRowIndex, _columnIndex, config.RowCount,
				GenerateFormula(config));
			return request;
		}
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
		/// The number of rows affected by this request (e.g., the number of game rows, or the number of teams)
		/// </summary>
		public int RowCount { get; set; }
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
		private Func<int, int, string, string> _formulaGeneratorMethod;

		protected ScoreBasedStandingsRequestCreator(FormulaGenerator formGen, string columnHeader, Func<int, int, string, string> formulaGeneratorMethod)
			: base(formGen, columnHeader)
		{
			_formulaGeneratorMethod = formulaGeneratorMethod;
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig cfg)
		{
			ScoreBasedStandingsRequestCreatorConfig config = (ScoreBasedStandingsRequestCreatorConfig)cfg;
			string addLastRoundValueFormula = GetAddLastRoundValueFormula(config.LastRoundStartRowNum);
			return $"={_formulaGeneratorMethod(config.StartGamesRowNum, config.EndGamesRowNum, config.FirstTeamsSheetCell)}{addLastRoundValueFormula}";
		}
	}

	/// <summary>
	/// Base class for classes that create a <see cref="Request"/> to build a column based on the game results in the standings table
	/// where the game may or may not be a scrimmage (e.g., games played, wins)
	/// </summary>
	public abstract class ScrimmageBasedStandingsRequestCreator : ScoreBasedStandingsRequestCreator
	{
		protected ScrimmageBasedStandingsRequestCreator(FormulaGenerator formGen, string columnHeader, Func<int, int, string, string> formulaGeneratorMethod)
			: base(formGen, columnHeader, formulaGeneratorMethod)
		{
		}

		public override Request CreateRequest(StandingsRequestCreatorConfig cfg)
		{
			ScoreBasedStandingsRequestCreatorConfig config = (ScoreBasedStandingsRequestCreatorConfig)cfg;
			if (config.RoundCountsForStandings)
				return base.CreateRequest(cfg);

			// when it's the game doesn't count for standings, enter a zero
			return new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = config.SheetId,
						StartRowIndex = config.SheetStartRowIndex,
						StartColumnIndex = _columnIndex,
						EndRowIndex = config.SheetStartRowIndex + config.RowCount,
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

		protected override string GenerateFormula(StandingsRequestCreatorConfig config)
			=> _formulaGenerator.GetGameWinnerFormula(config.StartGamesRowNum);
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

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> _formulaGenerator.GetGamePointsFormula(config.StartGamesRowNum);
	}

	/// <summary>
	/// Base class for creating a <see cref="Request"/> for building the column for rank
	/// </summary>
	public abstract class RankRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		protected RankRequestCreator(FormulaGenerator formGen, string columnHeader) : base(formGen, columnHeader)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) 
			=> $"={_formulaGenerator.GetTeamRankFormula(config.SheetStartRowIndex + 1, config.SheetStartRowIndex + config.RowCount)}";
	}

	/// <summary>
	/// Creates a <see cref="Request"/> for building the column for team rank
	/// </summary>
	public class TeamRankRequestCreator : RankRequestCreator
	{
		public TeamRankRequestCreator(FormulaGenerator formGen) 
			: base(formGen, Constants.HDR_RANK)
		{
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

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) =>_formulaGenerator.GetGoalDifferentialFormula(config.StartGamesRowNum);
	}

	public abstract class CheckboxRequestCreator : StandingsRequestCreator, IStandingsRequestCreator
	{
		protected CheckboxRequestCreator(FormulaGenerator formGen, string columnHeader) 
			: base(formGen, columnHeader)
		{
		}

		protected override string GenerateFormula(StandingsRequestCreatorConfig config) => throw new NotImplementedException();

		public override Request CreateRequest(StandingsRequestCreatorConfig config)
		{
			Request request = new Request
			{
				RepeatCell = new RepeatCellRequest
				{
					Range = new GridRange
					{
						SheetId = config.SheetId,
						StartRowIndex = config.SheetStartRowIndex,
						EndRowIndex = config.SheetStartRowIndex + config.RowCount,
						StartColumnIndex = _columnIndex,
						EndColumnIndex = _columnIndex + 1,
					},
					Cell = new CellData
					{
						DataValidation = new DataValidationRule
						{
							Condition = new BooleanCondition
							{
								Type = "BOOLEAN"
							}
						},
						UserEnteredValue = new ExtendedValue
						{
							BoolValue = false, // default to unchecked (if we don't set, it will be null, which causes weird things when sorting)
						}
					},
					Fields = nameof(CellData.DataValidation).ToCamelCase(),
				},
			};
			return request;
		}
	}
}
