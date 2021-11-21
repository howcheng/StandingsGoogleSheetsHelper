using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StandingsGoogleSheetsHelper
{
    public static class FormulaGenerator
    {
		/// <summary>
		/// Gets the formula for determining who won the game
		/// </summary>
		/// <param name="rowNum"></param>
		/// <param name="homeGoalsColumnName"></param>
		/// <param name="awayGoalsColumnName"></param>
		/// <returns></returns>
		/// <remarks>=IFS(OR(ISBLANK(B3),ISBLANK(C3)),"",B3=C3,"D",B3&gt;C3,"H",B3&lt;C3,"A")
		/// = if home or away goals scored is blank, then return blank
		/// otherwise if goals are equal return D, if home &gt; away return H, else return A
		/// </remarks>
		public static string GetGameWinnerFormula(int rowNum, string homeGoalsColumnName, string awayGoalsColumnName)
		{
			string homeGoalsCell = $"{homeGoalsColumnName}{rowNum}";
			string awayGoalsCell = $"{awayGoalsColumnName}{rowNum}";
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
		/// <param name="homeTeamsCellRange">Range of cells where the home team is listed (e.g., "A$21:A$28")</param>
		/// <param name="awayTeamsCellRange">Range of cells where the away team is listed (e.g., "D$21:D$28")</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2)+COUNTIFS(D$21:D$28,"="&Teams!A2)
		/// = count of number of times team name appears in Home column + same for Away column</remarks>
		public static string GetGamesPlayedFormula(string homeTeamsCellRange, string awayTeamsCellRange, string firstTeamCell)
		{
			return $"COUNTIFS({homeTeamsCellRange},\"=\"&{firstTeamCell})+COUNTIFS({awayTeamsCellRange},\"=\"&{firstTeamCell})";
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="homeTeamsCellRange">Range of cells where the home team is listed (e.g., "A$21:A$28")</param>
		/// <param name="awayTeamsCellRange">Range of cells where the away team is listed (e.g., "D$21:D$28")</param>
		/// <param name="whoWonCellRange">Range of cells where the winner is listed (e.g., "E$21:E$28")</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"H")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"A")
		/// = count number of times team name appears in Home column AND winning team = H + same for away column AND winning team = A</remarks>
		public static string GetGamesWonFormula(string homeTeamsCellRange, string awayTeamsCellRange, string whoWonCellRange, string firstTeamCell)
		{
			return $"COUNTIFS({homeTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.HOME_TEAM_INDICATOR}\")+COUNTIFS({awayTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.AWAY_TEAM_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="homeTeamsCellRange">Range of cells where the home team is listed (e.g., "A$21:A$28")</param>
		/// <param name="awayTeamsCellRange">Range of cells where the away team is listed (e.g., "D$21:D$28")</param>
		/// <param name="whoWonCellRange">Range of cells where the winner is listed (e.g., "E$21:E$28")</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"A")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"H")
		/// = count number of times team name appears in Home column AND winning team = A + same for away column AND winning team = H</remarks>
		public static string GetGamesLostFormula(string homeTeamsCellRange, string awayTeamsCellRange, string whoWonCellRange, string firstTeamCell)
		{
			return $"COUNTIFS({homeTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.AWAY_TEAM_INDICATOR}\")+COUNTIFS({awayTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.HOME_TEAM_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the number of games won
		/// </summary>
		/// <param name="homeTeamsCellRange">Range of cells where the home team is listed (e.g., "A$21:A$28")</param>
		/// <param name="awayTeamsCellRange">Range of cells where the away team is listed (e.g., "D$21:D$28")</param>
		/// <param name="whoWonCellRange">Range of cells where the winner is listed (e.g., "E$21:E$28")</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <returns></returns>
		/// <remarks>COUNTIFS(A$21:A$28,"="&Teams!A2,E$21:E$28,"")+COUNTIFS(D$21:D$28,"="&Teams!A2,E$21:E$28,"")
		/// = count number of times team name appears in Home column AND winning team = (empty) + same for away column</remarks>
		public static string GetGamesDrawnFormula(string homeTeamsCellRange, string awayTeamsCellRange, string whoWonCellRange, string firstTeamCell)
		{
			return $"COUNTIFS({homeTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.DRAW_INDICATOR}\")+COUNTIFS({awayTeamsCellRange},\"=\"&{firstTeamCell},{whoWonCellRange},\"{Constants.DRAW_INDICATOR}\")";
		}

		/// <summary>
		/// Gets the formula for determining the team rank
		/// </summary>
		/// <param name="columnName">Name of the column that contains the points value (e.g., 'M')</param>
		/// <param name="startRowNum">Row number of the first team</param>
		/// <param name="teamCount">Number of teams</param>
		/// <returns></returns>
		/// <remarks>RANK(M3,M$3:M$18)</remarks>
		public static string GetTeamRankFormula(string columnName, int startRowNum, int teamCount)
		{
			return $"RANK({columnName}{startRowNum},{columnName}${startRowNum}:{columnName}${startRowNum + teamCount - 1})";
		}

		/// <summary>
		/// Gets the formula for determining the number of goals scored
		/// </summary>
		/// <param name="homeTeamsCellRange">Range of cells where the home team is listed (e.g., "A$21:A$28")</param>
		/// <param name="firstTeamCell">Cell where the first team is listed (e.g., "Teams!A2")</param>
		/// <param name="awayTeamsCellRange">Range of cells where the away team is listed (e.g., "D$21:D$28")</param>
		/// <param name="homeGoalsCellRange">Range of cells for the goals scored by the home team</param>
		/// <param name="awayGoalsCellRange">Range of cells for the goals scored by the away team</param>
		/// <param name="goalsFor">Flag indicating we are calculating goals for (set to false for goals against)</param>
		/// <returns></returns>
		/// <remarks>SUMIFS(B$21:B$28, A$21:A$28,"="&Teams!A2)+SUMIFS(C$21:C$28, D$21:D$28,"="&Teams!A2)
		/// = sum of home goals column where home team = team name + sum of away goals column where away team = team name
		/// When doing goals against, swap the home and away goal columns</remarks>
		public static string GetGoalsScoredFormula(string homeTeamsCellRange, string firstTeamCell, string awayTeamsCellRange, string homeGoalsCellRange, string awayGoalsCellRange, bool goalsFor)
		{
			return $"SUMIFS({(goalsFor ? homeGoalsCellRange : awayGoalsCellRange)}, {homeTeamsCellRange},\"=\"&{firstTeamCell})+SUMIFS({(goalsFor ? awayGoalsCellRange : homeGoalsCellRange)}, {awayTeamsCellRange},\"=\"&{firstTeamCell})";
		}

		public static string GetGoalDifferentialFormula(string goalsForStartCell, string goalsAgainstStartCell)
		{
			return $"{goalsForStartCell} - {goalsAgainstStartCell}";
		}
    }
}
