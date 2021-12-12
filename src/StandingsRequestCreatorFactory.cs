using System.Collections.Generic;
using System.Linq;

namespace StandingsGoogleSheetsHelper
{
	/// <summary>
	/// Factory class to get the <see cref="IStandingsRequestCreator"/> instance that is applicable for a given column in the standings table
	/// </summary>
	public class StandingsRequestCreatorFactory
	{
		/// <summary>
		/// The collection <see cref="IStandingsRequestCreator"/> instances being used
		/// </summary>
		public IList<IStandingsRequestCreator> Creators { get; set; }

		public StandingsRequestCreatorFactory(IEnumerable<IStandingsRequestCreator> creators)
		{
			Creators = new List<IStandingsRequestCreator>(creators);
		}

		/// <summary>
		/// Gets the <see cref="IStandingsRequestCreator"/> from the <see cref="Creators"/> collection that is applicable to the given column header, or null if none found
		/// </summary>
		/// <param name="columnHeader"></param>
		/// <returns></returns>
		public IStandingsRequestCreator GetRequestCreator(string columnHeader) => Creators.SingleOrDefault(x => x.IsApplicableToColumn(columnHeader));
	}
}
