using System.Collections.Generic;

namespace StandingsGoogleSheetsHelper
{
	public interface IStandingsRequestCreatorFactory
	{
		IList<IStandingsRequestCreator> Creators { get; set; }

		IStandingsRequestCreator GetRequestCreator(string columnHeader);
	}
}