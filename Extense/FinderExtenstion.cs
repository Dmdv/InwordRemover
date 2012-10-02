using System;
using System.Diagnostics;
using Microsoft.Office.Interop.Word;

namespace Extense
{
	public static class FinderExtenstion
	{
		public static int ReplaceAll(this Find finder, Action func = null)
		{
			var count = 0;
			while (Execute(finder, ref count))
			{
				if (func != null)
					func();
			}
			return count;
		}

		private static bool Execute(Find finder, ref int count)
		{
			try
			{
				var result = finder.Execute(Replace: WdReplace.wdReplaceOne);
				if (result) count++;
				return result;
			}
			catch (Exception ex)
			{
				// ¬ыводить пользователю или записывать в лог.
				Trace.WriteLine(ex.Message);
				return true;
			}
		}
	}
}