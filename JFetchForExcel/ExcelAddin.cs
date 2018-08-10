using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;
using System.Net.Http;
using JFetch;

namespace JFetch {
	public static class ExcelAddin {

		private static HttpClient client = new HttpClient();

		[ExcelFunction(Description = "Print info about Kings")]
		public static object GetKings() {
			return ExcelAsyncUtil.Run("GetKings", new object[] { }, () => JFetch.JFetchSync("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", client));
		}

	}
}