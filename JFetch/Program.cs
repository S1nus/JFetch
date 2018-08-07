using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using System.ComponentModel;

namespace JFetch {
	
	class Program {

		public static void DoWithResponse (HttpWebRequest request, object otherstuff, Action<HttpWebResponse, object> responseAction) {
			Action wrapperAction = () => {
				request.BeginGetResponse(new AsyncCallback((iar) => {
					var response = (HttpWebResponse)((HttpWebRequest)iar.AsyncState).EndGetResponse(iar);
					responseAction(response, otherstuff);
				}), request);
			};
			wrapperAction.BeginInvoke(new AsyncCallback((iar) => {
				var action = (Action)iar.AsyncState;
				action.EndInvoke(iar);
			}), wrapperAction);
		}

		//public static void SetJsonDict(HttpWebResponse response, out object target) {
		//	var body = new StreamReader(response.GetResponseStream()).ReadToEnd();
		//	target = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);
		//}

		public static void JFetch(HttpWebRequest request, out object target) {
			object toRet = null;
			DoWithResponse(request, toRet, (response, otherstuff) => {
				var body = new StreamReader(response.GetResponseStream()).ReadToEnd();
				otherstuff = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);
			});
			target = toRet;
			Console.WriteLine(toRet);
		}

		static void Main(string[] args) {

			//http://mysafeinfo.com/api/data?list=englishmonarchs&format=json
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json");
			object table = null;
			JFetch(request, out table);
			Console.Read();

		}

	}
}