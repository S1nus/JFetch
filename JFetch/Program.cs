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

		public static void DoWithResponse (HttpWebRequest request, Action<HttpWebResponse> responseAction) {
			Action wrapperAction = () => {
				request.BeginGetResponse(new AsyncCallback((iar) => {
					var response = (HttpWebResponse)((HttpWebRequest)iar.AsyncState).EndGetResponse(iar);
					responseAction(response);
				}), request);
			};
			wrapperAction.BeginInvoke(new AsyncCallback((iar) => {
				var action = (Action)iar.AsyncState;
				action.EndInvoke(iar);
			}), wrapperAction);
		}

		public static void JFetch(string url, List<Dictionary<string, string>> result) {
			HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

			DoWithResponse(request, (response) => {
				var body = new StreamReader(response.GetResponseStream()).ReadToEnd();
				var res = JsonConvert.DeserializeObject<List<KeyValuePair<string, string>>>(body).ToDictionary(x => x.Key, y => y.Value);
			});

		}

		public static void setValue(out string toSet, string val) {
			toSet = val;
		}

		static void Main(string[] args) {

			List<Dictionary<string, string>> table = null; 
			JFetch("http://mysafeinfo.com/api/data?list=englishmonarchs&format=json", table);
			Console.Read();
		}

	}
}