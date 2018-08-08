using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.IO;
using Newtonsoft.Json;

namespace JFetch {
	class Class1 {

		//public static void JFetch (HttpWebRequest request) {
		//	List<Dictionary<string, string>> result;
		//	Action wrapperAction = (List<Dictionary<string, string>> res) => {
		//		request.BeginGetResponse(new AsyncCallback((iar) => {
		//			var response = (HttpWebResponse)((HttpWebRequest)iar.AsyncState).EndGetResponse(iar);
		//			var body = new StreamReader(response.GetResponseStream()).ReadToEnd();
		//			List<Dictionary<string, string>> table = JsonConvert.DeserializeObject<List<Dictionary<string, string>>>(body);
		//			result = table;
		//		}), request);
		//	};
		//	wrapperAction.BeginInvoke(new AsyncCallback((iar) => {
		//		var action = (Action)iar.AsyncState;
		//		action.EndInvoke(iar);
		//	}), wrapperAction);
		//}
		//static void Main(string[] args) {
		//}
	}
}
