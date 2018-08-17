using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using JFetch;

namespace OrionStuff {
	public static class Orion {

		private static string uname = "intern4@fpcm.net";
		private static string pword = "Unix15cool";

		private static bool loggedIn = false;
		private static bool loggingIn = false;
		private static HttpClient client = new HttpClient();
		private static string token = "";

		private static List<string> stratGroups = new List<string>();
		public static List<KeyValuePair<string, string>> accountIds = new List<KeyValuePair<string, string>>();

		public static async Task AuthAsync(string username, string password) {
			if (!loggingIn && !loggedIn) {
				loggingIn = true;
				client.DefaultRequestHeaders.Add("Authorization", "Basic " + Base64Encode(username + ":" + password));
				var response = await client.GetAsync("https://api.orionadvisor.com/api/v1/Security/Token").ConfigureAwait(false);
				//response.EnsureSuccessStatusCode();
				var j = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
				try {
					Dictionary<string, string> respDict = JsonConvert.DeserializeObject<Dictionary<string, string>>(j);
					token = respDict["access_token"];
					loggedIn = true;
					client.DefaultRequestHeaders.Clear();
					client.DefaultRequestHeaders.Add("Authorization", "Session " + token);
				}
				catch (Exception ex) {
					Console.WriteLine(ex.ToString());
				}
			}
		}

		public static async Task<object[,]> FP_Focus(string account, string date) {
			if (!loggedIn) {
				await AuthAsync(uname, pword);
			}
			if (stratGroups.Count == 0) {
				await GetStratGroups();
			}
			if (stratGroups.Contains(account.ToLower())) {
				var jstring = @"
				{
					'prompts': [
						{
							'id': 17307,
							'code': '@asof',
							'prompt': 'As Of Date',
							'promptDescription': '',
							'promptType': 'Date',
							'defaultValue': '',
							'isPromptUser': true,
							'sortOrder': null
						},
						{
							'id': 23342,
							'code': '@group',
							'prompt': 'Group',
							'promptDescription': 'Enter FPSUP, CMSUP, OASUP, ACSUP, CCSUP, ACBALA1, ACBALA2, ACBALA3, MISUP, or EQUITY',
							'promptType': 'Text',
							'defaultValue': '',
							'isPromptUser': true,
							'sortOrder': null
						}
					]
				}
				";
				JObject jobj = JObject.Parse(jstring);
				jobj["prompts"][0]["defaultValue"] = date;
				jobj["prompts"][1]["defaultValue"] = account;
				var payload = new StringContent(jobj.ToString(), Encoding.UTF8, "application/json");
				var array = await JFetch.JFetch.JFetchAsync("https://api.orionadvisor.com/api/v1/Reporting/custom/13095/Generate/Table", client, "post", payload);
				return array;
			}
			else {
				//return new object[,] { { "Not implemented" } };
				int idToUse = 0;
				foreach (KeyValuePair<string, string> kvp in accountIds) {
					if (kvp.Key.ToLower() == account.ToLower()) {
						idToUse = Int32.Parse(kvp.Value);
						break;
					}
				}
				return new object[,] { { idToUse } };
			}
		}

		public static async Task GetStratGroups() {
			if (!loggedIn) {
				await AuthAsync(uname, pword);
			}
			//We're gonna query Orion report #13962 for the list of all the consolidated groups.
			string jstring = @"
			{
				'prompts': [
					{
						'id': 18928,
						'code': '@asof',
						'prompt': 'As Of Date',
						'promptDescription': '',
						'promptType': 'Date',
						'defaultValue': '8/16/2018',
						'isPromptUser': true,
						'sortOrder': null
					}
				]
			}
			";
			var payload = new StringContent(jstring, Encoding.UTF8, "application/json");
			object[,] queryResult = await JFetch.JFetch.JFetchAsync("https://api.orionadvisor.com/api/v1/Reporting/custom/13962/generate/table", client, "post", payload);
			for (int i = 0; i < queryResult.GetLength(0); i++) {
				stratGroups.Add(((string) queryResult[i, 0]).ToLower());
			}
		}

		public static async Task GetAccountIds() {
			if (!loggedIn) {
				await AuthAsync(uname, pword);
			}
			string jstring = @"
			{
			'prompts': [
				{
				  'id': 20782,
				  'code': '@asOF',
				  'prompt': 'Enter Date',
				  'promptDescription': '',
				  'promptType': 'Date',
				  'defaultValue': '8/17/2018',
				  'isPromptUser': true,
				  'sortOrder': null
				}
			  ]
			}
			";
			var payload = new StringContent(jstring, Encoding.UTF8, "application/json");
			var result = await client.PostAsync("https://api.orionadvisor.com/api/v1/reporting/custom/14926/generate/table", payload).ConfigureAwait(false);
			var resultstring = await result.Content.ReadAsStringAsync().ConfigureAwait(false);
			JArray jres = JArray.Parse(resultstring);
			foreach (JObject jobj in jres) {
				if (!((string) jobj["shortname"] == "SHORTNAME")) {
					accountIds.Add(new KeyValuePair<string, string>((string) jobj["shortname"], (string) jobj["account ID"]));
				}
			}
		}

		internal static string Base64Encode(string plaintext) {
			var plaintextbytes = System.Text.Encoding.UTF8.GetBytes(plaintext);
			return System.Convert.ToBase64String(plaintextbytes);
		}

	}
}