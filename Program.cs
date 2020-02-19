/* 
 * RunApp URL Protocol
 * by Noah Coad, http://noahcoad.com, http://coadblog.com
 *
 * Created: Oct 12, 2006
 * An example of creating a URL Protocol Handler in Windows
 * 
 * For information, references, resources, etc, see:
 *
 * Register a Custom URL Protocol Handler
 * http://blogs.msdn.com/noahc/archive/2006/10/19/register-a-custom-url-protocol-handler.aspx
 * 
 */

#region Namespace Inclusions
using System;
using System.IO;
using System.Xml;
using System.Reflection;
using System.Diagnostics;
using System.Windows.Forms;
using System.Text.RegularExpressions;
#endregion

namespace RunAppUrlProtocol
{
	class Program
	{
		static void Main(string[] args)
		{
			// Example of a Microsoft Visio protocol link found in SharePoint 2019
			//ms-visio:ofe|u|http://sharepoint2019/sites/engineering/NetworkEngineering/Network%20Diagrams/Infinera_Network_Map_Plotter_12_05_18.vsd

			// The URL handler for this app
			string prefix = "ms-visio";

			// The name of this app for user messages
			string title = "RunApp URL Protocol Handler";

			// Verify the command line arguments
			if (args.Length == 0 || !args[0].StartsWith(prefix))
			{ MessageBox.Show("Syntax:\nrunapp://<key>", title); return; }

			// Obtain the part of the protocol we're interested in
			string key = Regex.Match(args[0], @"(?<=://).+?(?=:|/|\Z)").Value;

			// Convert the Url format to a UNC path
			string url = Regex.Match(args[0], @"(?<=://).+.").Value;
			url = url.Replace("/", "\\").Replace("%20", " ");
			url = "\"\\\\" + url + "\"";

			// Path to the configuration file
			string file = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "RegisteredApps.xml");

			// Verify the config file exists
			if (!File.Exists(file))
			{ MessageBox.Show("Could not find configuration file.\n" + file, title, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

			// Load the config file
			XmlDocument xml = new XmlDocument();
			try { xml.Load(file); }
			catch (XmlException e) 
			{ MessageBox.Show(String.Format("Error loading the XML config file.\n{0}\n{1}", file, e.Message), title, MessageBoxButtons.OK, MessageBoxIcon.Error); return; }

			// Locate the app to run
			XmlNode node = xml.SelectSingleNode(String.Format("/RunApp/App[@key='{0}']", key));

			// If the app is not found, let the user know
			if (node == null)
			{ MessageBox.Show("Key not found: " + key, title); return; }
	
			// Resolve the target app name
			string target = Environment.ExpandEnvironmentVariables(node.SelectSingleNode("@target").Value);

			// Pull the command line args for the target app if they exist
			string procargs = Environment.ExpandEnvironmentVariables(
				node.SelectSingleNode("@args") != null ? 
				node.SelectSingleNode("@args").Value : "");

			// Start the application
			//Process.Start(target, procargs);
			Process.Start(target, url);
		}
	}
}
