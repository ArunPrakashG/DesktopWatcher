using System;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.ServiceProcess;
using System.Threading.Tasks;

namespace DesktopWatcher {
	public partial class ServiceCore : ServiceBase {
		private string IMAGE_DIRECTORY => Path.Combine(DesktopPath, "Images");
		private string EDITOR_DIRECTORY => Path.Combine(DesktopPath, "Edited");

		private FileSystemWatcher Watcher;
		private string DesktopPath;
		private static readonly List<string> ImageExtensions = new List<string>() {
			".jpg",
			".png",
			".jpeg",
			".bmp",
			".tiff",
			".gif"
		};

		private static readonly List<string> EditorExtensions = new List<string>() {
			".psd"
		};

		public ServiceCore() {
			InitializeComponent();			
		}

		private void OnCreated(object sender, FileSystemEventArgs e) {

		}

		private async void OnChanged(object sender, FileSystemEventArgs e) {
			if (!File.Exists(e.FullPath)) {
				return;
			}

			await WaitForFile(e.FullPath).ConfigureAwait(false);
			EventLog.WriteEntry($"File change ({e.ChangeType}) detected => " + e.FullPath);
			switch (e.ChangeType) {
				case WatcherChangeTypes.Changed:
				case WatcherChangeTypes.Created:
					if (!File.Exists(e.FullPath)) {
						return;
					}

					FileInfo info = new FileInfo(e.FullPath);

					if (!Path.GetFullPath(info.Directory.FullName).Equals(Path.GetFullPath(DesktopPath))) {
						return;
					}

					string ext = info.Extension.ToLower();
					string fileName = Path.GetFileName(e.FullPath);

					if (ImageExtensions.Contains(ext)) {
						if (!Directory.Exists(IMAGE_DIRECTORY)) {
							Directory.CreateDirectory(IMAGE_DIRECTORY);
						}

						string targetPath = Path.Combine(IMAGE_DIRECTORY, fileName);
						File.Move(e.FullPath, targetPath);
						EventLog.WriteEntry("Image file moved.");
						return;
					}

					if (EditorExtensions.Contains(ext)) {
						if (!Directory.Exists(EDITOR_DIRECTORY)) {
							Directory.CreateDirectory(EDITOR_DIRECTORY);
						}

						string targetPath = Path.Combine(EDITOR_DIRECTORY, fileName);
						File.Move(e.FullPath, targetPath);
						EventLog.WriteEntry("Editor change moved");
						return;
					}

					break;
				case WatcherChangeTypes.Deleted:
				case WatcherChangeTypes.Renamed:
				case WatcherChangeTypes.All:
					break;
			}
		}

		protected override void OnStart(string[] args) {
			DesktopPath = AddQuotesIfRequired(GetDesktopPath());
			EventLog.WriteEntry(DesktopPath);
			Watcher = new FileSystemWatcher(Path.GetDirectoryName(DesktopPath), "*");
			Watcher.Changed += OnChanged;
			Watcher.Created += OnCreated;
			Watcher.IncludeSubdirectories = false;
			Watcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.CreationTime;
			EventLog.WriteEntry("File watcher started for path -> " + DesktopPath);
			Watcher.EnableRaisingEvents = true;
			base.OnStart(args);
		}

		protected override void OnStop() {
			Watcher.EnableRaisingEvents = false;
			Watcher.Dispose();
			base.OnStop();
		}

		private string AddQuotesIfRequired(string path) {
			return !string.IsNullOrWhiteSpace(path) ?
				path.Contains(" ") && (!path.StartsWith("\"") && !path.EndsWith("\"")) ?
					"\"" + path + "\"" : path :
					string.Empty;
		}

		private string GetDesktopPath() => Path.Combine(Directory.GetParent(GetWindowsFolder()).Parent.Name, "Users", GetWindowsUserAccountName(), "Desktop");

		public static string GetWindowsFolder() {
			string windowsFolder = string.Empty;
			ManagementScope ms = new ManagementScope("\\\\.\\root\\cimv2");
			ObjectQuery query = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(ms, query);
			foreach (ManagementObject m in searcher?.Get()) {
				windowsFolder = m["WindowsDirectory"]?.ToString();
			}
			windowsFolder = windowsFolder.Substring(0, windowsFolder.IndexOf(@"\"));
			return windowsFolder;
		}

		private static string GetWindowsUserAccountName() {
			string userName = string.Empty;
			ManagementScope ms = new ManagementScope("\\\\.\\root\\cimv2");
			ObjectQuery query = new ObjectQuery("select * from win32_computersystem");
			ManagementObjectSearcher searcher = new ManagementObjectSearcher(ms, query);

			foreach (ManagementObject mo in searcher?.Get()) {
				userName = mo["username"]?.ToString();
			}
			userName = userName?.Substring(userName.IndexOf(@"\") + 1);

			return userName;
		}

		private static async Task WaitForFile(string filename) {
			while (!IsFileReady(filename)) {
				await Task.Delay(5).ConfigureAwait(false);
			}
		}

		private static bool IsFileReady(string filename) {
			try {
				using (FileStream inputStream = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None))
					return inputStream.Length > 0;
			}
			catch (Exception) {
				return false;
			}
		}
	}
}
