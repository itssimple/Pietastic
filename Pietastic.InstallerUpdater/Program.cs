using IWshRuntimeLibrary;
using NuGet;
using System;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Pietastic.InstallerUpdater
{
	class Program
	{
		[STAThread]
		static int Main(params string[] args)
		{
			var cBase = Assembly.GetExecutingAssembly().CodeBase.Replace("file:///", "").Replace("/", "\\");
			var fileName = cBase.Substring(cBase.LastIndexOf('\\') + 1);
			fileName = fileName.Substring(0, fileName.LastIndexOf('.'));
			Console.WriteLine("Package: " + fileName);
			bool uninstall = false;
			bool upgrade = false;
			string packageName = fileName;
			if (args.Length > 0)
			{
				switch (args[0])
				{
					case "uninstall":
						uninstall = true;
						break;
					case "upgrade":
						upgrade = false;
						break;
					default:
						upgrade = true;
						break;
				}
			}
			if (string.IsNullOrWhiteSpace(packageName)) return -1;

			IPackageRepository pr = PackageRepositoryFactory.Default.CreateRepository("http://192.168.0.17/Pietastic.WebUpdateService/nuget");

			var latestVersion = pr.FindPackagesById(packageName).LastOrDefault(p => p.IsLatestVersion && p.Listed);
			if (latestVersion != null)
			{
				// Check if package is installed

				var packagePath = Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
					packageName
				);

				if (uninstall)
				{
					Uninstall(packagePath);
					Process.GetCurrentProcess().Kill();
					return 0;
				}

				Console.WriteLine("Creating App directory");
				Console.WriteLine(packagePath);
				Directory.CreateDirectory(packagePath);

				var launcherFile = Path.Combine(packagePath, packageName + ".exe");
				if (!System.IO.File.Exists(launcherFile))
				{
					System.IO.File.Copy(cBase, launcherFile);
				}

				var programPackages = Path.Combine(packagePath, "packages");

				Console.WriteLine("Creating Packages directory");
				Console.WriteLine(programPackages);
				Directory.CreateDirectory(programPackages);

				var latestPackage = Path.Combine(
					programPackages,
					string.Format("{0}-{1}.nupkg", latestVersion.Id, latestVersion.Version.ToFullString())
				);

				var latestInstalled = Path.Combine(
					packagePath,
					string.Format("app-{0}", latestVersion.Version.ToFullString())
				);

				Console.WriteLine(latestInstalled);

				if (!System.IO.File.Exists(latestPackage))
				{
					Console.WriteLine("Latest version not downloaded, removing old versions and downloading latest");
					DirectoryInfo di = new DirectoryInfo(programPackages);

					var allPackages = di.EnumerateFiles("*.nupkg").ToList();

					if (allPackages.Count > 0)
					{
						// Enumerate all packages, remove old packages, keep 3 latest
						var skipTheThreeLatest = allPackages.OrderByDescending(f => f.CreationTimeUtc).Skip(3);
						foreach (var pack in skipTheThreeLatest)
						{
							var appFolder = string.Format("app-{0}", pack.Name.Substring(pack.Name.LastIndexOf('-') + 1).Replace(".nupkg", ""));
							if (Directory.Exists(appFolder))
							{
								Directory.Delete(Path.Combine(packagePath, appFolder), true);
							}
						}
					}

					using (var fs = latestVersion.GetStream())
					{
						System.IO.File.WriteAllBytes(latestPackage, fs.ReadAllBytes());
					}
				}

				string symLinkPath = string.Empty;

				Console.WriteLine("Installing latest version");
				Directory.CreateDirectory(latestInstalled);

				var fz = new ZipPackage(latestPackage);
				foreach (var f in fz.GetFiles())
				{
					var lFile = Path.Combine(latestInstalled, f.Path);
					if (f.Path == latestVersion.Id + ".exe")
					{
						Console.WriteLine("Modifying/Creating symlink");

						symLinkPath = Path.Combine(packagePath, f.Path);
						if (SymbolicLink.Exists(symLinkPath))
							System.IO.File.Delete(symLinkPath);
						SymbolicLink.CreateFileLink(symLinkPath, lFile);
						Console.WriteLine(symLinkPath);
					}

					if (System.IO.File.Exists(lFile)) continue;

					using (var fs = f.GetStream())
					{
						System.IO.File.WriteAllBytes(lFile, fs.ReadAllBytes());
						if (f.Path != latestVersion.Id + ".exe") continue;

					}
				}
				
				var desktopShortcut = Path.Combine(
					Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
					(latestVersion.Title ?? packageName) + ".lnk"
				);

				if (!string.IsNullOrWhiteSpace(symLinkPath))
				{
					if (System.IO.File.Exists(desktopShortcut)) System.IO.File.Delete(desktopShortcut);
					CreateShortcut(desktopShortcut, symLinkPath);
					Console.WriteLine(desktopShortcut);
				}

				if (!upgrade)
				{
					ProcessStartInfo launch = new ProcessStartInfo(desktopShortcut);
					var p = Process.Start(launch);
					p.WaitForExit();
					Main("upgrade " + packageName);
				}

				return 0;
			}
			Console.WriteLine("Package: " + packageName + " was not found, please rename file");
			return 1;
		}

		internal static void CreateShortcut(string shortcutTarget, string targetFileLocation, string description = "")
		{
			WshShell shell = new WshShell();
			IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(shortcutTarget);

			shortcut.Description = description;   // The description of the shortcut
			shortcut.TargetPath = targetFileLocation;                 // The path of the file that will launch when the shortcut is run
			shortcut.Save();                                    // Save the shortcut
		}

		internal static void RegisterAsInstalled(string applicationName)
		{

		}

		internal static void Uninstall(string applicationPath)
		{
			/*ProcessStartInfo Info = new ProcessStartInfo();
			Info.Arguments = "/C choice /C Y /N /D Y /T 3 & RMDIR /Q /S " +
						   applicationPath;
			Info.WindowStyle = ProcessWindowStyle.Hidden;
			Info.CreateNoWindow = true;
			Info.FileName = "cmd.exe";
			Process.Start(Info);*/
		}
	}
}
