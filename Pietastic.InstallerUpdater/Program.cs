using IWshRuntimeLibrary;
using NuGet;
using System;
using System.Configuration;
using System.IO;
using System.Linq;

namespace Pietastic.InstallerUpdater
{
	class Program
	{
		[STAThread]
		static int Main(string[] args)
		{
			string packageName = ConfigurationManager.AppSettings["pietastic_packageName"] ?? string.Empty;
			if (args.Length > 0)
			{
				packageName = string.Join(" ", args);
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

				Console.WriteLine("Creating App directory");
				Console.WriteLine(packagePath);
				Directory.CreateDirectory(packagePath);

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
				if (!Directory.Exists(latestInstalled))
				{
					Console.WriteLine("Installing latest version");
					Directory.CreateDirectory(latestInstalled);

					var fz = new ZipPackage(latestPackage);
					foreach (var f in fz.GetFiles())
					{
						var lFile = Path.Combine(latestInstalled, f.Path);
						if (System.IO.File.Exists(lFile)) continue;

						using (var fs = f.GetStream())
						{
							System.IO.File.WriteAllBytes(lFile, fs.ReadAllBytes());
							if (f.Path != latestVersion.Id + ".exe") continue;
							Console.WriteLine("Modifying/Creating symlink");

							symLinkPath = Path.Combine(packagePath, f.Path);
							if (SymbolicLink.Exists(symLinkPath))
								System.IO.File.Delete(symLinkPath);
							SymbolicLink.CreateFileLink(symLinkPath, lFile);
							Console.WriteLine(symLinkPath);
						}
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
				
				return 0;
			}

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
	}
}
