using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using Microsoft.Win32;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;

class Program
{
    const string fontUrl = "https://github.com/be5invis/Iosevka/releases/download/v33.2.6/SuperTTC-SGr-Iosevka-33.2.6.zip";
    static readonly string tempDir = Path.Combine(Path.GetTempPath(), "iosevka-superttc");
    static readonly string fontsDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
    const string regPath = @"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts";
    const string regValueName = "Iosevka Super TTC (TrueType)";

    static async Task Main(string[] args)
    {
        try
        {
            if (!IsRunAsAdmin())
            {
                Console.WriteLine("This program must be run as Administrator.");
                return;
            }

            if (args.Length == 0)
            {
                ShowUsage();
                return;
            }

            string command = args[0].ToLowerInvariant();

            if (command == "install")
            {
                await InstallFont();
            }
            else if (command == "uninstall")
            {
                UninstallFont();
            }
            else
            {
                ShowUsage();
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
        finally
        {
            Cleanup();
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    static void ShowUsage()
    {
        Console.WriteLine("Usage:");
        Console.WriteLine("  iosevka.exe install   - Install the Iosevka font");
        Console.WriteLine("  iosevka.exe uninstall - Uninstall the Iosevka font");
    }

    static bool IsRunAsAdmin()
    {
        var id = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(id);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    static async Task InstallFont()
    {
        using (var key = Registry.LocalMachine.OpenSubKey(regPath, writable: false))
        {
            var installedFontFile = key?.GetValue(regValueName) as string;
            if (!string.IsNullOrEmpty(installedFontFile))
            {
                Console.WriteLine("Font is already installed.");
                return;
            }
        }

        Console.WriteLine("Downloading font...");
        Directory.CreateDirectory(tempDir);
        string zipPath = Path.Combine(tempDir, "iosevka.zip");
        await DownloadFile(fontUrl, zipPath);

        Console.WriteLine("Extracting font...");
        string ttcPath = ExtractTtc(zipPath);

        if (ttcPath == null)
        {
            Console.WriteLine("Failed to find TTC font in the archive.");
            return;
        }

        Console.WriteLine("Installing font...");
        string destPath = Path.Combine(fontsDir, Path.GetFileName(ttcPath));
        File.Copy(ttcPath, destPath, true);

        using (var key = Registry.LocalMachine.OpenSubKey(regPath, writable: true))
        {
            key.SetValue(regValueName, Path.GetFileName(destPath));
        }

        Console.WriteLine("Font installed successfully!");
    }

    static void UninstallFont()
    {
        Console.WriteLine("Checking if font is installed...");

        using (var key = Registry.LocalMachine.OpenSubKey(regPath, writable: true))
        {
            if (key == null)
            {
                Console.WriteLine("Registry key not found. Font is probably not installed.");
                return;
            }

            var installedFontFile = key.GetValue(regValueName) as string;
            if (string.IsNullOrEmpty(installedFontFile))
            {
                Console.WriteLine("Font is not installed.");
                return;
            }

            Console.WriteLine("Uninstalling font...");

            // Remove font file
            string fontPath = Path.Combine(fontsDir, installedFontFile);
            try
            {
                if (File.Exists(fontPath))
                {
                    File.Delete(fontPath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to delete font file: " + ex.Message);
                return;
            }

            // Remove registry value
            try
            {
                using (var writableKey = Registry.LocalMachine.OpenSubKey(regPath, writable: true))
                {
                    writableKey.DeleteValue(regValueName, throwOnMissingValue: false);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed to remove registry entry: " + ex.Message);
                return;
            }

            Console.WriteLine("Font uninstalled successfully!");
        }
    }

    static async Task DownloadFile(string url, string outputPath)
    {
        using var client = new HttpClient();
        using var stream = await client.GetStreamAsync(url);
        using var fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
        await stream.CopyToAsync(fileStream);
    }

    static string ExtractTtc(string zipPath)
    {
        using var archive = ZipFile.OpenRead(zipPath);
        foreach (var entry in archive.Entries)
        {
            if (entry.FullName.EndsWith(".ttc", StringComparison.OrdinalIgnoreCase))
            {
                string extractPath = Path.Combine(tempDir, entry.Name);
                entry.ExtractToFile(extractPath, overwrite: true);
                return extractPath;
            }
        }
        return null;
    }

    static void Cleanup()
    {
        try
        {
            if (Directory.Exists(tempDir))
                Directory.Delete(tempDir, true);
        }
        catch { }
    }
}
