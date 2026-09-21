using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SpotLight.Setup
{
    internal static class Program
    {
        private const string ProductName = "SpotLight";
        private const string ResourcePrefix = "SpotLight.Payload.";
        private const string UninstallRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall";

        [STAThread]
        private static int Main(string[] args)
        {
            bool verifyOnly = HasArgument(args, "--verify");
            try
            {
                if (verifyOnly)
                {
                    VerifyPayload();
                    return 0;
                }

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                return Install();
            }
            catch (Exception exception)
            {
                if (!verifyOnly)
                {
                    MessageBox.Show(
                        Localize(
                            "安装失败：\r\n\r\n" + exception.Message,
                            "Installation failed:\r\n\r\n" + exception.Message),
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                return 1;
            }
        }

        private static int Install()
        {
            if (IsExcelRunning())
            {
                MessageBox.Show(
                    Localize(
                        "请先关闭所有 Excel 窗口，然后重新运行安装程序。",
                        "Close all Excel windows, then run the installer again."),
                    ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return 2;
            }

            DialogResult confirmation = MessageBox.Show(
                Localize(
                    "将为当前 Windows 用户安装或更新 SpotLight。\r\n\r\n已有的 SpotLight VSTO 版本会先自动卸载。",
                    "SpotLight will be installed or updated for the current Windows user.\r\n\r\nAny existing SpotLight VSTO installation will be removed first."),
                ProductName,
                MessageBoxButtons.OKCancel,
                MessageBoxIcon.Information);
            if (confirmation != DialogResult.OK)
            {
                return 3;
            }

            string stagingDirectory = Path.Combine(
                Path.GetTempPath(),
                "SpotLightSetup-" + Guid.NewGuid().ToString("N"));
            string installDirectory = GetInstallDirectory();

            try
            {
                ExtractPayload(stagingDirectory);
                VerifyExtractedPayload(stagingDirectory);
                UninstallExistingVersions();
                ReplaceInstallDirectory(stagingDirectory, installDirectory);

                string bootstrapperPath = Path.Combine(installDirectory, "setup.exe");
                int exitCode = RunProcess(bootstrapperPath, string.Empty, installDirectory);
                if (exitCode != 0)
                {
                    throw new InvalidOperationException(
                        Localize(
                            "Microsoft VSTO 安装程序返回错误代码 " + exitCode + "。",
                            "The Microsoft VSTO installer returned error code " + exitCode + "."));
                }

                MessageBox.Show(
                    Localize(
                        "SpotLight 安装完成。现在可以重新打开 Excel。",
                        "SpotLight was installed successfully. You can now reopen Excel."),
                    ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return 0;
            }
            finally
            {
                TryDeleteDirectory(stagingDirectory);
            }
        }

        private static bool HasArgument(IEnumerable<string> args, string expected)
        {
            foreach (string argument in args)
            {
                if (string.Equals(argument, expected, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsExcelRunning()
        {
            Process[] processes = Process.GetProcessesByName("EXCEL");
            try
            {
                return processes.Length > 0;
            }
            finally
            {
                foreach (Process process in processes)
                {
                    process.Dispose();
                }
            }
        }

        private static string GetInstallDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "Programs",
                ProductName,
                "publish");
        }

        private static void VerifyPayload()
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            bool hasBootstrapper = false;
            bool hasDeploymentManifest = false;
            int payloadCount = 0;

            foreach (string resourceName in assembly.GetManifestResourceNames())
            {
                if (!resourceName.StartsWith(ResourcePrefix, StringComparison.Ordinal))
                {
                    continue;
                }

                payloadCount++;
                string relativePath = DecodeResourcePath(resourceName);
                hasBootstrapper |= string.Equals(relativePath, "setup.exe", StringComparison.OrdinalIgnoreCase);
                hasDeploymentManifest |= string.Equals(relativePath, "SpotLight.vsto", StringComparison.OrdinalIgnoreCase);
            }

            if (payloadCount < 3 || !hasBootstrapper || !hasDeploymentManifest)
            {
                throw new InvalidDataException("The embedded VSTO publish payload is incomplete.");
            }
        }

        private static void ExtractPayload(string destinationRoot)
        {
            Assembly assembly = Assembly.GetExecutingAssembly();
            string fullRoot = Path.GetFullPath(destinationRoot) + Path.DirectorySeparatorChar;
            Directory.CreateDirectory(fullRoot);

            foreach (string resourceName in assembly.GetManifestResourceNames())
            {
                if (!resourceName.StartsWith(ResourcePrefix, StringComparison.Ordinal))
                {
                    continue;
                }

                string relativePath = DecodeResourcePath(resourceName);
                string destinationPath = Path.GetFullPath(Path.Combine(fullRoot, relativePath));
                if (!destinationPath.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException("The embedded payload contains an invalid path.");
                }

                string destinationDirectory = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrEmpty(destinationDirectory))
                {
                    Directory.CreateDirectory(destinationDirectory);
                }

                using (Stream input = assembly.GetManifestResourceStream(resourceName))
                {
                    if (input == null)
                    {
                        throw new InvalidDataException("An embedded payload file could not be read.");
                    }

                    using (FileStream output = File.Create(destinationPath))
                    {
                        input.CopyTo(output);
                    }
                }
            }
        }

        private static string DecodeResourcePath(string resourceName)
        {
            string encoded = resourceName.Substring(ResourcePrefix.Length)
                .Replace('-', '+')
                .Replace('_', '/');
            int padding = encoded.Length % 4;
            if (padding != 0)
            {
                encoded = encoded.PadRight(encoded.Length + (4 - padding), '=');
            }

            return Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
        }

        private static void VerifyExtractedPayload(string directory)
        {
            if (!File.Exists(Path.Combine(directory, "setup.exe")) ||
                !File.Exists(Path.Combine(directory, "SpotLight.vsto")) ||
                Directory.GetFiles(directory, "SpotLight.dll.manifest", SearchOption.AllDirectories).Length == 0)
            {
                throw new InvalidDataException(
                    Localize(
                        "安装包中的 VSTO 发布文件不完整。",
                        "The VSTO publish files in this installer are incomplete."));
            }
        }

        private static void UninstallExistingVersions()
        {
            List<string> uninstallCommands = new List<string>();
            using (RegistryKey uninstallRoot = Registry.CurrentUser.OpenSubKey(UninstallRegistryPath))
            {
                if (uninstallRoot == null)
                {
                    return;
                }

                foreach (string subKeyName in uninstallRoot.GetSubKeyNames())
                {
                    using (RegistryKey entry = uninstallRoot.OpenSubKey(subKeyName))
                    {
                        if (entry == null)
                        {
                            continue;
                        }

                        string displayName = entry.GetValue("DisplayName") as string;
                        string uninstallString = entry.GetValue("UninstallString") as string;
                        if (string.Equals(displayName, ProductName, StringComparison.OrdinalIgnoreCase) &&
                            !string.IsNullOrWhiteSpace(uninstallString))
                        {
                            uninstallCommands.Add(uninstallString);
                        }
                    }
                }
            }

            foreach (string command in uninstallCommands)
            {
                string executable;
                string arguments;
                SplitCommandLine(command, out executable, out arguments);
                if (Path.GetFileName(executable).Equals("VSTOInstaller.exe", StringComparison.OrdinalIgnoreCase))
                {
                    arguments = NormalizeVstoUninstallArguments(arguments);
                }

                int exitCode = RunProcess(executable, arguments, Path.GetDirectoryName(executable));
                if (exitCode != 0)
                {
                    throw new InvalidOperationException(
                        Localize(
                            "无法自动卸载已安装的 SpotLight。请先在 Windows“已安装的应用”中卸载 SpotLight。",
                            "The existing SpotLight installation could not be removed automatically. Uninstall SpotLight from Windows Installed apps first."));
                }
            }
        }

        private static void SplitCommandLine(string command, out string executable, out string arguments)
        {
            string expanded = Environment.ExpandEnvironmentVariables(command).Trim();
            if (expanded.StartsWith("\"", StringComparison.Ordinal))
            {
                int closingQuote = expanded.IndexOf('"', 1);
                if (closingQuote < 0)
                {
                    throw new InvalidDataException("The existing uninstall command is invalid.");
                }

                executable = expanded.Substring(1, closingQuote - 1);
                arguments = expanded.Substring(closingQuote + 1).Trim();
                return;
            }

            int executableEnd = expanded.IndexOf(".exe", StringComparison.OrdinalIgnoreCase);
            if (executableEnd < 0)
            {
                throw new InvalidDataException("The existing uninstall command is invalid.");
            }

            executableEnd += 4;
            executable = expanded.Substring(0, executableEnd).Trim();
            arguments = expanded.Substring(executableEnd).Trim();
        }

        private static string NormalizeVstoUninstallArguments(string arguments)
        {
            const string uninstallSwitch = "/Uninstall";
            string normalized = arguments.Trim();
            if (!normalized.StartsWith(uninstallSwitch, StringComparison.OrdinalIgnoreCase))
            {
                return normalized;
            }

            string manifest = normalized.Substring(uninstallSwitch.Length).Trim();
            int silentIndex = manifest.LastIndexOf("/Silent", StringComparison.OrdinalIgnoreCase);
            if (silentIndex >= 0 &&
                string.IsNullOrWhiteSpace(manifest.Substring(silentIndex + "/Silent".Length)))
            {
                manifest = manifest.Substring(0, silentIndex).Trim();
            }

            manifest = manifest.Trim('"');
            return uninstallSwitch + " \"" + manifest + "\" /Silent";
        }

        private static int RunProcess(string executable, string arguments, string workingDirectory)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = arguments,
                WorkingDirectory = string.IsNullOrEmpty(workingDirectory)
                    ? Environment.CurrentDirectory
                    : workingDirectory,
                UseShellExecute = true
            };

            using (Process process = Process.Start(startInfo))
            {
                if (process == null)
                {
                    throw new InvalidOperationException("The installer process could not be started.");
                }

                process.WaitForExit();
                return process.ExitCode;
            }
        }

        private static void ReplaceInstallDirectory(string source, string destination)
        {
            string expectedDestination = Path.GetFullPath(GetInstallDirectory());
            string actualDestination = Path.GetFullPath(destination);
            if (!string.Equals(expectedDestination, actualDestination, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("The installation destination is invalid.");
            }

            if (Directory.Exists(actualDestination))
            {
                Directory.Delete(actualDestination, true);
            }

            CopyDirectory(source, actualDestination);
        }

        private static void CopyDirectory(string source, string destination)
        {
            string sourceRoot = Path.GetFullPath(source).TrimEnd(Path.DirectorySeparatorChar);
            Directory.CreateDirectory(destination);

            foreach (string sourceFile in Directory.GetFiles(sourceRoot, "*", SearchOption.AllDirectories))
            {
                string relativePath = sourceFile.Substring(sourceRoot.Length)
                    .TrimStart(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string destinationFile = Path.Combine(destination, relativePath);
                Directory.CreateDirectory(Path.GetDirectoryName(destinationFile));
                File.Copy(sourceFile, destinationFile, true);
            }
        }

        private static void TryDeleteDirectory(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        private static string Localize(string chinese, string english)
        {
            return string.Equals(
                CultureInfo.CurrentUICulture.TwoLetterISOLanguageName,
                "zh",
                StringComparison.OrdinalIgnoreCase)
                ? chinese
                : english;
        }
    }
}
