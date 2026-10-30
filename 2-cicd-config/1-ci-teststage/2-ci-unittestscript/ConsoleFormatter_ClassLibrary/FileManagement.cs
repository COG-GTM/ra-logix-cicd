// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:    FileManagement.cs
// FileType:    Visual C# Source File
// Author:      Rockwell Automation
// Created:     2024
// Description: This class provides the methods required to redirect console outputs to both a text file and the console.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using System.Text;
using System.Text.RegularExpressions;
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;

namespace ConsoleFormatter_ClassLibrary
{
    /// <summary>
    /// Static class to manage console output redirection.
    /// </summary>
    public static class FileManagement
    {
        private static TextWriter? _oldOut;
        private static StreamWriter? _fileWriter;

        /// <summary>
        /// Starts redirecting the console output to both the console and the file.
        /// </summary>
        /// <param name="path">The path to the output file.</param>
        public static void StartLogging(string path)
        {
            FileStream ostrm = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Write);
            _fileWriter = new StreamWriter(ostrm) { AutoFlush = true };
            _oldOut = Console.Out;

            DualWriter dualWriter = new DualWriter(Console.Out, _fileWriter);
            Console.SetOut(dualWriter);
        }

        /// <summary>
        /// Stops redirecting the console output and restores the original output.
        /// </summary>
        public static void StopLogging()
        {
            Console.SetOut(_oldOut!);
            _fileWriter!.Close();
        }

        /// <summary>
        /// Method to retain only a specified number of files of a certain type (extension) in a specified folder.
        /// </summary>
        /// <param name="folderPath">The file path to the folder.</param>
        /// <param name="numberOfFilesToRetain">The number of files to be retained.</param>
        /// <param name="fileExtension">Valid input examples: .txt, .xlsx, .ACD, etc.</param>
        public static void RetainMostRecentFiles(string folderPath, int numberOfFilesToRetain, string fileExtension)
        {
            // Get all .txt files in the directory
            var txtFiles = new DirectoryInfo(folderPath).GetFiles("*" + fileExtension);

            // Order the files by the last write time, descending
            var sortedFiles = txtFiles.OrderByDescending(f => f.CreationTime).ToList();

            // Retain only the specified number of recent files
            var filesToRetain = sortedFiles.Take(numberOfFilesToRetain);

            foreach (var file in filesToRetain)
            {
                ConsoleMessage($"Retained '{file.Name}'");
            }

            // Determine files to delete
            var filesToDelete = sortedFiles.Skip(numberOfFilesToRetain);

            // Delete the older files
            foreach (var file in filesToDelete)
            {
                try
                {
                    file.Delete();
                    ConsoleMessage($"Deleted '{file.Name}'");
                }
                catch (Exception ex)
                {
                    ConsoleMessage($"Error deleting file: {file.Name}. Exception: {ex.Message}", "ERROR");
                }
            }
        }

        /// <summary>
        /// Rename a specified file.
        /// </summary>
        /// <param name="oldFilePath">A string containing the full file path of the file to be renamed.</param>
        /// <param name="newFileName">A string containing the new name of the file.</param>
        /// <exception cref="ArgumentException">Thrown if the new file name contains a directory component.</exception>
        public static void RenameFile(string oldFilePath, string newFileName)
        {
            if (string.IsNullOrWhiteSpace(newFileName) || newFileName.IndexOfAny(['\\', '/']) >= 0 || Path.GetFileName(newFileName) != newFileName
                || newFileName is "." or "..")
                throw new ArgumentException($"New file name '{newFileName}' must not contain a directory component.", nameof(newFileName));

            // Get the directory of the old file
            string directory = Path.GetDirectoryName(oldFilePath)!;

            // Combine the directory, new file name, and extension to form the new file path
            string newFilePath = Path.Combine(directory, newFileName);

            // Rename the file
            File.Move(oldFilePath, newFilePath);
        }

        /// <summary>
        /// Resolve a path read from an input excel workbook against a root folder and ensure the result stays inside that root folder.<br/>
        /// Both '\' and '/' are treated as directory separators.
        /// </summary>
        /// <param name="rootFolderPath">The folder the resolved path must stay within (for example, the local GitHub folder path).</param>
        /// <param name="relativePath">The path relative to the root folder.</param>
        /// <returns>The full path of the resolved file.</returns>
        /// <exception cref="InvalidDataException">Thrown if the path is empty, rooted, resolves outside the root folder, or goes through a link.</exception>
        public static string ResolvePathWithinFolder(string rootFolderPath, string relativePath)
        {
            string normalizedRelativePath = relativePath.Replace('\\', Path.DirectorySeparatorChar).Replace('/', Path.DirectorySeparatorChar);
            if (string.IsNullOrWhiteSpace(normalizedRelativePath) || Path.IsPathRooted(normalizedRelativePath))
                throw new InvalidDataException($"Path '{relativePath}' must be a non-empty path relative to '{rootFolderPath}'.");

            string fullRootPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(rootFolderPath)) + Path.DirectorySeparatorChar;
            string fullPath = Path.GetFullPath(Path.Combine(fullRootPath, normalizedRelativePath));
            StringComparison comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            if (fullPath.Length <= fullRootPath.Length || !fullPath.StartsWith(fullRootPath, comparison))
                throw new InvalidDataException($"Path '{relativePath}' resolves to '{fullPath}', which is outside '{fullRootPath}'.");

            // Reject symbolic links and junctions below the root folder so the resolved path cannot be redirected outside it.
            string currentPath = fullRootPath;
            foreach (string pathPart in fullPath[fullRootPath.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries))
            {
                currentPath = Path.Combine(currentPath, pathPart);
                if ((File.Exists(currentPath) || Directory.Exists(currentPath) || new FileInfo(currentPath).LinkTarget != null)
                    && File.GetAttributes(currentPath).HasFlag(FileAttributes.ReparsePoint))
                    throw new InvalidDataException($"Path '{relativePath}' goes through the link '{currentPath}', which is not allowed.");
            }

            return fullPath;
        }

        // Studio 5000 Logix Designer component names: a letter or underscore, then letters, digits, or underscores (40 characters max).
        private static readonly Regex logixComponentNameRegex = new(@"\A[A-Za-z_][A-Za-z0-9_]{0,39}\z", RegexOptions.CultureInvariant);

        /// <summary>
        /// Validate that a name read from an L5X file is a valid Studio 5000 Logix Designer component name, so it is safe to use in file names.
        /// </summary>
        /// <param name="name">The name to validate (for example, the RSLogix5000Content TargetName attribute).</param>
        /// <param name="description">Where the name came from, used in the error message.</param>
        /// <returns>The validated name.</returns>
        /// <exception cref="InvalidDataException">Thrown if the name is not a valid Logix component name.</exception>
        public static string ValidateLogixComponentName(string? name, string description)
        {
            if (name == null || !logixComponentNameRegex.IsMatch(name))
                throw new InvalidDataException($"{description} '{name}' is not a valid Logix component name (letters, digits, and underscores only, " +
                    "starting with a letter or underscore, 40 characters max).");

            return name;
        }

        /// <summary>
        /// Insert a custom, third line to a text file.
        /// </summary>
        /// <param name="newThirdLine">The contents of the third line to insert into the text file.</param>
        /// <param name="textFilePath">The file path to the text file.</param>
        public static void AddThirdLineToTextFile(string newThirdLine, string textFilePath)
        {
            // Read all lines from the file
            string[] lines = File.ReadAllLines(textFilePath);

            // Create a new list to store the updated lines
            var updatedLines = new List<string>();

            // Add the first line to the updated list
            if (lines.Length > 0)
            {
                updatedLines.Add(lines[0]);
            }
            if (lines.Length > 1)
            {
                updatedLines.Add(lines[1]);
            }

            // Add the new line at the third line position
            updatedLines.Add(newThirdLine);

            // Add the rest of the original lines
            for (int i = 2; i < lines.Length; i++)
            {
                updatedLines.Add(lines[i]);
            }

            // Write the updated lines back to the file
            File.WriteAllLines(textFilePath, updatedLines);
        }
    }

    /// <summary>
    /// DualWriter class that writes to both the console and a file.
    /// </summary>
    public class DualWriter : TextWriter
    {
        private readonly TextWriter _consoleWriter;
        private readonly TextWriter _fileWriter;

        /// <summary>
        /// Initializes a new instance of the DualWriter class.
        /// </summary>
        /// <param name="consoleWriter">The TextWriter for the console.</param>
        /// <param name="fileWriter">The TextWriter for the file.</param>
        public DualWriter(TextWriter consoleWriter, TextWriter fileWriter)
        {
            _consoleWriter = consoleWriter;
            _fileWriter = fileWriter;
        }

        /// <summary>
        /// Gets the Encoding of the writer.
        /// </summary>
        public override Encoding Encoding => _consoleWriter.Encoding;

        /// <summary>
        /// Writes a character to both the console and the file.
        /// </summary>
        /// <param name="value">The character to write.</param>
        public override void Write(char value)
        {
            _consoleWriter.Write(value);
            _fileWriter.Write(value);
        }

        /// <summary>
        /// Flushes both writers.
        /// </summary>
        public override void Flush()
        {
            _consoleWriter.Flush();
            _fileWriter.Flush();
        }

        /// <summary>
        /// Disposes of the writers.
        /// </summary>
        /// <param name="disposing">Indicates whether to release managed resources.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _consoleWriter.Dispose();
                _fileWriter.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}