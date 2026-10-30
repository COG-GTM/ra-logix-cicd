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
using static ConsoleFormatter_ClassLibrary.ConsoleFormatter;
using static ConsoleFormatter_ClassLibrary.PathSafety;

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
        public static void RenameFile(string oldFilePath, string newFileName)
        {
            // Get the directory of the old file
            string directory = Path.GetDirectoryName(oldFilePath)!;

            // Combine the directory and new file name to form the new file path. The new name must not leave the directory.
            string newFilePath = EnsurePathWithinFolder(Path.Combine(directory, ValidateFileName(newFileName)), directory, "Renamed file path");

            // Rename the file
            File.Move(oldFilePath, newFilePath);
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