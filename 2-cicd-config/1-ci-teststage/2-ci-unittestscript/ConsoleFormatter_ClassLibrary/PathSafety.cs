// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:    PathSafety.cs
// FileType:    Visual C# Source File
// Author:      Rockwell Automation
// Created:     2026
// Description: This class validates file paths and object names read from input excel workbooks and L5X files before they are used to read or
//              write files, so that a workbook cell or L5X attribute cannot point the test program outside its repository or output folders.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using System.Text.RegularExpressions;

namespace ConsoleFormatter_ClassLibrary
{
    /// <summary>
    /// Exception thrown when a value from an input excel workbook or L5X file would make the test program read or write outside its allowed folders.
    /// </summary>
    public class UnsafeInputException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the UnsafeInputException class.
        /// </summary>
        /// <param name="message">A description of the rejected value.</param>
        public UnsafeInputException(string message) : base(message) { }
    }

    /// <summary>
    /// Static class containing path and name validation methods.
    /// </summary>
    public static class PathSafety
    {
        // Studio 5000 Logix Designer tag-name rule: starts with a letter or underscore, letters/digits/underscores only, 40 characters maximum.
        private static readonly Regex logixNamePattern = new Regex(@"^[A-Za-z_][A-Za-z0-9_]{0,39}\z", RegexOptions.CultureInvariant);

        private static StringComparison PathComparison => OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

        /// <summary>
        /// Resolve a relative path against a root folder and require the result to stay inside that root folder.
        /// </summary>
        /// <param name="rootFolderPath">The folder the resolved path must stay inside (for example, the local GitHub folder).</param>
        /// <param name="relativePath">The relative path to resolve (for example, the contents of an input excel workbook cell).</param>
        /// <param name="description">A description of where the value came from, used in the error message.</param>
        /// <returns>The fully qualified path.</returns>
        /// <exception cref="UnsafeInputException">The path is empty, absolute, or resolves outside the root folder.</exception>
        public static string ResolvePathWithinFolder(string rootFolderPath, string relativePath, string description)
        {
            if (string.IsNullOrWhiteSpace(relativePath))
                throw new UnsafeInputException($"{description} is empty.");

            if (Path.IsPathRooted(relativePath))
                throw new UnsafeInputException($"{description} '{relativePath}' must be a path relative to '{rootFolderPath}'.");

            return EnsurePathWithinFolder(Path.Combine(rootFolderPath, relativePath), rootFolderPath, description);
        }

        /// <summary>
        /// Require a path to resolve to a location inside a specified folder.
        /// </summary>
        /// <param name="path">The path to check.</param>
        /// <param name="folderPath">The folder the path must stay inside.</param>
        /// <param name="description">A description of the path, used in the error message.</param>
        /// <returns>The fully qualified path.</returns>
        /// <exception cref="UnsafeInputException">The path resolves outside the folder.</exception>
        public static string EnsurePathWithinFolder(string path, string folderPath, string description)
        {
            string fullFolderPath = Path.TrimEndingDirectorySeparator(Path.GetFullPath(folderPath)) + Path.DirectorySeparatorChar;
            string fullPath = Path.GetFullPath(path);

            if (!fullPath.StartsWith(fullFolderPath, PathComparison))
                throw new UnsafeInputException($"{description} '{path}' resolves to '{fullPath}', which is outside '{fullFolderPath}'.");

            return fullPath;
        }

        /// <summary>
        /// Require a Studio 5000 Logix Designer object name (AOI name, controller name, etc.) to follow the Logix tag-name rule.
        /// </summary>
        /// <param name="name">The name to check.</param>
        /// <param name="description">A description of where the name came from, used in the error message.</param>
        /// <returns>The validated name.</returns>
        /// <exception cref="UnsafeInputException">The name is missing or does not follow the Logix tag-name rule.</exception>
        public static string ValidateLogixName(string? name, string description)
        {
            if (name == null || !logixNamePattern.IsMatch(name))
                throw new UnsafeInputException($"{description} '{name}' is not a valid Logix name (letters, digits and underscores only, starting " +
                    "with a letter or underscore, 40 characters maximum).");

            return name;
        }

        /// <summary>
        /// Require a value to be a plain file name with no directory components.
        /// </summary>
        /// <param name="fileName">The file name to check.</param>
        /// <returns>The validated file name.</returns>
        /// <exception cref="UnsafeInputException">The value contains directory separators or invalid file name characters.</exception>
        public static string ValidateFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) || fileName == "." || fileName == ".." || Path.GetFileName(fileName) != fileName ||
                fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                throw new UnsafeInputException($"'{fileName}' is not a valid file name.");

            return fileName;
        }
    }
}
