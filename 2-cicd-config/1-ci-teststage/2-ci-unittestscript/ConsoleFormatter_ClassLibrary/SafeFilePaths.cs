// ------------------------------------------------------------------------------------------------------------------------------------------------------------
//
// FileName:    SafeFilePaths.cs
// FileType:    Visual C# Source File
// Author:      Rockwell Automation
// Created:     2026
// Description: This class provides guards for building generated file paths from names read out of L5X/ACD files, which may come from untrusted authors.
//
// ------------------------------------------------------------------------------------------------------------------------------------------------------------

using System.Text.RegularExpressions;

namespace ConsoleFormatter_ClassLibrary
{
    /// <summary>
    /// Static class to validate Logix names and keep generated file paths inside their intended folder.
    /// </summary>
    public static class SafeFilePaths
    {
        /// <summary>
        /// Logix component names: letters, digits, and underscores; must start with a letter or underscore; at most 40 characters.
        /// </summary>
        private static readonly Regex logixNamePattern = new(@"\A[A-Za-z_][A-Za-z0-9_]{0,39}\z", RegexOptions.CultureInvariant);

        /// <summary>
        /// Human-readable description of the accepted Logix name format, used in error messages.
        /// </summary>
        public const string LogixNameRules = "letters, digits, and underscores only; must start with a letter or underscore; at most 40 characters";

        /// <summary>
        /// Check whether a name read from an L5X/ACD file is a valid Logix component name (and therefore safe to use in a file name).
        /// </summary>
        /// <param name="name">The name to check.</param>
        /// <returns>True if the name is a valid Logix component name.</returns>
        public static bool IsValidLogixName(string? name)
        {
            return name is not null && logixNamePattern.IsMatch(name);
        }

        /// <summary>
        /// Return the name if it is a valid Logix component name, otherwise throw.
        /// </summary>
        /// <param name="name">The name to check.</param>
        /// <param name="description">Where the name came from (used in the exception message).</param>
        /// <returns>The validated name.</returns>
        /// <exception cref="InvalidDataException">The name is missing or is not a valid Logix component name.</exception>
        public static string RequireValidLogixName(string? name, string description)
        {
            if (!IsValidLogixName(name))
                throw new InvalidDataException($"{description} is not a valid Logix name ({LogixNameRules}).");

            return name!;
        }

        /// <summary>
        /// Combine a folder path and a file name, guaranteeing that the resulting full path is a file directly inside that folder.
        /// </summary>
        /// <param name="folderPath">The folder in which the file must be located.</param>
        /// <param name="fileName">The file name (no directory components).</param>
        /// <returns>The full path of the file inside the folder.</returns>
        /// <exception cref="ArgumentException">The file name contains directory components or resolves outside the folder.</exception>
        public static string CombineWithinFolder(string folderPath, string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName) || fileName != Path.GetFileName(fileName) || fileName == "." || fileName == "..")
                throw new ArgumentException($"File name '{fileName}' must not contain directory components.", nameof(fileName));

            string fullFolderPath = Path.GetFullPath(folderPath);
            if (!Path.EndsInDirectorySeparator(fullFolderPath))
                fullFolderPath += Path.DirectorySeparatorChar;

            string fullFilePath = Path.GetFullPath(Path.Combine(fullFolderPath, fileName));
            string? parentFolderPath = Path.GetDirectoryName(fullFilePath);
            StringComparison comparison = OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

            if (parentFolderPath is null || !string.Equals(Path.TrimEndingDirectorySeparator(parentFolderPath),
                Path.TrimEndingDirectorySeparator(fullFolderPath), comparison))
                throw new ArgumentException($"File name '{fileName}' resolves outside of folder '{fullFolderPath}'.", nameof(fileName));

            return fullFilePath;
        }
    }
}
