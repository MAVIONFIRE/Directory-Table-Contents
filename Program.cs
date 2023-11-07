using System;
using System.IO;
using Xceed.Words.NET;

class Program
{
    static void Main(string[] args)
    {
        string directoryPath = "";
        string outputDir = "";

        GenerateTableOfContents(directoryPath, outputDir);
    }

    static void GenerateTableOfContents(string directoryPath, string outputDir)
    {
        using (DocX document = DocX.Create(Path.Combine(outputDir, "TableOfContents.docx")))
        {
            document.InsertParagraph("Table of Contents").Bold().FontSize(16);

            GenerateDirectoryTOC(directoryPath, document, 0);

            document.Save();
        }

        Console.WriteLine($"Table of Contents saved to {Path.Combine(outputDir, "TableOfContents.docx")}");
    }

    static void GenerateDirectoryTOC(string directoryPath, DocX document, int level)
    {
        var levelBuffer = level + 1;
        foreach (var item in Directory.GetFileSystemEntries(directoryPath))
        {
            if (File.Exists(item))
            {
                string fileName = Path.GetFileName(item);

                // Add the dashes before the text and the level in parentheses after the name
                string tabSpaces = new string(' ', level * 8);
                document.InsertParagraph(tabSpaces + new string('-', level) + " - " + fileName + " (" + levelBuffer + ")").SpacingBefore(5);
            }
            else if (Directory.Exists(item))
            {
                string folderName = Path.GetFileName(item);

                // Insert folder names, add an asterisk (*) after them, add dashes before the text, and include the level in parentheses after the name
                string tabSpaces = new string(' ', level * 8);
                document.InsertParagraph(tabSpaces + new string('-', level) + " - " + folderName + "* (" + levelBuffer + ")").SpacingBefore(5);

                // Recursively generate the table of contents for subdirectories with increased horizontal spacing
                GenerateDirectoryTOC(item, document, level + 1);
            }
        }
    }
}
