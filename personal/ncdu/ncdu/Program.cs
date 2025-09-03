// C# program to find the estimate size of the folder
using System;
using System.IO;
using System.Linq;

class GFG
{

    // Driver code
    static public void Main( )
    {
        for (int j = 0; j < 2; j++) // repeat  times
        {
            //waiting effect
            Console.Write("Analyzing");
            for (int i = 0; i < 10; i++) // repeat 10 times
            {
                Console.Write(".");
                Thread.Sleep(100); // wait 0.1 second
                if ((i + 1) % 3 == 0) // after 3 dots, reset line
                {
                    Console.SetCursorPosition(0, Console.CursorTop);
                    Console.Write("    Analyzing   "); // overwrite old text
                    Console.SetCursorPosition(0, Console.CursorTop);
                }
            }
        }
        Console.WriteLine("\nDone!");
        /////////////


        Console.Write("Which folder u wanna analyze:");
        string ChoiceFolder = Console.ReadLine();

        // Get the directory information using directoryInfo() method
        DirectoryInfo folder = new DirectoryInfo(ChoiceFolder);

        // Calling a folderSize() method
        long totalFolderSize = GetFolderSize(folder);


        Console.WriteLine("\nTotal folder size in bytes: " + totalFolderSize);

        // Calling human-readable conversion
        Console.WriteLine($"Your total in human readable version: {HumanReadableConversion(totalFolderSize)}");

    }

 
    // LINQ version of folder size calculation
    static long GetFolderSize(DirectoryInfo folder) =>
        folder.EnumerateFiles("*", SearchOption.AllDirectories)
              .Sum(file => file.Length);


    // Function to calculate the size of the folder
    static string HumanReadableConversion(long bytes)
    {
        string[] sizes = { "B", "KB", "MB", "GB", "TB" };
        double len = bytes;
        int order = 0;


        while (len >= 1024 && order < sizes.Length) {
            order++;
            len/=1024;
        }

        return $"{len:0.###} {sizes[order]}";
    }
}
