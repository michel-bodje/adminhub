using System;
using System.IO;
using System.Threading;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.WriteLine("Usage: cleanTempDoc.exe <path>");
            return;
        }

        string filePath = args[0];
        Console.WriteLine("Watching temp file for release: " + filePath);

        const int timeoutMinutes = 30;
        const int checkIntervalMs = 30000; // 30 seconds
        int waited = 0;

        while (waited < timeoutMinutes * 60 * 1000)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File already deleted.");
                return;
            }

            if (IsFileUnlocked(filePath))
            {
                try
                {
                    File.Delete(filePath);
                    Console.WriteLine("Temp file deleted successfully.");
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Failed to delete file: " + ex.Message);
                }
            }

            Thread.Sleep(checkIntervalMs);
            waited += checkIntervalMs;
        }

        Console.WriteLine("Timeout reached. File still in use or deletion failed.");
    }

    static bool IsFileUnlocked(string path)
    {
        try
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                return true;
            }
        }
        catch (IOException)
        {
            return false; // File is locked
        }
    }
}
