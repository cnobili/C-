/*
 * Program: Decompress.cs
 * Author:  Craig Nobili
 * 
 * Description: Decompress a file.
 * 
 * Note: To compile add references to System.IO.Compression.dll
 *       and System.IO.Compression.FileSystem.dll.
 */

using System;
using System.IO;
using System.IO.Compression;

class Decompress
{
    public static void MoveFiles(String srcDir, String dstDir)
    {
        String dstFile;
        String[] fileList = Directory.GetFiles(srcDir, "*.zip");
        
        foreach(String file in fileList)
        {
          dstFile = dstDir + Path.DirectorySeparatorChar + Path.GetFileName(file);
          Console.WriteLine("Moving file {0} to {1}", file, dstFile);
          File.Move(file, dstFile);
        }
    
    } // MoveFiles()
    
    public static void UnZip(String fileToDecompress, String extractToDirectory)
    {
        
        ZipFile.ExtractToDirectory(fileToDecompress, extractToDirectory);

    } // Decompress()
    
    public static void Main(String[] args)
    {
        String dir = null;
        String moveToDir = null;
        
        if ( !(args.Length == 1 || args.Length == 2) )
        {
            Console.WriteLine("Usage: Decompress directoryWithZipFiles [moveToDirectory]");
            Console.WriteLine();
            Console.WriteLine("    directoryWithZipFiles - Directory containing the zip file or files to decompress.");
            Console.WriteLine("    moveToDirectory       - Optional directory to move the zip file(s) to after decompressing them.");
            return;
        }
        dir = args[0];
        
        if (args.Length == 2)
        {
            moveToDir = args[1];
        }
            
        Console.WriteLine("Process Zip files in directory = {0}", dir);
                
        String[] fileList = Directory.GetFiles(dir, "*.zip");
        
        foreach(String file in fileList)
        {
          Console.WriteLine("Unzipping file = {0}", file);
          UnZip(file, dir);
        }
        
        if (moveToDir != null)
        {
            MoveFiles(dir, moveToDir);
        }
    
    } // Main()

} // class Decompress

