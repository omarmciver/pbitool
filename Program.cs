using System;
using System.IO.Compression;
using Microsoft.AnalysisServices.AdomdClient;

namespace PBIXAnalyzer
{

    public class Program
    {
    public static void Main(string[] args)
    {
        if (args.Length != 1)
        {
            Console.WriteLine("Usage: PBIXAnalyzer <path_to_pbix_file>");
            return;
        }

        var analyzer = new PBIXAnalyzer(args[0]);
        analyzer.PrintContents();
        analyzer.AnalyzeModel();
    }
    }
