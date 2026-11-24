using System;
using System.IO;
using System.Text;

namespace WordGit
{
    // Cross-implementation check driven by tests/rust-compat.ps1: the original C# history code and
    // the Rust core must read each other's repositories and produce identical snapshot trees.
    internal static class RustCompat
    {
        private static int Main(string[] args)
        {
            Console.OutputEncoding = new UTF8Encoding(false);
            try
            {
                var storage = args[1];
                var document = args[2];
                var repo = new GitHistory(document, storage);
                switch (args[0])
                {
                    case "location":
                        Console.WriteLine(GitHistory.RepositoryFor(document, storage));
                        break;
                    case "commit":
                        Console.WriteLine(repo.CommitFile(document, args[3], "C# 客户端", repo.Head()) ?? "unchanged");
                        break;
                    case "prepare":
                        var prepared = repo.PrepareSnapshotFile(args[3]);
                        Console.WriteLine(prepared.Tree + " " + prepared.Checksum);
                        break;
                    case "extract":
                        using (var output = new FileStream(args[4], FileMode.CreateNew)) repo.ExtractTo(args[3], output);
                        break;
                    case "head":
                        Console.WriteLine(repo.ActiveBranch() + " " + repo.Head());
                        break;
                    case "history":
                        foreach (var revision in repo.History()) Console.WriteLine(revision.Id + " " + revision.Subject);
                        break;
                    default:
                        throw new ArgumentException("unknown command " + args[0]);
                }
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error.GetType().Name + ": " + error.Message);
                return 1;
            }
        }
    }
}
