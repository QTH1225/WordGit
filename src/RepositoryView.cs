// SPDX-License-Identifier: GPL-3.0-only
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordGit
{
    internal sealed class RepositoryView
    {
        internal RemoteState Remote;
        internal string[] Branches;
        internal List<Revision> History;
    }

    internal sealed partial class GitHistory
    {
        // Read all branch tips and the relevant configuration in two Git processes.
        internal RepositoryView ReadView()
        {
            if (!File.Exists(Path.Combine(root, "HEAD")))
                return new RepositoryView { Remote = new RemoteState { LocalBranch = "main", RemoteBranch = "main", RemoteReference = RemotePrefix + "main" },
                    Branches = new string[0], History = new List<Revision>() };

            var tips = new Dictionary<string, string>(StringComparer.Ordinal);
            string local = null;
            var references = Text(Run(null, null, "for-each-ref", "--sort=refname",
                "--format=%(HEAD)%00%(refname)%00%(objectname)", Heads, RemotePrefix));
            foreach (var line in references.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var fields = line.Split('\0');
                if (fields.Length != 3) throw new InvalidOperationException("无法读取仓库分支。");
                tips[fields[1]] = Id(fields[2]);
                if (fields[0].Trim() == "*" && fields[1].StartsWith(Heads, StringComparison.Ordinal))
                    local = fields[1].Substring(Heads.Length);
            }
            if (local == null) local = ActiveBranch(); // An unborn branch has no ref yet.
            var settings = new Dictionary<string, string>(StringComparer.Ordinal);
            var config = Invoke(null, null, "config", "--local", "--null", "--get-regexp",
                "^(remote\\.origin\\.url|branch\\..*\\.merge)$");
            if (config.Code != 1)
                foreach (var entry in Utf8.GetString(Checked(config).Bytes).Split(new[] { '\0' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var separator = entry.IndexOf('\n');
                    if (separator >= 0) settings[entry.Substring(0, separator)] = entry.Substring(separator + 1);
                }
            string merge, url, tip;
            settings.TryGetValue("branch." + local + ".merge", out merge);
            settings.TryGetValue("remote.origin.url", out url);
            var remoteBranch = merge != null && merge.StartsWith(Heads, StringComparison.Ordinal) ? merge.Substring(Heads.Length) : local;
            var state = new RemoteState { Url = url, LocalBranch = local, RemoteBranch = remoteBranch,
                RemoteReference = RemotePrefix + remoteBranch };
            if (tips.TryGetValue(Heads + local, out tip)) state.LocalTip = tip;
            if (state.Linked && tips.TryGetValue(state.RemoteReference, out tip)) state.RemoteTip = tip;
            CompleteSyncState(state);
            return new RepositoryView { Remote = state,
                Branches = tips.Keys.Where(reference => reference.StartsWith(Heads, StringComparison.Ordinal))
                    .Select(reference => reference.Substring(Heads.Length)).OrderBy(name => name, StringComparer.Ordinal)
                    .Concat(state.Linked ? tips.Keys.Where(reference => reference.StartsWith(RemotePrefix, StringComparison.Ordinal) &&
                        reference != RemotePrefix + "HEAD").OrderBy(name => name, StringComparer.Ordinal) : Enumerable.Empty<string>()).ToArray(),
                History = History() };
        }
    }
}
