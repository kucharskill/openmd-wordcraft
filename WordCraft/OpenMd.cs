using System;
using System.IO;
using System.Text;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Threading;
using System.Net;
using System.Net.Sockets;
using System.Windows.Forms;
using System.Drawing;
using System.IO.Compression;
using System.Web.Script.Serialization;
using System.Text.RegularExpressions;
using System.Reflection;

[assembly: AssemblyTitle("OpenMD for WordCraft")]
[assembly: AssemblyProduct("OpenMD for WordCraft")]
[assembly: AssemblyVersion("0.1.0.0")]
[assembly: AssemblyFileVersion("0.1.0.0")]
[assembly: AssemblyCopyright("Copyright (c) 2026 kucharskill")]

namespace OpenMdWordCraft
{
    public sealed class Settings { public string WordCraft; public string Pandoc; }
    public sealed class ImageRecord { public string Hash; public string Target; public string FullPath; }
    public sealed class Session
    {
        public string Source, Directory, Docx, Pandoc, WordCraft, SourceHash, DocxHash;
        public string FrontMatter, NewLine, AssetsDirectory;
        public bool Bom;
        public bool UsePipeTables;
        public List<string> Comments = new List<string>();
        public List<ImageRecord> Images = new List<ImageRecord>();
        public int Port, Pid;
    }

    internal static class Program
    {
        static readonly JavaScriptSerializer Json = new JavaScriptSerializer { MaxJsonLength = 134217728, RecursionLimit = 512 };
        static readonly string Home = Path.Combine(Path.GetTempPath(), "openmd-wordcraft");
        static bool Headless;
        static string ResultPath;
        static string ActiveDirectory;
        const string LegacyCommentNotice = "<!-- openmd: Komentarze HTML przeniesiono na początek dokumentu z ich pierwotnych miejsc w treści z powodu konwersji MD -> DOCX -> MD. -->";
        const string CommentNotice = "<!-- openmd: HTML comments were moved from their original locations to the top of this document because of MD -> DOCX -> MD conversion. -->";
        static readonly UTF8Encoding Utf8 = new UTF8Encoding(false, true);

        [STAThread]
        static int Main(string[] args)
        {
            Headless = args.Contains("--headless");
            Application.EnableVisualStyles();
            try
            {
                ResultPath = Option(args, "--result");
                if (args.Contains("--version"))
                {
                    WriteResult(new { version = "0.1.0" }); Report("OpenMD for WordCraft 0.1.0"); return 0;
                }
                if (args.Contains("--help"))
                {
                    Report("openmd.exe document.md [--wordcraft EXE] [--pandoc EXE]\n--version: show version; --prepare: prepare only; --sync SESSION.JSON: export saved DOCX; --result FILE: write result JSON; --headless: no dialogs.");
                    return 0;
                }
                string sync = Option(args, "--sync");
                if (sync != null)
                {
                    Session saved = Read<Session>(sync);
                    ActiveDirectory = saved.Directory;
                    Export(saved);
                    WriteResult(new { session = Path.Combine(saved.Directory, "session.json"), source = saved.Source });
                    return 0;
                }
                string source = Positional(args);
                if (source == null)
                {
                    if (Headless) throw new Exception("Missing Markdown path.");
                    source = Pick("Select a Markdown document", "Markdown (*.md;*.markdown)|*.md;*.markdown");
                    if (source == null) return 0;
                }
                source = Path.GetFullPath(source);
                if (!File.Exists(source)) throw new Exception("Document not found: " + source);
                if (!new[] { ".md", ".markdown" }.Contains(Path.GetExtension(source).ToLowerInvariant()))
                    throw new Exception("Select an .md or .markdown file.");
                Directory.CreateDirectory(Home);
                string configPath = Path.Combine(Home, "config.json");
                Settings settings = File.Exists(configPath) ? Read<Settings>(configPath) : new Settings();
                string wc = Option(args, "--wordcraft") ?? RunningWordCraft() ?? settings.WordCraft;
                string pandoc = Option(args, "--pandoc") ?? OnPath("pandoc.exe") ?? settings.Pandoc;
                wc = ResolveExecutable(wc, "Locate wordcraft.exe");
                if (wc == null) return 0;
                pandoc = ResolveExecutable(pandoc, "Locate pandoc.exe");
                if (pandoc == null) return 0;
                settings.WordCraft = wc; settings.Pandoc = pandoc;
                WriteJson(configPath, settings);
                bool created;
                using (Mutex mutex = new Mutex(true, "Local\\OpenMdWordCraft-" + Hash(Utf8.GetBytes(source.ToUpperInvariant())).Substring(0, 24), out created))
                {
                    if (!created) throw new Exception("This document is already open in OpenMD. Close its existing session before opening it again.");
                    try
                    {
                        Session session = Prepare(source, settings);
                        WriteResult(new { session = Path.Combine(session.Directory, "session.json"), docx = session.Docx });
                        if (!args.Contains("--prepare")) Monitor(session);
                    }
                    finally { mutex.ReleaseMutex(); }
                }
                return 0;
            }
            catch (Exception e)
            {
                string error = e.Message + (ActiveDirectory == null ? "" : "\nSession files: " + ActiveDirectory);
                if (ActiveDirectory != null) File.WriteAllText(Path.Combine(ActiveDirectory, "error.txt"), e.ToString(), Utf8);
                if (ResultPath != null) WriteJson(ResultPath, new { error = error });
                Report(error);
                return 1;
            }
        }

        static string Option(string[] args, string key)
        {
            int index = Array.IndexOf(args, key);
            if (index < 0) return null;
            if (index + 1 >= args.Length || args[index + 1].StartsWith("--")) throw new Exception("Missing value: " + key);
            return args[index + 1];
        }
        static string Positional(string[] args)
        {
            string found = null;
            for (int i = 0; i < args.Length; i++)
            {
                if (new[] { "--result", "--wordcraft", "--pandoc", "--sync" }.Contains(args[i])) { i++; continue; }
                if (new[] { "--headless", "--prepare", "--help" }.Contains(args[i])) continue;
                if (args[i].StartsWith("--")) throw new Exception("Unknown option: " + args[i]);
                if (found != null) throw new Exception("Open one Markdown file per invocation.");
                found = args[i];
            }
            return found;
        }
        static void Report(string message)
        {
            if (Headless) Console.Error.WriteLine(message);
            else MessageBox.Show(message, "OpenMD / WordCraft", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        static string Pick(string title, string filter)
        {
            using (OpenFileDialog dialog = new OpenFileDialog { Title = title, Filter = filter, CheckFileExists = true })
                return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        }
        static string ResolveExecutable(string path, string title)
        {
            if (!String.IsNullOrWhiteSpace(path) && File.Exists(path)) return Path.GetFullPath(path);
            if (Headless) throw new Exception(title);
            return Pick(title, "Executable (*.exe)|*.exe");
        }
        static string RunningWordCraft()
        {
            foreach (Process p in Process.GetProcessesByName("wordcraft"))
                using (p) { try { return p.MainModule.FileName; } catch { } }
            return null;
        }
        static string OnPath(string name)
        {
            foreach (string directory in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
            {
                try { string path = Path.Combine(directory.Trim().Trim('"'), name); if (File.Exists(path)) return Path.GetFullPath(path); }
                catch (ArgumentException) { }
            }
            return null;
        }
        static string Hash(byte[] bytes) { using (SHA256 h = SHA256.Create()) return BitConverter.ToString(h.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant(); }
        static T Read<T>(string path) { return Json.Deserialize<T>(File.ReadAllText(path, Utf8)); }
        static void WriteJson(string path, object value) { File.WriteAllText(path, Json.Serialize(value), Utf8); }
        static void WriteResult(object value) { if (ResultPath != null) WriteJson(ResultPath, value); }
        static Dictionary<string, object> Obj(object value) { return value as Dictionary<string, object>; }
        static IList List(object value) { return value as IList; }
        static string Quote(string value)
        {
            // CommandLineToArgvW escaping, including a trailing backslash before a quote.
            StringBuilder result = new StringBuilder("\""); int slashes = 0;
            foreach (char c in value)
            {
                if (c == '\\') { slashes++; continue; }
                result.Append('\\', c == '"' ? slashes * 2 + 1 : slashes);
                result.Append(c); slashes = 0;
            }
            result.Append('\\', slashes * 2); result.Append('"'); return result.ToString();
        }
        static void Pandoc(Session s, params string[] args)
        {
            ProcessStartInfo info = new ProcessStartInfo(s.Pandoc, String.Join(" ", args.Select(Quote))) {
                UseShellExecute = false, CreateNoWindow = true, RedirectStandardError = true,
                WorkingDirectory = Path.GetDirectoryName(s.Source)
            };
            using (Process p = Process.Start(info))
            {
                string error = p.StandardError.ReadToEnd();
                p.WaitForExit();
                if (p.ExitCode != 0) throw new Exception("Pandoc: " + error);
            }
        }
        static void Walk(object value, Action<Dictionary<string, object>> action)
        {
            Dictionary<string, object> obj = Obj(value);
            if (obj != null) { action(obj); foreach (object child in obj.Values.ToArray()) Walk(child, action); }
            else { IList list = List(value); if (list != null) foreach (object child in list) Walk(child, action); }
        }
        static bool Is(object value, string type) { Dictionary<string, object> d = Obj(value); return d != null && d.ContainsKey("t") && (string)d["t"] == type; }
        static int TableCount(object ast)
        {
            int count = 0; Walk(ast, delegate(Dictionary<string, object> node) { if (Is(node, "Table")) count++; }); return count;
        }
        static string LocalImage(string target, string baseDirectory)
        {
            Uri uri;
            if (Uri.TryCreate(target, UriKind.Absolute, out uri)) return uri.IsFile ? uri.LocalPath : null;
            return Path.GetFullPath(Path.Combine(baseDirectory, Uri.UnescapeDataString(target.Replace('/', Path.DirectorySeparatorChar))));
        }

        static Session Prepare(string source, Settings settings)
        {
            byte[] bytes = File.ReadAllBytes(source);
            bool bom = bytes.Take(3).SequenceEqual(new byte[] { 239, 187, 191 });
            string text = Utf8.GetString(bytes, bom ? 3 : 0, bytes.Length - (bom ? 3 : 0));
            string documentName = Path.GetFileNameWithoutExtension(source) + ".docx";
            string pathHash = Hash(Utf8.GetBytes(Path.GetFullPath(source).ToUpperInvariant()));
            string directory = Path.Combine(Home, "sessions", pathHash);
            // Reuse one location per MD, unless an existing DOCX cannot be replaced.
            string reusableDocx = Path.Combine(directory, documentName);
            if (File.Exists(reusableDocx))
            {
                try
                {
                    using (FileStream probe = new FileStream(reusableDocx, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
                }
                catch (IOException) { directory = Path.Combine(Home, "sessions", pathHash + "-" + Guid.NewGuid().ToString("N")); }
                catch (UnauthorizedAccessException) { directory = Path.Combine(Home, "sessions", pathHash + "-" + Guid.NewGuid().ToString("N")); }
            }
            Directory.CreateDirectory(directory); ActiveDirectory = directory;
            Session s = new Session { Source = source, Directory = directory, Pandoc = settings.Pandoc,
                WordCraft = settings.WordCraft, SourceHash = Hash(bytes), Bom = bom,
                NewLine = text.Contains("\r\n") ? "\r\n" : "\n", FrontMatter = "",
                AssetsDirectory = Path.GetFileNameWithoutExtension(source) + "_img" };
            s.Docx = Path.Combine(directory, documentName);
            File.WriteAllBytes(Path.Combine(directory, "original.md"), bytes);
            Match front = Regex.Match(text, @"\A---[ \t]*\r?\n.*?^(?:---|\.\.\.)[ \t]*(?:\r?\n|$)", RegexOptions.Multiline | RegexOptions.Singleline);
            if (front.Success) { s.FrontMatter = front.Value; text = text.Substring(front.Length); }
            string clean = Path.Combine(directory, "body.md"); File.WriteAllText(clean, text, Utf8);
            string astPath = Path.Combine(directory, "input.json");
            Pandoc(s, clean, "-f", "markdown+mark", "-t", "json", "-o", astPath);
            Dictionary<string, object> ast = Obj(Json.DeserializeObject(File.ReadAllText(astPath, Utf8)));
            // Compare actual parsed tables, so examples in code are not mistaken for pipe tables.
            if (TableCount(ast) > 0)
            {
                string withoutPipes = Path.Combine(directory, "without-pipe-tables.json");
                Pandoc(s, clean, "-f", "markdown+mark-pipe_tables", "-t", "json", "-o", withoutPipes);
                s.UsePipeTables = TableCount(ast) > TableCount(Json.DeserializeObject(File.ReadAllText(withoutPipes, Utf8)));
            }
            ast["meta"] = new Dictionary<string, object>();
            bool chosenDirectory = false;
            Walk(ast, delegate(Dictionary<string, object> node) {
                if (Is(node, "RawBlock") || Is(node, "RawInline"))
                {
                    IList content = List(node["c"]);
                    if ((string)content[0] == "html") content[1] = Regex.Replace((string)content[1], @"<!--[\s\S]*?-->", delegate(Match m) { if (m.Value != CommentNotice && m.Value != LegacyCommentNotice) s.Comments.Add(m.Value); return ""; });
                }
                if (!Is(node, "Image")) return;
                IList target = List(List(node["c"])[2]); string link = (string)target[0];
                string local = LocalImage(link, Path.GetDirectoryName(source));
                if (local == null) return;
                if (!File.Exists(local)) throw new Exception("Local image not found: " + local);
                s.Images.Add(new ImageRecord { Hash = Hash(File.ReadAllBytes(local)), Target = link, FullPath = local });
                if (!chosenDirectory && !Path.IsPathRooted(link) && !Uri.IsWellFormedUriString(link, UriKind.Absolute))
                {
                    string folder = Path.GetDirectoryName(Uri.UnescapeDataString(link.Replace('/', Path.DirectorySeparatorChar)));
                    if (!String.IsNullOrEmpty(folder)) { s.AssetsDirectory = folder; chosenDirectory = true; }
                }
                target[0] = local;
            });
            WriteJson(astPath, ast);
            Pandoc(s, astPath, "-f", "json", "-t", "docx", "-o", s.Docx);
            s.DocxHash = Hash(File.ReadAllBytes(s.Docx)); SaveSession(s);
            return s;
        }
        static void SaveSession(Session s) { WriteJson(Path.Combine(s.Directory, "session.json"), s); }

        static byte[] Snapshot(string path)
        {
            using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (MemoryStream copy = new MemoryStream())
            {
                stream.CopyTo(copy); byte[] bytes = copy.ToArray();
                using (ZipArchive zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read))
                    if (zip.GetEntry("word/document.xml") == null) throw new IOException("Incomplete DOCX file.");
                return bytes;
            }
        }
        static bool Export(Session s)
        {
            byte[] snapshot = Snapshot(s.Docx); string digest = Hash(snapshot);
            if (digest == s.DocxHash) return false;
            string pass = Path.Combine(s.Directory, "exports", Guid.NewGuid().ToString("N")); Directory.CreateDirectory(pass);
            string copy = Path.Combine(pass, "saved.docx"); File.WriteAllBytes(copy, snapshot);
            string astPath = Path.Combine(pass, "output.json"), media = Path.Combine(pass, "media");
            Pandoc(s, copy, "--track-changes=all", "--extract-media=" + media, "-t", "json", "-o", astPath);
            object ast = Json.DeserializeObject(File.ReadAllText(astPath, Utf8));
            List<KeyValuePair<string, byte[]>> pending = new List<KeyValuePair<string, byte[]>>();
            Walk(ast, delegate(Dictionary<string, object> node) {
                if (!Is(node, "Image")) return;
                IList target = List(List(node["c"])[2]); string local = LocalImage((string)target[0], Path.GetDirectoryName(s.Source));
                if (local == null || !File.Exists(local)) throw new Exception("Cannot extract image: " + target[0]);
                byte[] image = File.ReadAllBytes(local); string hash = Hash(image);
                ImageRecord original = s.Images.FirstOrDefault(i => i.Hash == hash);
                // Preserve external edits to image files too: use a new name for the DOCX copy.
                if (original != null && File.Exists(original.FullPath) && Hash(File.ReadAllBytes(original.FullPath)) != hash) original = null;
                if (original != null)
                {
                    target[0] = original.Target;
                    if (!File.Exists(original.FullPath)) pending.Add(new KeyValuePair<string, byte[]>(original.FullPath, image));
                }
                else
                {
                    string extension = Path.GetExtension(local).ToLowerInvariant();
                    if (!Regex.IsMatch(extension, @"^\.[a-z0-9]{1,8}$")) extension = ".bin";
                    string name = "image_" + hash.Substring(0, 24) + extension;
                    string full = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(s.Source), s.AssetsDirectory, name));
                    pending.Add(new KeyValuePair<string, byte[]>(full, image));
                    target[0] = (s.AssetsDirectory.TrimEnd('/', '\\') + "/" + name).Replace('\\', '/');
                }
            });
            WriteJson(astPath, ast); string body = Path.Combine(pass, "body.md");
            string format = s.UsePipeTables ? "markdown+mark-simple_tables-multiline_tables-grid_tables+pipe_tables" : "markdown+mark";
            Pandoc(s, astPath, "-f", "json", "-t", format, "--wrap=none", "--markdown-headings=atx", "-o", body);
            string text = s.FrontMatter;
            if (text.Length > 0 && !text.EndsWith("\n")) text += s.NewLine;
            if (s.Comments.Count > 0) text += CommentNotice + s.NewLine + String.Join(s.NewLine, s.Comments) + s.NewLine + s.NewLine;
            text += File.ReadAllText(body, Utf8).Replace("\r\n", "\n").Replace("\n", s.NewLine);
            byte[] output = new UTF8Encoding(s.Bom).GetPreamble().Concat(Utf8.GetBytes(text)).ToArray();
            string candidate = Path.Combine(pass, "candidate.md"); File.WriteAllBytes(candidate, output);
            // Ask once per saved DOCX; keep the candidate when the user declines.
            if (!File.Exists(s.Source)) throw new Exception("The Markdown file was deleted. Recovered document: " + candidate);
            string expectedSourceHash = Hash(File.ReadAllBytes(s.Source));
            if (expectedSourceHash != s.SourceHash)
            {
                if (Headless) throw new Exception("The Markdown file changed outside OpenMD. Headless mode does not overwrite conflicts. Recovered document: " + candidate);
                DialogResult decision = MessageBox.Show(
                    "The Markdown file was changed outside OpenMD:\n" + s.Source +
                    "\n\nOverwrite it with the version saved in WordCraft?\n" +
                    "Yes: replace Markdown and keep its current version in the session folder.\n" +
                    "No: keep the current Markdown file. The WordCraft version will remain here:\n" + candidate,
                    "OpenMD: save conflict", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
                if (decision != DialogResult.Yes)
                {
                    s.DocxHash = digest; SaveSession(s);
                    File.AppendAllText(Path.Combine(s.Directory, "session.log"), DateTime.Now.ToString("s") + " conflict: user kept external MD; candidate " + candidate + "\r\n", Utf8);
                    return false;
                }
            }
            if (Hash(output) == expectedSourceHash)
            {
                s.SourceHash = expectedSourceHash; s.DocxHash = digest; SaveSession(s); return false;
            }
            foreach (KeyValuePair<string, byte[]> image in pending)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(image.Key));
                if (File.Exists(image.Key))
                {
                    if (Hash(File.ReadAllBytes(image.Key)) != Hash(image.Value)) throw new Exception("Image file conflict: " + image.Key);
                }
                else using (FileStream file = new FileStream(image.Key, FileMode.CreateNew, FileAccess.Write)) file.Write(image.Value, 0, image.Value.Length);
            }
            // File.Replace is atomic on the source volume; retain the previous MD for recovery.
            string sibling = s.Source + ".openmd-" + Guid.NewGuid().ToString("N") + ".tmp";
            File.WriteAllBytes(sibling, output);
            try
            {
                using (FileStream guard = new FileStream(s.Source, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete))
                {
                    using (MemoryStream bytes = new MemoryStream())
                    {
                        guard.CopyTo(bytes);
                        byte[] previous = bytes.ToArray();
                        if (Hash(previous) != expectedSourceHash) throw new IOException("Markdown changed during saving. Session files were retained. Please try again.");
                        File.WriteAllBytes(Path.Combine(pass, "previous.md"), previous);
                    }
                    File.Replace(sibling, s.Source, null);
                }
            }
            finally { if (File.Exists(sibling)) File.Delete(sibling); }
            s.SourceHash = Hash(output); s.DocxHash = digest; SaveSession(s);
            File.AppendAllText(Path.Combine(s.Directory, "session.log"), DateTime.Now.ToString("s") + " saved MD\r\n", Utf8);
            AutoCommit(s, pending.Select(image => image.Key));
            return true;
        }

        sealed class GitResult { public int Code; public string Output; }
        static GitResult Git(string exe, string directory, string index, params string[] args)
        {
            ProcessStartInfo info = new ProcessStartInfo(exe, String.Join(" ", args.Select(Quote))) {
                WorkingDirectory = directory, UseShellExecute = false, CreateNoWindow = true,
                RedirectStandardOutput = true, RedirectStandardError = true
            };
            info.EnvironmentVariables.Remove("GIT_INDEX_FILE");
            if (index != null) info.EnvironmentVariables["GIT_INDEX_FILE"] = index;
            info.EnvironmentVariables["GIT_TERMINAL_PROMPT"] = "0";
            info.EnvironmentVariables["GIT_EDITOR"] = "true";
            info.EnvironmentVariables["GCM_INTERACTIVE"] = "Never";
            using (Process p = Process.Start(info))
            {
                var output = p.StandardOutput.ReadToEndAsync();
                var error = p.StandardError.ReadToEndAsync();
                if (!p.WaitForExit(20000)) { p.Kill(); throw new IOException("Git operation timed out."); }
                string stdout = output.Result; string stderr = error.Result;
                return new GitResult { Code = p.ExitCode, Output = stdout + (p.ExitCode == 0 ? "" : stderr) };
            }
        }
        static void AutoCommit(Session s, IEnumerable<string> images)
        {
            // Optional and silent. A private index excludes unrelated staged work.
            try
            {
                string git = OnPath("git.exe"); if (git == null) return;
                string directory = Path.GetDirectoryName(s.Source);
                GitResult rootResult = Git(git, directory, null, "rev-parse", "--show-toplevel");
                if (rootResult.Code != 0) return;
                string root = Path.GetFullPath(rootResult.Output.Trim());
                string prefix = root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                if (!s.Source.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) return;
                string document = s.Source.Substring(prefix.Length).Replace('\\', '/');
                if (Git(git, root, null, "ls-files", "--error-unmatch", "--", document).Code != 0) return;
                if (Git(git, root, null, "ls-files", "-u", "-z").Output.Length != 0) return;
                List<string> paths = new List<string> { document };
                foreach (string image in images.Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    string full = Path.GetFullPath(image);
                    if (full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) paths.Add(full.Substring(prefix.Length).Replace('\\', '/'));
                }
                string index = Path.Combine(s.Directory, "git-index-" + Guid.NewGuid().ToString("N"));
                if (Git(git, root, index, "read-tree", "HEAD").Code != 0) return;
                string[] add = new[] { "add", "--" }.Concat(paths).ToArray();
                if (Git(git, root, index, add).Code != 0) return;
                if (Git(git, root, index, "diff", "--cached", "--quiet").Code == 0) return;
                GitResult commit = Git(git, root, index, "-c", "commit.gpgsign=false", "commit", "--quiet", "-m",
                    "Apply user changes made in an external editor (WordCraft)");
                if (commit.Code == 0)
                {
                    // Refresh only the committed paths in the real index. Other staged entries stay intact.
                    Git(git, root, null, new[] { "reset", "--quiet", "HEAD", "--" }.Concat(paths).ToArray());
                }
                File.AppendAllText(Path.Combine(s.Directory, "git.log"), DateTime.Now.ToString("s") + " exit=" + commit.Code + " " + commit.Output + "\r\n", Utf8);
            }
            catch (Exception e)
            {
                try { File.AppendAllText(Path.Combine(s.Directory, "git.log"), DateTime.Now.ToString("s") + " " + e.Message + "\r\n", Utf8); }
                catch { /* Optional Git integration must never interrupt a document save. */ }
            }
        }

        static object Control(Session s, string method, object parameters)
        {
            string profile = Environment.GetEnvironmentVariable("WORDCRAFT_PROFILE_DIR");
            if (String.IsNullOrEmpty(profile)) profile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordCraft");
            string key = File.ReadAllText(Path.Combine(profile, "control-key.local-" + s.Port)).Trim();
            using (TcpClient client = new TcpClient())
            {
                IAsyncResult connect = client.BeginConnect("127.0.0.1", s.Port, null, null);
                if (!connect.AsyncWaitHandle.WaitOne(1500)) throw new IOException("Control connection timed out.");
                client.EndConnect(connect);
                NetworkStream stream = client.GetStream(); stream.ReadTimeout = 2000; stream.WriteTimeout = 2000;
                StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(false)) { AutoFlush = true };
                StreamReader reader = new StreamReader(stream);
                writer.WriteLine(Json.Serialize(new { id = 1, key = key, method = method, @params = parameters }));
                Dictionary<string, object> reply = Obj(Json.DeserializeObject(reader.ReadLine()));
                if (!(bool)reply["ok"]) throw new IOException(Convert.ToString(reply["error"]));
                return reply["result"];
            }
        }
        static void Monitor(Session s)
        {
            TcpListener reservation = new TcpListener(IPAddress.Loopback, 0); reservation.Start();
            s.Port = ((IPEndPoint)reservation.LocalEndpoint).Port; reservation.Stop();
            ProcessStartInfo info = new ProcessStartInfo(s.WordCraft, "--control " + s.Port + " " + Quote(s.Docx)) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(s.WordCraft) };
            using (Process process = Process.Start(info))
            using (NotifyIcon tray = new NotifyIcon { Icon = SystemIcons.Application, Text = "OpenMD: " + Path.GetFileName(s.Source).Substring(0, Math.Min(45, Path.GetFileName(s.Source).Length)), Visible = !Headless })
            {
                s.Pid = process.Id; SaveSession(s);
                ContextMenu menu = new ContextMenu(); menu.MenuItems.Add("Open session folder", delegate { Process.Start("explorer.exe", Quote(s.Directory)); }); tray.ContextMenu = menu;
                bool seen = false; int missing = 0, ticks = 0; string pending = null, lastError = null;
                while (true)
                {
                    Application.DoEvents();
                    bool exited = process.HasExited;
                    try
                    {
                        string digest = Hash(Snapshot(s.Docx));
                        if (digest == s.DocxHash) lastError = null;
                        if (digest != s.DocxHash && (exited || digest == pending))
                        {
                            bool written = Export(s); lastError = null;
                            if (written) tray.ShowBalloonTip(1200, "OpenMD", "Saved " + Path.GetFileName(s.Source), ToolTipIcon.Info);
                        }
                        pending = digest;
                    }
                    catch (Exception e)
                    {
                        if (exited) throw;
                        if (lastError != e.Message)
                        {
                            lastError = e.Message; File.AppendAllText(Path.Combine(s.Directory, "session.log"), e + "\r\n", Utf8);
                            tray.ShowBalloonTip(5000, "OpenMD: save pending", e.Message.Substring(0, Math.Min(220, e.Message.Length)), ToolTipIcon.Warning);
                        }
                    }
                    if (exited) break;
                    if (ticks++ % 8 == 0)
                    {
                        try
                        {
                            Dictionary<string, object> state = Obj(Control(s, "ui.inspect", new { }));
                            IList docs = List(state["documents"]);
                            bool open = docs.Cast<object>().Any(d => String.Equals(Convert.ToString(Obj(d)["path"]), s.Docx, StringComparison.OrdinalIgnoreCase));
                            if (open) { seen = true; missing = 0; }
                            else if (seen && ++missing >= 2) { Export(s); break; }
                        }
                        catch (IOException) { /* Older apps: PID + saved file monitoring remains available. */ }
                        catch (SocketException) { }
                    }
                    Thread.Sleep(500);
                }
                if (lastError != null) throw new Exception(lastError);
                tray.Visible = false;
            }
        }
    }
}
