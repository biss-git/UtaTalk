using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Yomiage.SDK;
using Yomiage.SDK.Config;
using Yomiage.SDK.Settings;

namespace UtaLibrary
{
    public class UtaLibrary : IVoiceLibrary
    {
        private string ConfigDirectory;
        private string DllDirectory;
        private string VoiceDirectory;
        private List<Oto> OtoList = new List<Oto>();
        private int fs = 44100;

        private Version version
        {
            get
            {
                //自分自身のAssemblyを取得
                System.Reflection.Assembly asm =
                    System.Reflection.Assembly.GetExecutingAssembly();
                return asm.GetName().Version;
            }
        }
        public int MajorVersion => version.Major;
        public int MinorVersion => version.Minor;


        public LibraryConfig Config { get; private set; }

        public LibrarySettings Settings { get; set; }

        public bool IsActivated => true;

        public bool IsEnable => true;

        public string StateText { get; private set; } = string.Empty;

        public async Task<bool> Activate(string key)
        {
            StateText = "アクティベートされました。";
            return true;
        }

        public async Task<bool> DeActivate()
        {
            StateText = "ディアクティベートされました。";
            return false;
        }


        public void Dispose()
        {
        }

        public object GetValue(string command, string key)
        {
            StateText = "呼び出されました。command = " + command + ", key = " + key;

            var oto = SearchOto(key);
            if (oto == null)
            {
                key = FixKey(key);
                oto = SearchOto(key);
            }
            if (oto == null && key.Contains(" R"))
            {
                key = FixKey2(key);
                oto = SearchOto(key);
            }
            if (oto == null) { return null; }

            var breathVolume = 0.0;
            if (Settings?.Doubles?.ContainsKey("breath") == true)
            {
                breathVolume = Settings.Doubles["breath"].Value;
            }

            switch (command.ToLower())
            {
                case "wave":
                    return oto.GetWave(breathVolume);
                case "timing":
                    return oto.Preceding_Sample;
                case "start":
                    return oto.Overlap_Sample;
                case "fixed":
                    return oto.Fixed_Sample;
                case "fs":
                    return fs;
            }
            return null;
        }
        private Oto SearchOto(string key)
        {
            var oto = OtoList.FirstOrDefault(o => o.Alias == key); // まず完全一致
            if (oto != null) { return oto; }
            oto = OtoList.FirstOrDefault(o => o.Alias.Contains(key + "_")); // 次に_付き(スを探してスィがヒットしないように)
            if (oto != null) { return oto; }
            oto = OtoList.FirstOrDefault(o => o.Alias.Contains(key + " ")); // 次に半角スペース付き(スを探してスィがヒットしないように)
            if (oto != null) { return oto; }
            oto = OtoList.Where(o => o.Alias.Contains(key)).OrderBy(o => o.Alias.Length).FirstOrDefault(); // 最後にゆるく検索する、Ailas の短いものを優先、これでも見つからなければ諦める。
            if (oto != null) { return oto; }
            return null;
        }

        private string FixKey(string key)
        {
            switch (key)
            {
                case "a ー": return "a ア";
                case "i ー": return "i イ";
                case "u ー": return "u ウ";
                case "e ー": return "e エ";
                case "o ー": return "o オ";
                case "n ー": return "u ン"; // n ン は無いので u ン にしてみる。
                case "a ッ": return "a R";
                case "i ッ": return "i R";
                case "u ッ": return "u R";
                case "e ッ": return "e R";
                case "o ッ": return "o R";
                case "n ッ": return "n R";
                case "a ァ": return "a ア";
                case "i ァ": return "i ア";
                case "u ァ": return "u ア";
                case "e ァ": return "e ア";
                case "o ァ": return "o ア";
                case "n ァ": return "n ア";
                case "a ィ": return "a イ";
                case "i ィ": return "i イ";
                case "u ィ": return "u イ";
                case "e ィ": return "e イ";
                case "o ィ": return "o イ";
                case "n ィ": return "n イ";
                case "a ゥ": return "a ウ";
                case "i ゥ": return "i ウ";
                case "u ゥ": return "u ウ";
                case "e ゥ": return "e ウ";
                case "o ゥ": return "o ウ";
                case "n ゥ": return "n ウ";
                case "a ェ": return "a エ";
                case "i ェ": return "i エ";
                case "u ェ": return "u エ";
                case "e ェ": return "e エ";
                case "o ェ": return "o エ";
                case "n ェ": return "n エ";
                case "a ォ": return "a オ";
                case "i ォ": return "i オ";
                case "u ォ": return "u オ";
                case "e ォ": return "e オ";
                case "o ォ": return "o オ";
                case "n ォ": return "n オ";
            }
            return key;
        }
        private string FixKey2(string key)
        {
            switch (key)
            {
                case "a R": return "ア ・";
                case "i R": return "イ ・";
                case "u R": return "ウ ・";
                case "e R": return "エ ・";
                case "o R": return "オ ・";
                case "n R": return "ン ・";
            }
            return key;
        }

        public bool TryGetValue<T>(string command, string key, out T value)
        {
            var result = GetValue(command, key);
            if (command.ToLower() == "wave" &&
                typeof(T) == typeof(double[]))
            {
                value = (T)result;
                if (value == null) { return false; }
                return true;
            }
            if (typeof(T) == typeof(int) &&
                result != null)
            {
                switch (command.ToLower())
                {
                    case "timing":
                    case "start":
                    case "fixed":
                    case "fs":
                        value = (T)GetValue(command, key);
                        return true;
                }
            }
            value = default(T);
            return false;
        }

        public Dictionary<string, object> GetValues(string command, string[] keys)
        {
            var dict = new Dictionary<string, object>();
            keys = keys.Distinct().ToArray();
            foreach (var key in keys)
            {
                dict.Add(key, GetValue(command, key));
            }
            return dict;
        }

        public bool TryGetValues<T>(string command, string[] keys, out Dictionary<string, T> values)
        {
            values = new Dictionary<string, T>();
            keys = keys.Distinct().ToArray();
            foreach (var key in keys)
            {
                if (TryGetValue(command, key, out T value))
                {
                    values.Add(key, value);
                }
            }
            return values.Count > 0;
        }

        public void Initialize(string configDirectory, string dllDirectory, LibraryConfig config)
        {
            this.Config = config;
            this.ConfigDirectory = configDirectory;
            this.DllDirectory = dllDirectory;
            this.VoiceDirectory = Path.Combine(dllDirectory, "voice");
            var otoFile = Path.Combine(VoiceDirectory, "oto.ini");
            LoadOtoIni(otoFile);
            if (OtoList.Count == 0)
            {
                StateText = "初期化に失敗しました。oto.iniが読み込めません。" + otoFile;
                return;
            }
            StateText = "初期化されました。" + OtoList.Count + " 個の音を読み込みました。";
        }

        private void LoadOtoIni(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) ||
                !File.Exists(filePath))
            {
                return;
            }

            var lines = File.ReadAllLines(filePath, Encoding.GetEncoding("Shift_JIS"));
            OtoList.Clear();
            Parallel.ForEach(lines, line =>
            {
                var oto = Oto.CreateOto(line, VoiceDirectory);
                if (oto != null)
                {
                    lock (OtoList)
                    {
                        OtoList.Add(oto);
                    }
                }
            });
            if (OtoList.Count > 0)
            {
                fs = OtoList.First().fs;
            }
        }
    }
}
