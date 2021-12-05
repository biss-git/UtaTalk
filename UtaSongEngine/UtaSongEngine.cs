using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VSLIB.NET;
using Yomiage.SDK;
using Yomiage.SDK.Config;
using Yomiage.SDK.FileConverter;
using Yomiage.SDK.Settings;
using Yomiage.SDK.Talk;
using Yomiage.SDK.VoiceEffects;

namespace UtaSongEngine
{
    public class UtaSongEngine : IVoiceEngine
    {
        public IFileConverter FileConverter { get; } = new FileConverter();

        private bool isPlaying = false;
        private bool stopFlag = false;
        private string ConfigDirectory;
        private string DllDirectory;
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

        public EngineConfig Config { get; private set; }

        public EngineSettings Settings { get; set; }

        public bool IsActivated => true;

        public bool IsEnable => !isPlaying;

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
            FileConverter?.Dispose();
        }

        public void Initialize(string configDirectory, string dllDirectory, EngineConfig config)
        {
            StateText = "初期化されました。  ";
            this.Config = config;
            this.ConfigDirectory = configDirectory;
            this.DllDirectory = dllDirectory;

            if (FileConverter is FileConverter converter)
            {
                converter.Config = config;
            }
        }

        public async Task<double[]> Play(VoiceConfig voicePreset, VoiceConfig subPreset, TalkScript talkScript, MasterEffectValue masterEffect, Action<int> setSamplingRate_Hz, Action<double[]> submitPartWave)
        {
            StateText = "再生されました。";
            if (isPlaying) { return null; }
            isPlaying = true;
            stopFlag = false;
            var moraList = new List<(string, VoiceEffectValueBase)>();
            string LastChar = string.Empty;

            fs = 44100;
            {
                if (voicePreset.Library.TryGetValue<int>("fs", "", out var fs_lib))
                {
                    fs = fs_lib;
                }
            }

            if (talkScript.MoraCount == 0)
            {
                var pauseTime = 0;
                foreach (var section in talkScript.Sections)
                {
                    if (section.Pause.Type != PauseType.None)
                    {
                        pauseTime += section.Pause.Span_ms;
                    }
                }
                if (talkScript.EndSection.Pause.Type != PauseType.None)
                {
                    pauseTime += talkScript.EndSection.Pause.Span_ms;
                }
                isPlaying = false;
                stopFlag = false;
                return new double[fs * pauseTime / 1000];
            }

            foreach (var section in talkScript.Sections)
            {
                if (LastChar != string.Empty)
                {
                    // ポーズの処理
                    if (section.Pause.Type != PauseType.None)
                    {
                        var vowel = MoraUtility.GetVowel(LastChar);
                        moraList.Add((vowel + " R", section));
                        LastChar = string.Empty;
                    }
                }
                foreach (var mora in section.Moras)
                {
                    var vowel = MoraUtility.GetVowel(LastChar);
                    moraList.Add((vowel + " " + mora.Character, mora));
                    switch (mora.Character)
                    {
                        case "ッ":
                            LastChar = string.Empty;
                            break;
                        case "ー":
                            break;
                        default:
                            LastChar = mora.Character;
                            break;
                    }
                }
            }
            if (LastChar != string.Empty)
            {
                // ポーズの処理
                var vowel = MoraUtility.GetVowel(LastChar);
                moraList.Add((vowel + " R", talkScript.EndSection));
                LastChar = string.Empty;
            }

            voicePreset.VoiceEffect.Volume ??= voicePreset.Library.Config.VolumeSetting.DefaultValue;
            voicePreset.VoiceEffect.Speed ??= voicePreset.Library.Config.SpeedSetting.DefaultValue;
            voicePreset.VoiceEffect.Pitch ??= voicePreset.Library.Config.PitchSetting.DefaultValue;
            voicePreset.VoiceEffect.Emphasis ??= voicePreset.Library.Config.EmphasisSetting.DefaultValue;
            masterEffect.Volume ??= 1;
            masterEffect.Speed ??= 1;
            masterEffect.Pitch ??= 1;
            masterEffect.Emphasis ??= 1;

            var volumeValue = voicePreset.VoiceEffect.Volume.Value * masterEffect.Volume.Value;
            var speedValue = voicePreset.VoiceEffect.Speed.Value * masterEffect.Speed.Value;
            speedValue = Math.Clamp(speedValue, 0.1, 10);
            var pitchValue = voicePreset.VoiceEffect.Pitch.Value;
            {
                // シの音との音高差を適用する
                if (voicePreset?.Library?.Settings?.Ints != null &&
                    voicePreset.Library.Settings.Ints.TryGetSetting("ShiftKey", out var setting))
                {
                    pitchValue += setting.Value / 100.0;
                }
            }

            var prePause = 0.2;
            {
                // 子音用待機時間の取得
                if (Settings?.Ints != null &&
                    Settings.Ints.TryGetSetting("prePause", out var setting))
                {
                    if (setting.Value > 0)
                    {
                        prePause = setting.Value / 1000.0;
                    }
                }
                var firstPause = talkScript.Sections.FirstOrDefault()?.Pause;
                if (firstPause?.Type != PauseType.None &&
                    firstPause?.Span_ms > prePause * 1000)
                {
                    prePause = firstPause.Span_ms / 1000.0;
                }
            }

            setSamplingRate_Hz(fs);
            var waveList = new List<double>();  // 音声波形
            var jointList_Sample = new List<int>(); // 波形のつなぎ目あたり
            var startList_Sample = new List<int>(); // オーバーラップ
            var timingList_Sample = new List<int>(); // 先行発声
            var timingEditList_sec = new List<double>(); // 先行発声（目標位置）
            var fixedList_Sample = new List<int>(); // 固定範囲
            var volumeList = new List<double>();
            var pitchList = new List<double>();
            var velocityList = new List<double>();
            var pitchCurveList = new List<double[]>();
            var formantCurveList = new List<double[]>();
            var timingEdit = prePause / speedValue;
            foreach (var mora in moraList)
            {
                if (!voicePreset.Library.TryGetValue("wave", mora.Item1, out double[] wave))
                {
                    wave = new double[0];
                }
                if (!voicePreset.Library.TryGetValue("timing", mora.Item1, out int timing))
                {
                    timing = 0;
                }
                if (!voicePreset.Library.TryGetValue("start", mora.Item1, out int start))
                {
                    start = 0;
                }
                if (!voicePreset.Library.TryGetValue("fixed", mora.Item1, out int fix))
                {
                    fix = 0;
                }
                jointList_Sample.Add(waveList.Count);
                if (mora.Item1.Contains(" ァ") ||
                    mora.Item1.Contains(" ィ") ||
                    mora.Item1.Contains(" ゥ") ||
                    mora.Item1.Contains(" ェ") ||
                    mora.Item1.Contains(" ォ"))
                {
                    // 小文字の音量は小さくする。
                    volumeList.Add((double)mora.Item2.Volume * 0.7 * volumeValue);
                }
                else
                {
                    volumeList.Add((double)mora.Item2.Volume * volumeValue);
                }
                pitchList.Add((double)mora.Item2.Pitch + pitchValue);
                velocityList.Add(100 / mora.Item2.GetAdditionalValueOrDefault("V", 100));
                pitchCurveList.Add(mora.Item2.GetAdditionalValuesOrDefault("P", 0));
                formantCurveList.Add(mora.Item2.GetAdditionalValuesOrDefault("F", 0));
                timingEditList_sec.Add(timingEdit);
                switch (mora.Item2)
                {
                    case EndSection es:
                        timingEdit += es.Pause.Span_ms / 1000.0 / speedValue;
                        if (es.Pause.Span_ms == 0)
                        {
                            // 語尾にポーズが無い場合はブツってなっちゃう可能性があるので20ms だけゆとりを持たせる。
                            timingEditList_sec[timingEditList_sec.Count - 1] -= 0.02;
                        }
                        break;
                    case Section s:
                        timingEdit += s.Pause.Span_ms / 1000.0 / speedValue;
                        break;
                    default:
                        timingEdit += 60 / ((double)mora.Item2.Speed * speedValue);
                        break;
                }
                VoiceUtility.SmoothAdd(
                    waveList, wave,
                    timingList_Sample, timing,
                    startList_Sample, start,
                    fixedList_Sample, fix);
            }

            if (waveList.Count == 0)
            {
                // 音が一切ない場合はポーズ分の時間だけ無音を返す。
                var pause_ms = 0;
                talkScript.Sections.ForEach(s =>
                {
                    pause_ms += s.Pause.Span_ms;
                });
                pause_ms += talkScript.EndSection.Pause.Span_ms;
                isPlaying = false;
                stopFlag = false;
                return new double[pause_ms * fs / 1000];
            }

            var project = new VSProject();
            {
                var fileName = Path.Combine(ConfigDirectory, "temp.wav");
                using (var writer = new WaveFileWriter(fileName, new WaveFormat(fs, 16, 1)))
                {
                    writer.WriteSamples(waveList.Select(v => (float)v).ToArray(), 0, waveList.Count);
                }
                project.ItemList.Add(fileName);

                {
                    // パラメータの加工
                    var item = project.ItemList[0];
                    var ctrlPoints = item.CtrlPoints;
                    var timeCtrlPoints = item.TimeCtrlPoints;
                    foreach (var joint in jointList_Sample) // つなぎ目の音量を滑らかにする
                    {
                        if (joint == 0) { continue; }
                        var start = joint * 200 / fs - 40;
                        var end = start + 50;
                        for (int i = start; i <= end; i++)
                        {
                            if (ctrlPoints[i].DynOrg < 0.01) { continue; }
                            var range = Math.Min(Math.Min(i - start, end - i), 10);
                            var count = 1;
                            var sum = Math.Log(ctrlPoints[i].DynOrg);
                            for (int j = 1; j <= range; j++)
                            {
                                int index1 = i + j;
                                int index2 = i - j;
                                if (ctrlPoints[index1].DynOrg < 0.01 || ctrlPoints[index2].DynOrg < 0.01)
                                {
                                    break;
                                }
                                count += 2;
                                sum += Math.Log(ctrlPoints[index1].DynOrg);
                                sum += Math.Log(ctrlPoints[index2].DynOrg);
                            }
                            ctrlPoints[i].DynEdit = Math.Exp(sum / count);
                        }
                    }

                    {
                        // 大きすぎる音はつぶす（コンプレッサー）
                        var threshold = 0.2;
                        if (voicePreset.Library.Settings?.Doubles?.ContainsKey("comp") == true)
                        {
                            threshold = Math.Clamp(voicePreset.Library.Settings.Doubles["comp"].Value, 0.01, 1);
                        }
                        for (int i = 0; i < ctrlPoints.Length; i++)
                        {
                            if (ctrlPoints[i].DynEdit > 0.2)
                            {
                                ctrlPoints[i].DynEdit = Math.Sqrt(ctrlPoints[i].DynEdit * 0.2);
                            }
                        }
                    }

                    { // 伸ばし棒のときは音量を平滑化する
                        for (int i = 1; i < moraList.Count; i++)
                        {
                            if (moraList[i].Item1.Contains("ー"))
                            {
                                var ave = 0.0;
                                var start = startList_Sample[i] * 200 / fs;
                                var end = fixedList_Sample[i] * 200 / fs;
                                if (end - start <= 0) { continue; }
                                for (int j = start; j < end; j++)
                                {
                                    ave += Math.Log(ctrlPoints[j].DynEdit);
                                }
                                ave = Math.Exp(ave / (end - start));
                                for (int j = start; j < end; j++)
                                {
                                    ctrlPoints[j].DynEdit = ave;
                                }
                                if (!moraList[i - 1].Item1.Contains("ー")) { continue; } // ーが連続している場合。
                                var pre = fixedList_Sample[i - 1] * 200 / fs;
                                for (int j = pre; j < start; j++)
                                {
                                    ctrlPoints[j].DynEdit = ave;
                                }
                            }
                            else
                            {
                                if (!moraList[i - 1].Item1.Contains("ー")) { continue; }
                                // ーの終わり
                                var pre = fixedList_Sample[i - 1] * 200 / fs;
                                var start = startList_Sample[i] * 200 / fs;
                                if (start - pre <= 0) { continue; }
                                var dyn = ctrlPoints[pre - 1].DynEdit;
                                for (int j = pre; j < start; j++)
                                {
                                    ctrlPoints[j].DynEdit = (dyn * (start - j) + ctrlPoints[j].DynEdit * (j - pre)) / (start - pre);
                                }
                            }
                        }
                    }

                    {
                        // ピッチの平滑化
                        const int l = 10;
                        for (int i = l; i < ctrlPoints.Length - l; i++)
                        {
                            var pitch = 0;
                            var num = 0;
                            if (ctrlPoints[i].PitOrg <= 2000) { continue; }
                            for (int j = i - l; j < i + l; j++)
                            {
                                if (ctrlPoints[j].PitOrg > 2000)
                                {
                                    pitch += ctrlPoints[j].PitOrg;
                                    num += 1;
                                }
                            }
                            if (num > 1)
                            {
                                ctrlPoints[i].PitEdit = pitch / num;
                            }
                        }

                    }

                    // 音高の平滑化（ケロケロ化）
                    {
                        // 基準音高を取得
                        var BaseKey = 5900;
                        {
                            if (voicePreset?.Library?.Settings?.Ints != null &&
                                voicePreset.Library.Settings.Ints.TryGetSetting("BaseKey", out var setting))
                            {
                                BaseKey = setting.Value;
                            }
                        }
                        {
                            if (Settings?.Doubles != null &&
                                Settings.Doubles.TryGetSetting("kero", out var setting) == true)
                            {
                                var rate = Math.Clamp(setting.Value, 0, 1);
                                for (int i = 0; i < ctrlPoints.Length; i++)
                                {
                                    if (ctrlPoints[i].PitEdit < 10) { continue; }
                                    ctrlPoints[i].PitEdit = (int)Math.Round((1 - rate) * ctrlPoints[i].PitEdit + rate * BaseKey);
                                }
                            }
                        }
                    }

                    {
                        // 音量を適用する
                        var volume = volumeList[0];
                        var position = 0;
                        for (int i = 1; i < timingList_Sample.Count; i++)
                        {
                            var start = startList_Sample[i] * 200 / fs;
                            for (; position < start; position++)
                            {
                                ctrlPoints[position].Volume = volume;
                            }
                            var distVolume = volumeList[i];
                            var rate = 0.0;
                            var timing = timingList_Sample[i] * 200 / fs;
                            for (; position < timing; position++)
                            {
                                rate = (double)(position - start) / (timing - start);
                                ctrlPoints[position].Volume = distVolume * rate + volume * (1 - rate);
                            }
                            volume = distVolume;
                        }
                        for (; position < ctrlPoints.Length; position++)
                        {
                            ctrlPoints[position].Volume = volume;
                        }
                    }

                    {
                        // ピッチを適用する
                        var pitch = pitchList[0];
                        var position = 0;
                        for (int i = 1; i < timingList_Sample.Count; i++)
                        {
                            var start = (startList_Sample[i] + timingList_Sample[i]) / 2 * 200 / fs;
                            for (; position < start; position++)
                            {
                                ctrlPoints[position].PitEdit = PitchShift(ctrlPoints[position].PitEdit, pitch);
                            }
                            var distPitch = pitchList[i];
                            var rate = 0.0;
                            var timing = timingList_Sample[i] * 200 / fs;
                            for (; position < timing; position++)
                            {
                                rate = (double)(position - start) / (timing - start);
                                ctrlPoints[position].PitEdit = PitchShift(ctrlPoints[position].PitEdit, distPitch * rate + pitch * (1 - rate));
                            }
                            pitch = distPitch;
                        }
                        for (; position < ctrlPoints.Length; position++)
                        {
                            ctrlPoints[position].PitEdit = PitchShift(ctrlPoints[position].PitEdit, pitch);
                        }
                    }

                    {
                        var endTimingEdit_Sample = (int)Math.Round(timingEdit * fs);
                        timeCtrlPoints[1].time2 = endTimingEdit_Sample;
                        // 話速からタイミングを合わせる
                        {
                            var preTime = (int)Math.Round(timingEditList_sec[0] * fs) - timingList_Sample[0];
                            if (preTime > 0)
                            {
                                timeCtrlPoints.Add(1, preTime);
                            }
                        }
                        var timingEdits_Sample = new List<int>();
                        var startEdits_Sample = new List<int>();
                        var fixedEdits_Sample = new List<int>();
                        for (int i = 0; i < timingList_Sample.Count; i++)
                        {
                            var timingEdit_Sample = (int)Math.Round(timingEditList_sec[i] * fs);
                            timingEdits_Sample.Add(timingEdit_Sample);
                            startEdits_Sample.Add(timingEdit_Sample - (int)((timingList_Sample[i] - startList_Sample[i]) * velocityList[i]));
                            fixedEdits_Sample.Add(timingEdit_Sample + (int)((fixedList_Sample[i] - timingList_Sample[i]) * velocityList[i]));
                        }
                        {
                            startEdits_Sample[0] = Math.Max(startEdits_Sample[0], 0);
                        }
                        for (int i = 1; i < timingList_Sample.Count; i++)
                        {
                            if (startEdits_Sample[i] - fixedEdits_Sample[i - 1] < 1000)
                            {
                                var dt = timingEdits_Sample[i] - timingEdits_Sample[i - 1];
                                var d1 = fixedEdits_Sample[i - 1] - timingEdits_Sample[i - 1];
                                var d2 = timingEdits_Sample[i] - startEdits_Sample[i];
                                var dt1 = dt * d1 / (d1 + d2);
                                var middle = timingEdits_Sample[i - 1] + dt1;
                                startEdits_Sample[i] = middle + 500;
                                fixedEdits_Sample[i - 1] = middle - 500;
                            }
                        }
                        {
                            fixedEdits_Sample[fixedEdits_Sample.Count - 1] = Math.Min(fixedEdits_Sample[fixedEdits_Sample.Count - 1], endTimingEdit_Sample - 500);
                        }
                        for (int i = 0; i < timingList_Sample.Count; i++)
                        {
                            timeCtrlPoints.Add(timingList_Sample[i], timingEdits_Sample[i]);
                            timeCtrlPoints.Add(startList_Sample[i], startEdits_Sample[i]);
                            timeCtrlPoints.Add(fixedList_Sample[i], fixedEdits_Sample[i]);
                            //var timingEdit_Sample = (int)Math.Round(timingEditList_sec[i] * fs);
                            //timeCtrlPoints.Add(timingList_Sample[i], timingEdit_Sample);
                            //timeCtrlPoints.Add(startList_Sample[i], timingEdit_Sample - (int)((timingList_Sample[i] - startList_Sample[i]) * velocityList[i]));
                            //timeCtrlPoints.Add(fixedList_Sample[i], timingEdit_Sample + (int)((fixedList_Sample[i] - timingList_Sample[i]) * velocityList[i]));
                        }
                    }

                    // オートビブラート
                    {
                        var vDepth = voicePreset.VoiceEffect.GetAdditionalValueOrDefault("VibratoDepth") / 100; // cent to キー にしておく
                        var vSpan_sample = voicePreset.VoiceEffect.GetAdditionalValueOrDefault("VibratoSpan", 200) / 1000.0 * fs;
                        var vStart_sample = item.GetTimeEdt_Sample(fixedList_Sample[0]);
                        for (int i = 1; i < moraList.Count; i++)
                        {
                            // 本当はフェード入れたいけど面倒だからいいや。
                            if (!moraList[i - 1].Item1.Contains(" ー"))
                            {
                                var timingEdit_Sample = (int)Math.Round(timingEditList_sec[i] * fs);
                                vStart_sample = item.GetTimeEdt_Sample(fixedList_Sample[i]);
                                var start = (int)Math.Ceiling(fixedList_Sample[i - 1] * 200.0 / fs);
                                start = Math.Max(start, 0);
                                var end = timingList_Sample[i] * 200.0 / fs;
                                end = Math.Min(end, ctrlPoints.Length);
                                for (int j = start; j < end; j++)
                                {
                                    var sample = item.GetTimeEdt_Sample(j * fs / 200.0);
                                    var angle = (sample - vStart_sample) * 2 * Math.PI / vSpan_sample;
                                    ctrlPoints[j].PitEdit = PitchShift(ctrlPoints[j].PitEdit, vDepth * Math.Sin(angle));
                                }
                            }
                            else
                            {
                                var start = (int)Math.Ceiling(timingList_Sample[i - 1] * 200.0 / fs);
                                start = Math.Max(start, 0);
                                var end = (int)Math.Ceiling(timingList_Sample[i] * 200.0 / fs);
                                end = Math.Min(end, ctrlPoints.Length);
                                for (int j = start; j < end; j++)
                                {
                                    var sample = item.GetTimeEdt_Sample(j * fs / 200.0);
                                    var angle = (sample - vStart_sample) * 2 * Math.PI / vSpan_sample;
                                    ctrlPoints[j].PitEdit = PitchShift(ctrlPoints[j].PitEdit, vDepth * Math.Sin(angle));
                                }
                            }
                        }
                    }

                    // フォルマント揺らぎ
                    {
                        var formant = voicePreset.VoiceEffect.GetAdditionalValueOrDefault("Formant");
                        for (int i = 0; i < ctrlPoints.Length; i++)
                        {
                            ctrlPoints[i].Formant = (int)Math.Round(formant * Math.Sin(i / 10.0));
                        }
                    }

                    {
                        // ピッチ補正を適用する
                        for (int i = 1; i < pitchCurveList.Count; i++)
                        {
                            var pitchCurve = pitchCurveList[i - 1];
                            var startSample = timingList_Sample[i - 1];
                            var startSampleEdit = item.GetTimeEdt_Sample(startSample);
                            var start = startSample * 200 / fs;
                            var endSample = timingList_Sample[i];
                            var endSampleEdit = item.GetTimeEdt_Sample(endSample);
                            var end = endSample * 200 / fs;
                            for (int j = start; j < end; j++)
                            {
                                var sampleEdit = item.GetTimeEdt_Sample(j * fs / 200);
                                var index = (sampleEdit - startSampleEdit) / (endSampleEdit - startSampleEdit) * 10;
                                var pitchShift = Interpolation(pitchCurve, index);
                                ctrlPoints[j].PitEdit = PitchShift(ctrlPoints[j].PitEdit, pitchShift / 100);
                            }
                        }
                    }
                    {
                        // フォルマント補正を適用する
                        for (int i = 1; i < formantCurveList.Count; i++)
                        {
                            var formantCurve = formantCurveList[i - 1];
                            var startSample = timingList_Sample[i - 1];
                            var startSampleEdit = item.GetTimeEdt_Sample(startSample);
                            var start = startSample * 200 / fs;
                            var endSample = timingList_Sample[i];
                            var endSampleEdit = item.GetTimeEdt_Sample(endSample);
                            var end = endSample * 200 / fs;
                            for (int j = start; j < end; j++)
                            {
                                var sampleEdit = item.GetTimeEdt_Sample(j * fs / 200);
                                var index = (sampleEdit - startSampleEdit) / (endSampleEdit - startSampleEdit) * 10;
                                var formantShift = (int)Interpolation(formantCurve, index);
                                ctrlPoints[j].Formant += formantShift;
                            }
                        }
                    }

                }

                var projectName = Path.Combine(ConfigDirectory, "temp.vshp");
                project.Save(projectName);
            }


            isPlaying = false;
            stopFlag = false;
            (int[] dataL, int[] dataR) = project.ReadMixData(1, 0, 100000000);
            {
                // 最初の無音区間を無音として扱う（たまに異音が入るので）
                var startSample = timingEditList_sec.FirstOrDefault() * fs;
                startSample -= (int)((timingList_Sample[0] - startList_Sample[0]) * velocityList[0]);
                for (int i = 0; i < startSample; i++)
                {
                    dataL[i] = 0;
                }
            }

            var result = new List<double>();
            result.AddRange(dataL.Select(v => (double)v / 32768));

            // 歌では文末ポーズは適用しない
            //var endPauseSampleNum = (int)(fs * masterEffect.EndPause / 1000);
            //if (endPauseSampleNum > 0) { result.AddRange(new double[endPauseSampleNum]); }

            return result.ToArray();
        }

        private static int PitchShift(int pitOrg, double keyShift)
        {
            if (pitOrg < 10) { return pitOrg; }
            var shift = (int)Math.Round(keyShift * 100);
            return pitOrg + shift;
        }

        private static double Interpolation(double[] values, double index)
        {
            if (index <= 0) { return values[0]; }
            if (index >= values.Length - 1) { return values.Last(); }
            int i1 = (int)Math.Floor(index);
            int i2 = (int)Math.Ceiling(index);
            if (i1 == i2) { return values[i1]; }
            var rate = index - i1;
            return values[i2] * rate + values[i1] * (1 - rate);
        }

        public async Task Save(VoiceConfig voicePreset, VoiceConfig subPreset, TalkScript[] talkScripts, MasterEffectValue masterEffect, string filePath, int startPause, int endPause, bool saveWithText, Encoding encoding)
        {
            var fs = 44100;
            var waveList = new List<double>();

            int startPauseSample = fs * startPause / 1000;
            if (startPauseSample > 0) { waveList.AddRange(new double[startPauseSample]); }

            foreach (var script in talkScripts)
            {
                var wave = await Play(voicePreset, subPreset, script, masterEffect, x => fs = x, x => { });
                if (wave != null && wave.Length > 0)
                    waveList.AddRange(wave);
            }

            int endPauseSample = fs * endPause / 1000;
            if (endPauseSample > 0) { waveList.AddRange(new double[endPauseSample]); }


            SaveWav(waveList.ToArray(), filePath, fs);

            if (saveWithText)
            {
                var textPath = Path.ChangeExtension(filePath, ".txt");
                var text = "";
                foreach (var s in talkScripts)
                {
                    text += s.OriginalText;
                }
                SaveText(textPath, text, encoding);
            }

            if (Settings?.Bools != null &&
                Settings.Bools.TryGetSetting("withVshp", out var setting) &&
                setting.Value)
            {
                var project = new VSProject();
                project.ItemList.Add(filePath);
                var projectName = Path.ChangeExtension(filePath, ".vshp");
                project.Save(projectName);
            }
        }

        private void SaveWav(double[] Dbuffer, string filePath, int fs)
        {
            using (var writer = new WaveFileWriter(filePath, new WaveFormat(fs, 16, 1)))
            {
                writer.WriteSamples(Dbuffer.Select(v => (float)v).ToArray(), 0, Dbuffer.Length);
            }
        }

        private void SaveText(string filePath, string text, Encoding enc)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
                File.WriteAllText(filePath, text, enc);
            }
            catch
            {

            }
        }

        public async Task<bool> Stop()
        {
            if (!isPlaying) { return true; }
            stopFlag = true;
            for (int i = 0; i < 10; i++)
            {
                await Task.Delay(100);
                if (!stopFlag) { break; }
            }
            return !stopFlag;
        }

        public async Task<TalkScript> GetDictionary(string text)
        {
            return null;
        }
    }
}
