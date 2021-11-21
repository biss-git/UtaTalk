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

namespace UtaTalkEngine
{
    public class UtaTalkEngine : IVoiceEngine
    {
        public IFileConverter FileConverter { get; }
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
        }

        public void Initialize(string configDirectory, string dllDirectory, EngineConfig config)
        {
            StateText = "初期化されました。  ";
            this.Config = config;
            this.ConfigDirectory = configDirectory;
            this.DllDirectory = dllDirectory;
        }

        public async Task<double[]> Play(VoiceConfig voicePreset, VoiceConfig subPreset, TalkScript talkScript, MasterEffectValue masterEffect, Action<int> setSamplingRate_Hz)
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
                var isSecondSection =
                    section != talkScript.Sections.FirstOrDefault() &&
                    section.Pause.Type == PauseType.None; // ２番目以降のアクセント句ではちょっとピッチを下げる。そのためのフラグ。
                foreach (var mora in section.Moras)
                {
                    mora.Volume = section.Volume;
                    mora.Pitch = section.Pitch;
                    mora.Speed = section.Speed * (mora == section.Moras.Last() ? 0.98 : 1);
                    mora.Emphasis = section.Emphasis;
                    {
                        // アクセントの計算
                        var accentValue =
                            (mora == section.Moras.FirstOrDefault() || mora == section.Moras.Last()) ? 0.5 : 1.5
                            + (isSecondSection ? -1 : 0);
                        if (section.Moras.All(m => m.Accent == section.Moras.FirstOrDefault().Accent))
                        {
                            // h タイプ
                            // 特になし
                        }
                        else if (section.Moras.FirstOrDefault()?.Accent == true && !section.Moras.FirstOrDefault()?.Accent == true)
                        {
                            // hl タイプ
                            if (mora.Accent)
                            {
                                accentValue += 2;
                            }
                            else
                            {
                                var index = section.Moras.IndexOf(mora);
                                var num = section.Moras.Count(m => !m.Accent);
                                accentValue += -2 + 4.0 * (section.Moras.Count - index - 1) / num;
                            }
                        }
                        else if (!section.Moras.FirstOrDefault()?.Accent == true && section.Moras.FirstOrDefault()?.Accent == true)
                        {
                            // lh タイプ
                            if (mora.Accent)
                            {
                                accentValue += 0;
                            }
                            else
                            {
                                var index = section.Moras.IndexOf(mora);
                                var num = section.Moras.Count(m => !m.Accent);
                                accentValue += -2 + 2.0 * index / num;
                            }
                        }
                        else
                        {
                            // lhl タイプ
                            if (mora.Accent)
                            {
                                accentValue += 2;
                            }
                            else if (section.Moras.Count > 0)
                            {
                                var hindex = section.Moras.IndexOf(section.Moras.First(m => m.Accent));
                                var index = section.Moras.IndexOf(mora);
                                if (index < hindex)
                                {
                                    accentValue += -2 + 4.0 * index / hindex;
                                }
                                else
                                {
                                    var num = section.Moras.Count(m => !m.Accent);
                                    accentValue += -2 + 4.0 * (section.Moras.Count - index - 1) / (num - hindex);
                                }
                            }
                        }
                        if (section.Moras.FirstOrDefault() == mora && section.Moras.Count > 1 && section.Moras[0].Accent && section.Moras[1].Accent)
                        {
                            accentValue -= 1;
                        }
                        if (section.Moras.Count >= 5)
                        {
                            var index = section.Moras.IndexOf(mora);
                            accentValue += (double)index / section.Moras.Count;
                        }
                        mora.SetAccent(accentValue);
                    }
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

            if (talkScript.EndSection.EndSymbol == "？")
            {
                talkScript.EndSection.SetAccent(10);
                var lastMora = talkScript.Sections.Last().Moras.Last();
                lastMora.SetAccent(lastMora.GetAccent() + 2);
            }

            var volumeValue = voicePreset.VoiceEffect.Volume.Value * masterEffect.Volume.Value;
            var speedValue = voicePreset.VoiceEffect.Speed.Value * masterEffect.Speed.Value;
            speedValue = Math.Clamp(speedValue, 0.1, 10);

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
                pitchList.Add(Math.Log((double)mora.Item2.Pitch, 2) * 12);
                velocityList.Add(1.0 / Math.Sqrt(((double)mora.Item2.Speed + 1) / 2));
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
                        // timingEdit += 0.85 * MoraUtility.GetMoraSpan_ms(mora.Item1) / 1000.0 / ((double)mora.Item2.Speed * speedValue);
                        // timingEdit += 0.85 * 200 / 1000.0 / ((double)mora.Item2.Speed * speedValue);
                        // timingEdit += 0.85 * (MoraUtility.GetMoraSpan_ms(mora.Item1) + 200) / 2 / 1000.0 / ((double)mora.Item2.Speed * speedValue);
                        var index = moraList.IndexOf(mora);
                        if (index + 1 < moraList.Count)
                        {
                            // 次の発音の子音に依存するので、次の発音の長さをやや重視する。
                            timingEdit += 0.85 * (MoraUtility.GetMoraSpan_ms(mora.Item1) + MoraUtility.GetMoraSpan_ms(moraList[index + 1].Item1) + 1 * 200) / 3 / 1000.0 / ((double)mora.Item2.Speed * speedValue);
                        }
                        else
                        {
                            timingEdit += 0.85 * (MoraUtility.GetMoraSpan_ms(mora.Item1) + 1 * 200) / 2 / 1000.0 / ((double)mora.Item2.Speed * speedValue);
                        }
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
                            ctrlPoints[i].DynEdit = (ctrlPoints[i].DynEdit + Math.Exp(sum / count)) / 2;
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
                                    ctrlPoints[j].DynEdit = (ctrlPoints[j].DynOrg + ave) / 2;
                                }
                                if (!moraList[i - 1].Item1.Contains("ー")) { continue; } // ーが連続している場合。
                                var pre = fixedList_Sample[i - 1] * 200 / fs;
                                for (int j = pre; j < start; j++)
                                {
                                    ctrlPoints[j].DynEdit = (ctrlPoints[j].DynOrg + ave) / 2;
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
                        const int l = 30;
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
                        // 音量の順序が入れ替わっているときに修正する
                        for (int i = 1; i < timingList_Sample.Count; i++)
                        {
                            //if (startEdits_Sample[i] - fixedEdits_Sample[i - 1] < 1000)
                            var mora = moraList[i].Item1;
                            if (!mora.Contains("- ") ||
                                startEdits_Sample[i] - fixedEdits_Sample[i - 1] < 300)
                            {
                                var dt = timingEdits_Sample[i] - timingEdits_Sample[i - 1];
                                var d1 = fixedEdits_Sample[i - 1] - timingEdits_Sample[i - 1];
                                var d2 = timingEdits_Sample[i] - startEdits_Sample[i];
                                var dt1 = dt * d1 / (d1 + d2);
                                var middle = timingEdits_Sample[i - 1] + dt1;
                                startEdits_Sample[i] = middle;
                                fixedEdits_Sample[i - 1] = middle;
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

                    // 抑揚とアクセント
                    {
                        var accentList = moraList.Select(m => m.Item2.GetAccent()).ToArray();
                        // ピッチを適用する
                        // var pitch = moraList[0].Item2.GetAdditionalValueOrDefault("Accent", 0) * (moraList[0].Item2.Emphasis.HasValue ? (double)moraList[0].Item2.Emphasis : 1);
                        var pitch = 0.0;
                        var position = startList_Sample[0] * 200 / fs;
                        for (int i = 1; i < timingList_Sample.Count; i++)
                        {
                            var start = position;
                            // var distPitch = moraList[i].Item2.GetAdditionalValueOrDefault("Accent", 0) * (moraList[i].Item2.Emphasis.HasValue ? (double)moraList[i].Item2.Emphasis : 1);
                            var distPitch = moraList[i - 1].Item2.GetAccent() * (moraList[i - 1].Item2.Emphasis.HasValue ? (double)moraList[i - 1].Item2.Emphasis : 1);
                            var center2 = (timingList_Sample[i] + timingList_Sample[i - 1]) / 2 * 200 / fs;
                            var center = (startList_Sample[i] + fixedList_Sample[i - 1]) / 2 * 200 / fs;
                            var t1 = item.GetTimeEdt_Sec(start * 0.005);
                            var t2 = item.GetTimeEdt_Sec(center * 0.005);
                            for (; position < center; position++)
                            {
                                var t = item.GetTimeEdt_Sec(position * 0.005);
                                var rate = (double)(t - t1) / (t2 - t1);
                                var key = distPitch * rate + pitch * (1 - rate);
                                ctrlPoints[position].PitEdit = PitchShift(ctrlPoints[position].PitEdit, key);
                                ctrlPoints[position].Formant = (int)(10 * key);
                            }
                            pitch = distPitch;
                        }
                        {
                            var start = position;
                            var distPitch = talkScript.EndSection.GetAccent();
                            var center = ctrlPoints.Length;
                            var t1 = item.GetTimeEdt_Sec(start * 0.005);
                            var t2 = item.GetTimeEdt_Sec(center * 0.005);
                            for (; position < ctrlPoints.Length; position++)
                            {
                                var t = item.GetTimeEdt_Sec(position * 0.005);
                                var rate = (double)(t - t1) / (t2 - t1);
                                var key = distPitch * rate + pitch * (1 - rate);
                                ctrlPoints[position].PitEdit = PitchShift(ctrlPoints[position].PitEdit, key);
                                ctrlPoints[position].Formant = (int)(10 * key);
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
            var endPauseSampleNum = (int)(fs * masterEffect.EndPause / 1000);
            if (endPauseSampleNum > 0) { result.AddRange(new double[endPauseSampleNum]); }

            return result.ToArray();
        }

        private static int PitchShift(int pitOrg, double keyShift)
        {
            if (pitOrg < 10) { return pitOrg; }
            var shift = (int)Math.Round(keyShift * 100);
            return pitOrg + shift;
        }

        public async Task Save(VoiceConfig voicePreset, VoiceConfig subPreset, TalkScript[] talkScripts, MasterEffectValue masterEffect, string filePath, int startPause, int endPause, bool saveWithText, Encoding encoding)
        {
            var fs = 44100;
            var waveList = new List<double>();

            int startPauseSample = fs * startPause / 1000;
            if (startPauseSample > 0) { waveList.AddRange(new double[startPauseSample]); }

            foreach (var script in talkScripts)
            {
                var wave = await Play(voicePreset, subPreset, script, masterEffect, x => fs = x);
                if (wave != null && wave.Length > 0)
                {
                    waveList.AddRange(wave);
                }
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

    }
}
