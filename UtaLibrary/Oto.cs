using Microsoft.International.Converters;
using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace UtaLibrary
{
    class Oto
    {
        /// <summary>
        /// 音
        /// ファイル名
        /// </summary>
        public string WavPath;
        /// <summary>
        /// エイリアス
        /// "i あ" みたいなやつ
        /// </summary>
        public string Alias;
        /// <summary>
        /// 左ブランク
        /// 単位：ms
        /// </summary>
        public double LeftOffset;
        /// <summary>
        /// オーバーラップ
        /// </summary>
        public double Overlap;
        public int Overlap_Sample => (int)Math.Round(fs * Overlap / 1000);
        /// <summary>
        /// 先行発音
        /// </summary>
        public double Preceding;
        public int Preceding_Sample => (int)Math.Round(fs * Preceding / 1000);
        /// <summary>
        /// 固定範囲
        /// 子音の終了位置。
        /// </summary>
        public double Fixed;
        public int Fixed_Sample => (int)Math.Round(fs * Fixed / 1000);
        /// <summary>
        /// 右ブランク
        /// 正の場合は右ブランク時間[ms]
        /// 負の場合は有効音声時間[ms]
        /// </summary>
        public double RightOffset;
        /// <summary>
        /// サンプリング周波数[Hz]
        /// ファイル読み込み時に再設定される
        /// </summary>
        public int fs = 44100;

        public double[] GetWave(double breathVolume)
        {
            if (!File.Exists(this.WavPath)) { return new double[0]; }
            using var reader = new WaveFileReader(this.WavPath);
            var wave = new List<double>();
            int fs = reader.WaveFormat.SampleRate;
            int start = Math.Max((int)Math.Round(fs * this.LeftOffset / 1000), 0);
            long end = (long)Math.Round(((this.RightOffset <= 0) ? start : reader.Length / 2) - fs * this.RightOffset / 1000);
            start *= reader.BlockAlign;
            end *= reader.BlockAlign;
            end = Math.Min(end, reader.Length);
            reader.Position = start;
            while (reader.Position < end)
            {
                var samples = reader.ReadNextSampleFrame();
                wave.Add(samples.First());
            }

            if (Alias.Contains(" R"))
            {
                for (int i = Preceding_Sample; i < wave.Count; i++)
                {
                    wave[i] *= breathVolume;
                }
            }

            return wave.ToArray();
        }

        public static Oto CreateOto(string line, string directory)
        {
            var oto = new Oto();
            var texts = line.Split(',');
            if (texts.Length < 6) { return null; } // 要素が足りてない
            var otoAlias = texts[0].Split('=');
            if (otoAlias.Length != 2) { return null; } // 要素が足りてない
            oto.WavPath = Path.Combine(directory, otoAlias[0]);
            if (!File.Exists(oto.WavPath)) { return null; }
            oto.Alias = KanaConverter.HiraganaToKatakana(otoAlias[1]);
            if (string.IsNullOrWhiteSpace(oto.Alias)) { return null; } // エイリアスがない
            if (!double.TryParse(texts[1], out var left)) { return null; }
            oto.LeftOffset = left;
            if (!double.TryParse(texts[2], out var fix)) { return null; }
            oto.Fixed = fix;
            if (!double.TryParse(texts[3], out var right)) { return null; }
            oto.RightOffset = right;
            if (!double.TryParse(texts[4], out var pre)) { return null; }
            oto.Preceding = pre;
            if (!double.TryParse(texts[5], out var over)) { return null; }
            oto.Overlap = over;

            // サンプリングレート計算
            using var reader = new WaveFileReader(oto.WavPath);
            oto.fs = reader.WaveFormat.SampleRate;


            return oto;
        }

    }
}
