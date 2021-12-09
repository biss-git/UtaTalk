using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VSLIB.NET;

namespace UtaTalkEngine
{
    static class VoiceUtility
    {
        const int minLength = 2000;
        const int minLength2 = 100;
        const double minValue = 0.001;
        const double saveRate = 0.7;
        const int sideLength = 1500;
        const int sideLength2 = 500;

        /// <summary>
        /// Songエンジンとは異なり、なるべく詰めてつなぎ合わせる。
        /// </summary>
        public static void SmoothAdd(
            List<double> waveList, double[] addWaveIn,
            List<int> timingList, int timing,
            List<int> startList, int start,
            List<int> fixedList, int fix)
        {
            if (waveList.Count < minLength || addWaveIn.Length < minLength)
            {
                // 十分な音声が無いときは適当につなげる
                timingList.Add(waveList.Count + timing);
                startList.Add(waveList.Count + start);
                fixedList.Add(waveList.Count + fix);
                var remove2 = Math.Max(addWaveIn.Length - fix - sideLength2, 0);
                waveList.AddRange(addWaveIn.Take(addWaveIn.Length - remove2));
                return;
            }

            bool noVoice = true;
            for (int i = 0; i < minLength2; i++)
            {
                if (Math.Abs(waveList[waveList.Count - i - 1]) > minValue ||
                    Math.Abs(addWaveIn[i]) > minValue)
                {
                    noVoice = false;
                    break;
                }
            }
            if (noVoice)
            {
                // 無音同士のときはそのままつなげる
                timingList.Add(waveList.Count + timing);
                startList.Add(waveList.Count + start);
                fixedList.Add(waveList.Count + fix);
                var remove2 = Math.Max(addWaveIn.Length - fix - sideLength2, 0);
                waveList.AddRange(addWaveIn.Take(addWaveIn.Length - remove2));
                return;
            }


            // 音がある場所同士でつなげる場合はゼロクロスを探して良い感じでつなげる。
            double[] addWave;
            {
                // 前と後ろを取り除いて、伸ばす音を短くする。
                var remove1 = Math.Max(start - sideLength, 0);
                var remove2 = Math.Max(addWaveIn.Length - fix - sideLength2, 0);
                addWave = new double[addWaveIn.Length - remove1 - remove2];
                Array.Copy(addWaveIn, remove1, addWave, 0, addWave.Length);
                timing -= remove1;
                start -= remove1;
                fix -= remove1;
            }

            {
                // waveList のおしりがゼロクロスになるように削る
                var list = new List<int>();
                for (int i = 0; i < minLength; i++)
                {
                    if (waveList[waveList.Count - i - 2] >= 0 &&
                        waveList[waveList.Count - i - 1] < 0)
                    {
                        list.Add(i + 1); // ゼロクロスの左になる点を追加している
                    }
                }
                var sub = new List<int>();
                for (int i = 0; i < list.Count - 1; i++)
                {
                    sub.Add(list[i + 1] - list[i]); // ゼロクロスの距離を算出
                }
                var maxLength = sub.Count > 0 ? sub.Max() : 0; // 最大の距離
                var limit = maxLength * saveRate;
                int removeNum = 0; // 削るサンプル数
                for (int i = 0; i < sub.Count; i++)
                {
                    if (sub[i] > limit)
                    {
                        removeNum = list[i];
                        break;
                    }
                }
                if (removeNum > 0)
                {
                    waveList.RemoveRange(waveList.Count - removeNum, removeNum);
                }
            }

            {
                // addWave の頭がゼロクロスになるように削る
                var list = new List<int>();
                for (int i = 0; i < minLength; i++)
                {
                    if (addWave[i] >= 0 &&
                        addWave[i + 1] < 0)
                    {
                        list.Add(i + 1); // ゼロクロスの右になる点を追加している
                    }
                }
                var sub = new List<int>();
                for (int i = 0; i < list.Count - 1; i++)
                {
                    sub.Add(list[i + 1] - list[i]); // ゼロクロスの距離を算出
                }
                var maxLength = sub.Count > 0 ? sub.Max() : 0; // 最大の距離
                var limit = maxLength * saveRate;
                int removeNum = 0; // 削るサンプル数
                for (int i = 0; i < sub.Count; i++)
                {
                    if (sub[i] > limit)
                    {
                        removeNum = list[i];
                        break;
                    }
                }
                if (removeNum > 0)
                {
                    addWave = addWave.ToList().GetRange(removeNum, addWave.Length - removeNum).ToArray();
                }

                int overlap = 0; // オーバーラップさせるサンプル数
                for (int i = sub.Count - 1; i >= 0; i--)
                {
                    if (sub[i] > limit)
                    {
                        overlap = list[i + 1];
                        break;
                    }
                }

                // ここは単純に線形補完で重ね合わせている
                for (int i = 0; i < overlap; i++)
                {
                    var rate = (double)i / overlap;
                    waveList[waveList.Count - overlap + i] =
                        waveList[waveList.Count - overlap + i] * (1 - rate) +
                        addWave[i] * rate;
                }
                if (overlap > 0)
                {
                    addWave = addWave.ToList().GetRange(overlap, addWave.Length - overlap).ToArray();
                }

                timingList.Add(waveList.Count + Math.Max(timing - overlap - removeNum, 0));
                startList.Add(Math.Max(waveList.Count + start - overlap - removeNum, fixedList.LastOrDefault()));
                fixedList.Add(waveList.Count + Math.Max(fix - overlap - removeNum, 0));
            }

            waveList.AddRange(addWave);
        }


    }
}
