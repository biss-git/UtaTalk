using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using VSLIB.NET;

namespace UtaSongEngine
{
    static class VoiceUtility
    {
        const int minLength = 3000;
        const int minLength2 = 100;
        const double minValue = 0.001;
        const double saveRate = 0.7;

        public static void SmoothAdd(
            List<double> waveList, double[] addWave,
            List<int> timingList, int timing,
            List<int> startList, int start,
            List<int> fixedList, int fix)
        {
            if (waveList.Count < minLength || addWave.Length < minLength)
            {
                // 十分な音声が無いときは適当につなげる
                timingList.Add(waveList.Count + timing);
                startList.Add(waveList.Count + start);
                fixedList.Add(waveList.Count + fix);
                waveList.AddRange(addWave);
                return;
            }

            bool noVoice = true;
            for (int i = 0; i < minLength2; i++)
            {
                if (Math.Abs(waveList[waveList.Count - i - 1]) > minValue ||
                    Math.Abs(addWave[i]) > minValue)
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
                waveList.AddRange(addWave);
                return;
            }


            // 音がある場所同士でつなげる場合はゼロクロスを探して良い感じでつなげる。
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
                var maxLength = sub.Max(); // 最大の距離
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
                var maxLength = sub.Max(); // 最大の距離
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

                // ここは単純に線形補完で重ね合わせているが、本当はちゃんと処理しないとダメ。
                // 単純な線形補完だと異なる周波数の似た音を足し算しているだけなのでハモってるように聞こえる。
                // まあ、それはそれで良い感じだし、処理が簡単なのでとりあえずはこれでいいかな。
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
                startList.Add(waveList.Count + Math.Max(start - overlap - removeNum, 0));
                fixedList.Add(waveList.Count + Math.Max(fix - overlap - removeNum, 0));
            }

            waveList.AddRange(addWave);
        }


    }
}
