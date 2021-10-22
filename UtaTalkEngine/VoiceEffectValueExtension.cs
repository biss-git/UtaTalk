using System;
using System.Collections.Generic;
using System.Text;
using Yomiage.SDK.VoiceEffects;

namespace UtaTalkEngine
{
    static class VoiceEffectValueExtension
    {
        public static void SetAccent(this VoiceEffectValueBase voiceEffectValueBase, double Accent)
        {
            voiceEffectValueBase.SetAdditionalValue("Accent", Accent);
        }
        public static double GetAccent(this VoiceEffectValueBase voiceEffectValueBase)
        {
            return voiceEffectValueBase.GetAdditionalValueOrDefault("Accent", 0);
        }
    }
}
