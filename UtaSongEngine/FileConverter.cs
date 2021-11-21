using System;
using System.Collections.Generic;
using System.Linq;
using Yomiage.SDK.Config;
using Yomiage.SDK.FileConverter;
using Yomiage.SDK.Talk;

namespace UtaSongEngine
{
    internal class FileConverter : FileConverterBase
    {

        public EngineConfig Config { get; set; }

        public override string[] ImportFilterList { get; } = new string[]
        {
            "ustファイル(UtaSong)|*.ust",
            "vsqxファイル(UtaSong)|*.vsqx",
        };

        public override (string, TalkScript[]) Open(string filepath, string filter)
        {
            List<Note> notes = null;
            if (filter.Contains("ust"))
            {
                notes = UstUtil.ReadUST(filepath);
            }
            else if (filter.Contains("vsqx"))
            {
                notes = VsqxUtil.ReadVSQX(filepath);
            }

            return MakeScripts(notes);
        }


        private (string, TalkScript[]) MakeScripts(List<Note> notes)
        {
            if (notes == null || notes.Count == 0)
            {
                return (null, null);
            }

            notes = PreProcessing(notes);

            var groups = SplitNotes(notes);

            var dict = new List<TalkScript>();

            for (int i = 0; i < groups.Count; i++)
            {
                var group = groups[i];

                var pauseScript = group.GetPauseScript(i, Config?.Key);
                if (pauseScript.Sections.First().Pause.Span_ms > 0)
                {
                    dict.Add(pauseScript);
                }

                var mainScript = group.GetMainScript(i, Config?.Key);

                if (mainScript.MoraCount > 0)
                {
                    dict.Add(mainScript);
                }
            }

            var text = "";
            foreach (var d in dict)
            {
                if (!string.IsNullOrWhiteSpace(text))
                {
                    text += Environment.NewLine;
                }
                text += d.OriginalText;
            }

            return (text, dict.ToArray());
        }

        private List<Note> PreProcessing(List<Note> notes)
        {
            List<Note> newNotes = new List<Note>();

            // 無音区間は "R" に統一
            notes.ForEach(n =>
            {
                if (n.OutputMora == "")
                {
                    n.OutputMora = "R";
                }
            });

            // Rはマージ、"ー"も足して１秒以下ならマージ
            for (int i = 0; i < notes.Count; i++)
            {
                if (i == notes.Count - 1)
                {
                    newNotes.Add(notes[i]);
                    break;
                }

                var note1 = notes[i];
                var note2 = notes[i + 1];

                // R はマージ
                if (note1.OutputMora == "R" && note2.OutputMora == "R")
                {
                    newNotes.Add(new Note()
                    {
                        Time_sec = note1.Time_sec,
                        Length_sec = note1.Length_sec + note2.Length_sec,
                        OutputMora = "R",
                        Key = note1.Key,
                        Pitch_Hz = note1.Pitch_Hz,
                    });
                    i += 1;
                    continue;
                }

                // １秒未満の伸ばし棒ならマージ
                if (note2.OutputMora == "ー" && note1.Length_sec + note2.Length_sec < 1.0 &&
                    Math.Abs(Math.Log(note1.Pitch_Hz / note2.Pitch_Hz, 2)) < 0.3)
                {
                    newNotes.Add(new Note()
                    {
                        Time_sec = note1.Time_sec,
                        Length_sec = note1.Length_sec + note2.Length_sec,
                        OutputMora = note1.OutputMora,
                        Key = note2.Key,
                        Pitch_Hz = note2.Pitch_Hz,
                    });
                    i += 1;
                    continue;
                }

                // 1秒以上なら伸ばし棒に変える
                if (note1.OutputMora != "R" && note1.Length_sec > 1.0)
                {
                    var num = (int)Math.Round(note1.Length_sec / 0.6);

                    note1.Length_sec = note1.Length_sec / num;
                    newNotes.Add(note1);

                    for (int j = 1; j < num; j++)
                    {
                        newNotes.Add(new Note()
                        {
                            Time_sec = note1.Time_sec + i * note1.Length_sec,
                            Length_sec = note1.Length_sec,
                            OutputMora = "ー",
                            Key = note1.Key,
                            Pitch_Hz = note1.Pitch_Hz,
                        });
                    }

                    continue;
                }

                newNotes.Add(note1);
            }

            return newNotes;
        }

        private List<NoteGroup> SplitNotes(List<Note> notes)
        {
            var noteGroups = new List<NoteGroup>();

            List<Note> group = new List<Note>();
            double pause_sec = 0;

            for (int i = 0; i < notes.Count; i++)
            {
                var note = notes[i];
                if (note.OutputMora == "R" && note.Length_sec > 0.2)
                {
                    if (group.Count > 0)
                    {
                        noteGroups.Add(new NoteGroup(pause_sec, group.ToArray()));
                        group.Clear();
                    }
                    pause_sec = note.Length_sec;
                    continue;
                }
                group.Add(note);
            }

            if (group.Count > 0)
            {
                noteGroups.Add(new NoteGroup(pause_sec, group.ToArray()));
                group.Clear();
            }

            return noteGroups;
        }

    }

    internal class NoteGroup
    {
        public double Pause_sec { get; set; }
        public Note[] Notes { get; set; }

        public NoteGroup(double pause_sec, Note[] notes)
        {
            Pause_sec = pause_sec;
            Notes = notes;
        }

        public TalkScript GetPauseScript(int i, string engineKey)
        {
            return new TalkScript()
            {
                OriginalText = "ポーズ＿" + i.ToString("0000") + "。",
                EngineName = engineKey,
                Sections = new List<Section>()
                    {
                        new Section()
                        {
                            Pause = new Pause()
                            {
                                Type = PauseType.Manual,
                                Span_ms = (int)Math.Round((Pause_sec - 0.2) * 1000),
                            },
                            Moras = new List<Mora>()
                            {
                                new Mora()
                                {
                                    Character = "ッ",
                                }
                            }
                        }
                    }
            };
        }

        public TalkScript GetMainScript(int i, string engineKey)
        {
            var text = "";
            foreach (var note in Notes)
            {
                if (note.OutputMora == "R")
                {
                    // 短い無音は ッ となる
                    note.OutputMora = "ッ";
                }
                text += note.OutputMora;
            }
            var mainScript = new TalkScript()
            {
                OriginalText = "フレーズ" + i.ToString("0000") + "_" + text + "。",
                EngineName = engineKey,
            };

            int pause = Math.Min(200, (int)Math.Round(Pause_sec * 1000));

            Section section = new Section()
            {
                Pause = new Pause
                {
                    Type = PauseType.Manual,
                    Span_ms = pause,
                }
            };

            foreach (var note in Notes)
            {
                section.Moras.Add(new Mora()
                {
                    Character = note.OutputMora,
                    Speed = 100 * 0.6 / note.Length_sec,
                    Pitch = Math.Log(note.Pitch_Hz / 247, 2) * 12
                }); ;
            }

            mainScript.Sections.Add(section);

            return mainScript;
        }
    }

}
