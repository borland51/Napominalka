using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace hakaton
{
    class TrshConfig
    {
        static public byte SettType = 0;
        static public string SettFile = null;
        static public List<int> SettDays = new List<int>();

        const string CFG = "config.ini";

        static string GetCOnfigPath(string file) 
        {
            return Path.Combine(Application.StartupPath, file);
        }

        static public byte GetType(string path)
        {
            byte res = 0;

            if (File.Exists(path))
                res = 1;
            else if (Directory.Exists(path))
                res = 2;

            return res;
        }

        static public bool GetConfig()
        {
            try
            {
                string iFile = GetCOnfigPath(CFG);
                if (!File.Exists(iFile)) 
                {
                    CreateConfig();
                    return false;
                }

                FileStream file = new FileStream(iFile, FileMode.Open, FileAccess.Read);
                StreamReader reader = new StreamReader(file, Encoding.UTF8);

                while (!reader.EndOfStream)
                {
                    string line = reader.ReadLine();
                    line = line.Trim();

                    if (String.IsNullOrWhiteSpace(line))
                        continue;

                    string[] kv = line.Split('=');
                    kv[0] = kv[0].Trim();
                    kv[1] = kv[1].Trim();

                    if (String.Compare(kv[0], "path") == 0)
                    {
                        SettFile = String.Copy(kv[1]);
                        SettType = GetType(SettFile);
                    } 
                    else if (String.Compare(kv[0], "days") == 0)
                    {
                        string[] days = kv[1].Split(',');
                        int cnt = days.Count();

                        for (int i = 0; i < cnt; i++)
                            SettDays.Add(Convert.ToInt32(days[i]));
                    }
                }

                reader.Close();
                file.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка!");
            }
            return false;
        }

        static public bool CreateConfig(bool reCreate = false) 
        {
            try 
            {
                FileStream file = new FileStream(GetCOnfigPath(CFG), reCreate ? FileMode.Truncate : FileMode.CreateNew, FileAccess.Write);
                StreamWriter writer = new StreamWriter(file, Encoding.UTF8);

                if (reCreate)
                {
                    writer.WriteLine("path = " + SettFile);
                    SettType = GetType(SettFile);

                    writer.Write("days = ");
                    for(int i = 0; i < SettDays.Count; i ++)
                        writer.Write((i > 0 ? "," : "") + SettDays[i].ToString());
                    writer.WriteLine();
                }
                else 
                {
                    writer.WriteLine("path = ");
                    writer.WriteLine("days = 15,45");

                    SettDays.Add(15);
                    SettDays.Add(45);
                }

                writer.Close();
                file.Close();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Ошибка!");
            }
            return false;
        }
    }
}
