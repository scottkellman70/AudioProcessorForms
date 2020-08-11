using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;
using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace AudioProcessorForms
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog1.SelectedPath;
                CreateCSV(folderBrowserDialog1.SelectedPath);
            }

            //
        }
        static void CreateCSV(string directory)
        {
            using (var stream = new StreamWriter(@"c:\cases\results\vocab.csv", false, Encoding.UTF8))
            {
                foreach (var f in new DirectoryInfo(directory).GetFiles())
                {
                    using (var so = ShellObject.FromParsingName(f.FullName))
                    {
                        //
                        // The Title property contains the vocabulary separated by " - "
                        //
                        var title = so.Properties.GetProperty(SystemProperties.System.Title).ToString();
                        if (title.Contains(" - "))
                        {
                            var index = title.IndexOf(" - ");
                            var left = title.Substring(0, index - 1);
                            var right = title.Substring(index + 3);

                            //
                            // The Album property stores the category (e.g. Food and Drink).
                            //
                            var category = so.Properties.GetProperty(SystemProperties.System.Music.AlbumTitle).ValueAsObject.ToString().Substring(3);
                            //
                            // The Artist property stores the type of vocabulary (verb, noun, ...).
                            //
                            var kind = so.Properties.GetProperty(SystemProperties.System.Music.DisplayArtist).ValueAsObject.ToString();

                            var line = string.Format("{0},{1},{2},{3}", left, right, category, kind);
                            stream.WriteLine(line);
                        }
                        else
                        {
                            //
                            // The Album property stores the category (e.g. Food and Drink).
                            //
                            var category = so.Properties.GetProperty(SystemProperties.System.Music.AlbumTitle).ValueAsObject.ToString().Substring(3);
                            //
                            // The Artist property stores the type of vocabulary (verb, noun, ...).
                            //
                            var kind = so.Properties.GetProperty(SystemProperties.System.Music.DisplayArtist).ValueAsObject.ToString();

                            var line = string.Format("{0},{1},{2}", title, category, kind);
                            stream.WriteLine(line);

                        }
                    }
                }
            }
        }
    }
}
