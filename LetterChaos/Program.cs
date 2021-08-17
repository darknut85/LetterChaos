using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace LetterChaos
{
    class Program
    {
        static void Main(string[] args)
        {
            string[] row = new string[] { "Calibri", "Arial", "Century", "Cambria", "Papyrus", "Impact", "Freestyle Script", "Juuce ITC" };

            Application word = new Application();
            string file = @"C:\Users\Marco\source\repos\itvitae creations\LetterChaos\random.doc";

            Document random = word.Documents.Open(file);
            word.Visible = true;

            Console.WriteLine("Please add some text");
            //string add = "This iqs dummy text. It does not contain anything of value. But something will happen it there is more than 1 line.";
            string add = random.Content.Text;
            random.Content.Text = "";
            random.Content.Text += "";

            //////////////////////////////////////////////////

            string temp = "";
            foreach (Paragraph q in random.Paragraphs)
            {
                temp = add;
            }
            Console.WriteLine(temp);

            string writeling = temp;

            char[] ch = new char[writeling.Length];
            for (int i = 0; i < writeling.Length; i++)
            {
                ch[i] = writeling[i];
            }
            foreach (char c in ch)
            {
                Console.WriteLine(c);
            }

            //////////////////////////////////////////////////
            bool x = true;
            bool y = true;
            int d = 0;

            foreach (Paragraph p in random.Paragraphs)
            {
                Random rnm = new Random();
                Random rnn = new Random();
                int start5 = rnm.Next(11, 26);
                int start4 = rnn.Next(row.Length);
                p.Range.Font.Size = start5;
                p.Range.Font.Name = $"{row[start4]}";
                if (x == true)
                {
                    foreach (char cd in ch)
                    {
                        if (y == true)
                        {
                            p.Range.Text = "";

                            y = false;
                        }
                        p.Range.Font.Size = start5;
                        p.Range.Font.Name = $"{row[start4]}";
                        p.Range.Text += cd;
                    }
                    x = false;

                }
                d++;
            }
            for (int i = 1; i < Math.Sqrt(d); i++)
            {
                if (file != null)
                {
                    Document doc = word.Documents.Open(file, ReadOnly: false, Visible: true);
                    doc.Activate();

                    object missingObject = null;

                    doc.ConvertNumbersToText();

                    object item = WdGoToItem.wdGoToPage;
                    object whichItem = WdGoToDirection.wdGoToFirst;
                    object replaceAll = WdReplace.wdReplaceAll;
                    object forward = true;
                    object matchWholeWord = false;
                    object matchWildcards = true;
                    object matchSoundsLike = false;
                    object matchAllWordForms = false;
                    object wrap = WdFindWrap.wdFindAsk;
                    object format = true;
                    object matchCase = false;
                    object originalText = "([!.:])^13([!_-_-_-_])";
                    object replaceText = @"\1\2";

                    doc.GoTo(ref item, ref whichItem, ref missingObject, ref missingObject);
                    foreach (Range rng in doc.StoryRanges)
                    {
                        rng.Find.Font.Bold = 0;

                        rng.Find.Execute(ref originalText, ref matchCase,
                    ref matchWholeWord, ref matchWildcards, ref matchSoundsLike, ref matchAllWordForms, ref forward,
                    ref wrap, ref format, ref replaceText, ref replaceAll, ref missingObject,
                    ref missingObject, ref missingObject, ref missingObject);
                    }
                }
            }
        }
    }
}
