using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Collections;
using System.Xml;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Net;


namespace Lexiconer
{
    public partial class Form1 : Form
    {
       

        Dictionary<string, int> words = new Dictionary<string, int>();   // все слова

        Dictionary<string, int> adverbs = new Dictionary<string, int>();
        Dictionary<string, int> verbs = new Dictionary<string, int>();
        Dictionary<string, int> nouns = new Dictionary<string, int>();
        Dictionary<string, int> adjectives = new Dictionary<string, int>();
        Dictionary<string, int> numerals = new Dictionary<string, int>();
        Dictionary<string, int> verbaladverbs = new Dictionary<string, int>();
        Dictionary<string, int> hyperlinks = new Dictionary<string, int>();
        Dictionary<string, int> particles = new Dictionary<string, int>();
        Dictionary<string, int> preposition = new Dictionary<string, int>();
        Dictionary<string, int> pronounce = new Dictionary<string, int>();
        Dictionary<string, int> unions = new Dictionary<string, int>();
        Dictionary<string, int> interjection = new Dictionary<string, int>();
        Dictionary<string, int> participium = new Dictionary<string, int>();
        Dictionary<string, int> lastnames = new Dictionary<string, int>();
        Dictionary<string, int> firstnames = new Dictionary<string, int>();
        Dictionary<string, int> abbreviatura = new Dictionary<string, int>();
        Dictionary<string, int> reduction = new Dictionary<string, int>();

        Dictionary<string, int> unknownObjects = new Dictionary<string, int>();   // словаь неизвестных слов

        Dictionary<string, int> listKeyWordsProcess = new Dictionary<string, int>(); // словарь ключевых слов
        Dictionary<string, int> positive = new Dictionary<string, int>();
        Dictionary<string, int> negative = new Dictionary<string, int>();
        Dictionary<string, int> stopslovs = new Dictionary<string, int>();

        Dictionary<string, int> listWords = new Dictionary<string, int>();  // вспомогательный словарь для метода AddDictionaryToListView
        Dictionary<string, int> toshnota = new Dictionary<string, int>(); // вспомогательный словарь для вычисления тошноты
        Dictionary<string, int> listDictionary = new Dictionary<string, int>(); // вспомогательный для метода CheckCustom
        Dictionary<string, int> dictionary = new Dictionary<string, int>(); // вспомогательный для методов CheckOne и CheckDouble
        Dictionary<string, int> t; // выходной в метиоду тошноты

        Dictionary<string, int> wordsAll = new Dictionary<string, int>();
        Dictionary<string, int> semanticCore = new Dictionary<string, int>();
        Dictionary<string, int> lkpxml = new Dictionary<string, int>();
        Dictionary<string, int> ldxml = new Dictionary<string, int>();

        List<string> globalKey = new List<string>(); // глобальные ключи
        List<string> globalValue = new List<string>(); // глобальные значения
        List<string> globalKeyZnach = new List<string>(); // ключи значимых
        List<string> globalValueZnach = new List<string>(); // значения значимых

        List<string> globalCustomKey = new List<string>(); // ключи пользовательских словарей
        List<string> globalVodaKey = new List<string>(); // ключи воды
        List<string> globalValueStop = new List<string>(); // значения стоп-слов

        List<double> toshClass = new List<double>(); // тошнота классическая 
        List<double> toshAcad = new List<double>(); // тошнота академическая

        List<string> keyMassiveSpace = new List<string>(); // все ключи с пробелами
        List<string> kmpNoSpace = new List<string>(); // все ключи с замененным пробелами

        List<string> extendDictionary = new List<string>();
        List<string> keysUnknown = new List<string>();

        DirectoryInfo di = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
        Thread thr;
        DateTime begin;
        TimeSpan processingTime;

        StringBuilder strText = new StringBuilder(120);

        string process;
        string keywordsclipboard = String.Empty;
        string fileName = "Clipboard";  // for report
        public string snew = "";

        string[] processOut;
        string[] cleanedTested;
        string[] toProcess = null;

        char[] splitArray = new char[] { ' ', ',', '.', '!', '?', '\n', '\r', ';', '"', '}', '{', '[', ']', ':', '(', ')', '>', '<', '№', '$', '*', '—', '+', '=', '&', '^', '#', '@', '«', '»', '/', '\\' };
        char[] splitArrayMini = new char[] { '\n', '\r' };
        string[] splitArrayProbel = new string[] { "814270824437" };

        int cds;
        int cdi;
        int textcharlenght;
        int toProcessLenght = 0;
        int counterIteratorabbreviatura;
        int countSumAbbr;
        int counterIteratoradjectives;
        int countSumAdjectives;
        int counterIteratoradverbs;
        int countSumAdverbs;
        int counterIteratoredizmerenie;
        int countSumEdizmerenie;
        int countIterFirstNames;
        int countSumFirstNames;
        int counterIteratorhyperlinks;
        int counterSummatorhyperlinks;
        int counterIteratorinterjection;
        int counterSummatorinterjection;
        int counterIteratorlastnames;
        int counterSummatorlastnames;
        int counterIteratornouns;
        int counterSummatornouns;
        int counterIteratornumerals;
        int counterSummatornumerals;
        int counterIteratorparticipium;
        int counterSummatorparticipium;
        int counterIteratorparticles;
        int counterSummatorparticles;
        int counterIteratorpreposition;
        int counterSummatorpreposition;
        int counterIteratorpronounce;
        int counterSummatorpronounce;
        int counterIteratorunions;
        int counterSummatorunions;
        int counterIteratorunknownObjects;
        int counterSummatorunknownObjects;
        int counterIteratorverbaladverbs;
        int counterSummatorverbaladverbs;
        int counterIteratorverbs;
        int counterSummatorverbs;
        int counterIteratorwords;
        float counterSummatorwords;
        int counterIteratorlistKeyWordsProcess;
        int counterSummatorlistKeyWordsProcess;
        int counterIteratornegative;
        int counterSummatornegative;
        int counterIteratorpositive;
        int counterSummatorpositive;
        int ciStopslov;
        int csStopslov;


        // интерфейс, вызывается в методе  listView1_ColumnClick()
        // т.е. только работает в интерфейсе
        public IComparer SortViaValue { get { return (IComparer)new SortByValue(); } }
        public IComparer SortViaFrequency { get { return (IComparer)new SortByFrequency(); } }
        public IComparer SortViaPartOfSpeech { get { return (IComparer)new SortByPartOfSpeech(); } }

        internal static  bool ascendOrDescend = true;   // true  - sort by ascending

        bool rus = true; // for choice of language
        bool engl = false; // for choice of language 


        public Form1()
        {   
            InitializeComponent(); 
            //englishItem.Enabled = false;
            //complexItem.Enabled = false;
            //richTextBox1.Text = "Откройте в этом окне файл (воспользуйтесь верхним меню) или вставьте через буфер (кнопка справа).";
            //richTextBox3.Text = "Вставьте в это окно ключевые слова. Воспользуйтесь для этого кнопкой справа.";
            this.button1.Enabled = false;
            //this.button8.Enabled = false;

            //richTextBox2.Text = "";

            //columnHeader1.Width = listView1.Width / 3;
            //columnHeader2.Width = listView1.Width / 3;
            //columnHeader3.Width = listView1.Width / 3;

            //columnHeader6.Width = listView3.Width / 2;
            //columnHeader7.Width = listView3.Width / 2;

            //columnHeader10.Width = listView5.Width / 3;
            //columnHeader11.Width = listView5.Width / 3;
            //columnHeader8.Width = listView5.Width / 3;

            //columnHeader9.Width = listView4.Width / 3;
            //columnHeader12.Width = listView4.Width / 3;
            //columnHeader13.Width = listView4.Width / 3;

            //columnHeader14.Width = listView4.Width / 2;
            //columnHeader15.Width = listView4.Width / 2;

            //columnHeader16.Width = listView4.Width / 4;
            //columnHeader17.Width = listView4.Width / 4;
            //columnHeader18.Width = listView4.Width / 4;
            //columnHeader19.Width = listView4.Width / 4;

            //this.button2.Enabled = false;
            this.saveAsTextToolStripMenuItem.Enabled = false;
            this.saveAsXMLToolStripMenuItem.Enabled = false;
        }

        //выход из программы
        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        //загузка файла
        private void loadFileToolStripMenuItem_Click(object sender, EventArgs e)
        {          

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Текстовые файлы(*.txt)|*.txt";
            ofd.RestoreDirectory = true;
            DialogResult dr = ofd.ShowDialog();

            if (dr == DialogResult.OK)
            {
                fileName = ofd.FileName;
                FileStream fs = new FileStream(ofd.FileName, FileMode.Open);
                StreamReader sr = new StreamReader(fs, Encoding.GetEncoding(0));
                richTextBox1.Text = sr.ReadToEnd();
                sr.Close();
                fs.Close();
                this.button1.Enabled = true;
                //this.button8.Enabled = true;
            }
        }


        

        private void ProcessText(object sourceText)
        {
            string text = null;
            if (sourceText is string)
                text =(string)sourceText;
            begin = DateTime.Now;

            try
            {
                textcharlenght = text.ToCharArray().Length;
                // строим глоальный массив ключей Воды
                string[] AbbreviaturaKeys = Properties.Resources.AbbreviaturaKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] AdverbsKeys = Properties.Resources.AdverbsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] FirstNamesKeys = Properties.Resources.FirstNamesKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] InterjectionKeys = Properties.Resources.Interjection.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] LastNamesKeys = Properties.Resources.LastnamesKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] NumeralsKeys = Properties.Resources.NumeralsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] ParticlesKeys = Properties.Resources.ParticlesKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] PrepositionsKeys = Properties.Resources.PrepositionsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] PronounsKeys = Properties.Resources.PronounsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] ReductionKeys = Properties.Resources.Reduction.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] UnionsKeys = Properties.Resources.UnionsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();

                GlobalVodaKey(NumeralsKeys, InterjectionKeys, PronounsKeys, AdverbsKeys, PrepositionsKeys, UnionsKeys, ParticlesKeys, LastNamesKeys, FirstNamesKeys, AbbreviaturaKeys, ReductionKeys, "Строю глобальный массив ключей стоп-слов", out globalVodaKey);

                CollectSpace(globalVodaKey, out keyMassiveSpace, "CollectSpace");
                OperationSpace(keyMassiveSpace, "OperationSpace", out kmpNoSpace);
                ReplaseElement(text, keyMassiveSpace, kmpNoSpace, "ReplaseElement", out process);
                SplitArray (process, out processOut);
                ReplaseElementReverce(processOut, "ReplaseElementReverce", out toProcess);

                toProcessLenght = 0;
                for (int i = 0; i < toProcess.Length; i++)
                {
                    toProcessLenght = toProcessLenght + toProcess[i].ToString().ToCharArray().Length;
                }


                // исследую части речи
                CheckOne(toProcess, AbbreviaturaKeys, out dictionary, out cds, out cdi, "Проверяю аббревиатуры");
                AddDictionaryToListView(listView1, dictionary, "Аббревиатура");
                //PrintSum(cds, cdi, "Аббревиатура");
                counterIteratorabbreviatura = cdi;
                countSumAbbr = cds;
                abbreviatura = new Dictionary<string, int>(dictionary);

                string [] AdjectivesKeys = Properties.Resources.AdjectivesKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] AdjectivesValue = Properties.Resources.AdjectivesValue.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, AdjectivesKeys, AdjectivesValue, out dictionary, out cds, out cdi, "Проверяю прилагательные");
                AddDictionaryToListView(listView1, dictionary, "Прилагательное");
                //PrintSum(cds, cdi, "Прилагательное");
                counterIteratoradjectives = cdi;
                countSumAdjectives = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                adjectives = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, AdverbsKeys, out dictionary, out cds, out cdi, "Проверяю наречия");
                AddDictionaryToListView(listView1, dictionary, "Наречие");
                //PrintSum(cds, cdi, "Наречие");
                counterIteratoradverbs = cdi;
                countSumAdverbs = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                adverbs = new Dictionary<string, int>(dictionary);

                string [] FirstNamesValues = Properties.Resources.FirstNamesValues.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, FirstNamesKeys, FirstNamesValues, out dictionary, out cds, out cdi, "Проверяю имена");
                AddDictionaryToListView(listView1, dictionary, "Имя");
                //PrintSum(cds, cdi, "Имя");
                countIterFirstNames = cdi;
                countSumFirstNames = cds;
                firstnames = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, InterjectionKeys, out dictionary, out cds, out cdi, "Проверяю междометия");
                AddDictionaryToListView(listView1, dictionary, "Междометие");
                //PrintSum(cds, cdi, "Междометие");
                counterIteratorinterjection = cdi;
                counterSummatorinterjection = cds;
                interjection = new Dictionary<string, int>(dictionary);

                string [] LastNamesValues = Properties.Resources.LastnamesValue.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, LastNamesKeys, LastNamesValues, out dictionary, out cds, out cdi, "Проверяю фамилии");
                AddDictionaryToListView(listView1, dictionary, "Фамилия");
                //PrintSum(cds, cdi, "Фамилия");
                counterIteratorlastnames = cdi;
                counterSummatorlastnames = cds;
                lastnames = new Dictionary<string, int>(dictionary);

                string [] NounsKeys = Properties.Resources.NounsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] NounsValues = Properties.Resources.NounsValue.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, NounsKeys, NounsValues, out dictionary, out cds, out cdi, "Проверяю существительные");
                AddDictionaryToListView(listView1, dictionary, "Существительное");
                //PrintSum(cds, cdi, "Существительное");
                counterIteratornouns = cdi;
                counterSummatornouns = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                nouns = new Dictionary<string, int>(dictionary);

                string [] NumeralsValue = Properties.Resources.NumeralsValue.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, NumeralsKeys, NumeralsValue, out dictionary, out cds, out cdi, "Проверяю числительные");
                AddDictionaryToListView(listView1, dictionary, "Числительное");
                //PrintSum(cds, cdi, "Числительное");
                counterIteratornumerals = cdi;
                counterSummatornumerals = cds;
                numerals = new Dictionary<string, int>(dictionary);

                string [] ParticipiumKeys = Properties.Resources.Participium.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckOne(toProcess, ParticipiumKeys, out dictionary, out cds, out cdi, "Проверяю причастия");
                AddDictionaryToListView(listView1, dictionary, "Причастие");
                //PrintSum(cds, cdi, "Причастие");
                counterIteratorparticipium = cdi;
                counterSummatorparticipium = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                participium = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, ParticlesKeys, out dictionary, out cds, out cdi, "Проверяю частицы");
                AddDictionaryToListView(listView1, dictionary, "Частица");
                //PrintSum(cds, cdi, "Частица");
                counterIteratorparticles = cdi;
                counterSummatorparticles = cds;
                particles = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, PrepositionsKeys, out dictionary, out cds, out cdi, "Проверяю предлоги");
                AddDictionaryToListView(listView1, dictionary, "Предлог");
                //PrintSum(cds, cdi, "Предлог");
                counterIteratorpreposition = cdi;
                counterSummatorpreposition = cds;
                preposition = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, PronounsKeys, out dictionary, out cds, out cdi, "Проверяю местоимения");
                AddDictionaryToListView(listView1, dictionary, "Местоимение");
                //PrintSum(cds, cdi, "Местоимение");
                counterIteratorpronounce = cdi;
                counterSummatorpronounce = cds;
                pronounce = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, ReductionKeys, out dictionary, out cds, out cdi, "Проверяю сокращения");
                AddDictionaryToListView(listView1, dictionary, "Сокращение");
                //PrintSum(cds, cdi, "Сокращение");
                counterIteratoredizmerenie = cdi;
                countSumEdizmerenie = cds;
                reduction = new Dictionary<string, int>(dictionary);

                CheckOne(toProcess, UnionsKeys, out dictionary, out cds, out cdi, "Проверяю союзы");
                AddDictionaryToListView(listView1, dictionary, "Союз");
                //PrintSum(cds, cdi, "Союз");
                counterIteratorunions = cdi;
                counterSummatorunions = cds;
                unions = new Dictionary<string, int>(dictionary);

                string [] VerbalAdverbsKey = Properties.Resources.VerbalAdverbs.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string [] VerbalAdverbsValues = Properties.Resources.VerbalAdverbs.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, VerbalAdverbsKey, VerbalAdverbsValues, out dictionary, out cds, out cdi, "Проверяю дееперичастия");
                AddDictionaryToListView(listView1, dictionary, "Деепричастие");
                //PrintSum(cds, cdi, "Деепричастие");
                counterIteratorverbaladverbs = cdi;
                counterSummatorverbaladverbs = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                verbaladverbs = new Dictionary<string, int>(dictionary);

                string[] VerbsKeys = Properties.Resources.VerbsKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] VerbsValue = Properties.Resources.VerbsValue.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                CheckDouble(toProcess, VerbsKeys, VerbsValue, out dictionary, out cds, out cdi, "Проверяю глаголы");
                AddDictionaryToListView(listView1, dictionary, "Глагол");
                //PrintSum(cds, cdi, "Глагол");
                counterIteratorverbs = cds;
                counterSummatorverbs = cds;
                ToshnotaDictionary(dictionary, out toshnota);
                Toshnota(toshnota, out toshClass, out toshAcad);
                AddDictionaryToListView(listView7, toshnota, toshClass, toshAcad);
                verbs = new Dictionary<string, int>(dictionary);
                
                // после частей речи - проверка через регулярки
                //CheckHyperLinks(toProcess, "Проверяю гиперссылки");
                //AddDictionaryToListView(listView1, hyperlinks, "Гиперссылка");
                //counterIteratorhyperlinks = cdi;
                //counterSummatorhyperlinks = cds;
                //hyperlinks = new Dictionary<string, int>(dictionary);

                // чтобы проверить неизвестные - сначала строим Глоальный массив ключей
                GlobalKey(VerbsKeys, VerbalAdverbsKey, AdjectivesKeys, NounsKeys, NumeralsKeys, InterjectionKeys, PronounsKeys, AdverbsKeys, PrepositionsKeys, ParticipiumKeys, UnionsKeys, ParticlesKeys, LastNamesKeys, FirstNamesKeys, AbbreviaturaKeys, ReductionKeys, "Строю глобальный массив ключей");
                GlobalValue(VerbsValue, VerbalAdverbsValues, AdjectivesValue, NounsValues, NumeralsValue, InterjectionKeys, PronounsKeys, AdverbsKeys, PrepositionsKeys, ParticipiumKeys, UnionsKeys, ParticlesKeys, LastNamesValues, FirstNamesValues, AbbreviaturaKeys, ReductionKeys, "Строю глобальный массив значений");
                CheckUnknown(toProcess, globalKey, "Проверяю неизвестные");
                AddDictionaryToListView(listView1, unknownObjects, "Не определено");
                //counterIteratorunknownObjects = cdi;
                //counterSummatorunknownObjects = cds;
                unknownObjects = new Dictionary<string, int>(dictionary);

                foreach (var st in unknownObjects)
                {
                    snew += st.Key + ", ";
                }

                // прсто считаем слова и строим словарь
                CheckWords(toProcess, "Считаю слова", out words, out cds, out cdi);
                AddDictionaryToListView(listView3, words);
                counterIteratorwords = cdi;
                counterSummatorwords = cds;
                wordsAll = new Dictionary<string, int>(words);

                // строим глоальный массив ключей Значащих слов
                GlobalKeyZnach(VerbsKeys, VerbalAdverbsKey, AdjectivesKeys, NounsKeys, ParticipiumKeys, "Строю глобальный массив ключей значимых слов");
                //GlobalValueZnach(VerbsValue, VerbalAdverbsValues, AdjectivesValue, NounsValues, ParticipiumKeys, "Строю глобальный массив значений значимых слов");

                // строим глоальный массив Значений Воды
                GlobalVodalValue(NumeralsValue, InterjectionKeys, PronounsKeys, AdverbsKeys, PrepositionsKeys, UnionsKeys, ParticlesKeys, LastNamesValues, FirstNamesValues, AbbreviaturaKeys, ReductionKeys, "Строю глобальный массив значений стоп-слов");


                // построить globalCustomKey из кастомных словарей
                string[] keywordsclipboardArray = keywordsclipboard.ToLower().Trim(' ').Split(splitArray, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] StopslovKeys = Properties.Resources.StopSlovoKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] PositiveKeys = Properties.Resources.PositiveKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                string[] NegativeKeys = Properties.Resources.NegativeKeys.ToLower().Trim(' ').Split(splitArrayMini, StringSplitOptions.RemoveEmptyEntries).ToArray();
                GlobalCustomKey(keywordsclipboardArray, StopslovKeys, PositiveKeys, NegativeKeys, out globalCustomKey);

                // вычленяем из предыдущего массива массив с пробелами
                CollectSpace(globalCustomKey, out keyMassiveSpace, "CollectSpace");
                OperationSpace(keyMassiveSpace, "OperationSpace", out kmpNoSpace);
                ReplaseElement(text, keyMassiveSpace, kmpNoSpace, "ReplaseElement", out process);
                SplitArray(process, out processOut);
                ReplaseElementReverce(processOut, "ReplaseElementReverce", out toProcess);

                // очищаем тестируемый массив от Воды
                CleanerTested(toProcess, globalVodaKey, "Очищаем тестируемый массив", out cleanedTested);

                // строим семантическое ядро
                CheckWords(cleanedTested, "Строю семантическое ядро", out words, out cds, out cdi);
                AddDictionaryToListView(listView6, words, "Ядро");
                semanticCore = new Dictionary<string, int>(words);

                // стргий поиск ключевых слов
                CheckKeyWords(cleanedTested, keywordsclipboardArray, "Проверяю ключевые слова");
                AddDictionaryToListView(listView4, listKeyWordsProcess, "Ключевое слово");
                //counterIteratorlistKeyWordsProcess = cdi;
                //counterSummatorlistKeyWordsProcess = cds;
                lkpxml = new Dictionary<string, int>(listKeyWordsProcess);

                // нестрогий поиск ключевых слов
                GenerateExtendDictionary(keywordsclipboardArray, globalKey, globalValue, "Веду нестрогий поиск ключевиков", out extendDictionary);
                CheckCustom(cleanedTested, extendDictionary, "Нестрогий поиск ключевиков", out listDictionary);
                AddDictionaryToListView(listView4, listDictionary, "Ключевик, нестрогий");
                ldxml = new Dictionary<string, int>(listDictionary);

                // посик стоп-слов
                CheckOne(cleanedTested, StopslovKeys, out dictionary, out cds, out cdi, "Проверяю стоп-слова");
                AddDictionaryToListView(listView5, dictionary, "Стоп-слово");
                //PrintSum(cds, cdi, "Стоп-слово");
                ciStopslov = cdi;
                csStopslov = cds;
                stopslovs = new Dictionary<string, int>(dictionary);

                // посик на позитив
                CheckOne(cleanedTested, PositiveKeys, out dictionary, out cds, out cdi, "Проверяю на позитив");
                AddDictionaryToListView(listView5, dictionary, "Позитив");
                //PrintSum(cds, cdi, "Позитив");
                counterIteratorpositive = cdi;
                counterSummatorpositive = cds;
                positive = new Dictionary<string, int>(dictionary);

                // поиск на негатив
                CheckOne(cleanedTested, NegativeKeys, out dictionary, out cds, out cdi, "Проверяю на негатив");
                AddDictionaryToListView(listView5, dictionary, "Негатив");
                //PrintSum(cds, cdi, "Негатив");
                counterIteratornegative = cdi;
                counterSummatornegative = cds;
                negative = new Dictionary<string, int>(dictionary);

                processingTime = DateTime.Now - begin;

                PrintResume();



                this.button1.Invoke(new Action(() =>
                {
                    this.label3.Text = string.Empty;
                    this.progressBar1.Visible = false;
                    loadFileToolStripMenuItem.Enabled = true;
                    this.button1.Enabled = true;
                    //this.button2.Enabled = true;
                    //this.button8.Enabled = true;
                    saveAsTextToolStripMenuItem.Enabled = true;
                    saveAsXMLToolStripMenuItem.Enabled = true;
                    listView1.Show();
                    this.progressBar1.Value = 0;
                }));

                MessageBox.Show("Готово", "Сообщение", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            catch (Exception ex)
            {
                 MessageBox.Show(ex.ToString());
                 this.button1.Invoke(new Action(() =>
                 {
                     this.loadFileToolStripMenuItem.Enabled = true;
                     this.button1.Enabled = true;
                     this.progressBar1.Visible = false;
                     this.progressBar1.Value = 0;
                       this.label3.Text = "";
                 }));
            }
        }


        //private void russianItem_Click_1(object sender, EventArgs e)
        //{
        //    if (russianItem.CheckState == CheckState.Unchecked || russianItem.CheckState == CheckState.Indeterminate)
        //    {
        //        //complexItem.Checked = false;
        //        //englishItem.Checked = false;
        //        //russianItem.Checked = true;
        //        rus = true;
        //        engl = false;
        //    }

        //     else
        //    {
        //        russianItem.Checked = false;
        //        rus = false;
        //    }
             
        //}

        //private void englishItem_Click(object sender, EventArgs e)
        //{
        //    if (englishItem.CheckState == CheckState.Unchecked || englishItem.CheckState == CheckState.Indeterminate)
        //    {
        //        englishItem.CheckState = CheckState.Checked;
        //        russianItem.Checked = false;
        //        complexItem.Checked = false;
        //        rus = false;
        //        engl = true;
        //    }

        //    else
        //    {
        //        englishItem.Checked = false;
        //        engl = false;
        //        russianItem.Checked = true;
        //        rus = true;
        //    }
        //}

        //private void complexItem_Click(object sender, EventArgs e)
        //{
        //    if (complexItem.CheckState == CheckState.Indeterminate || complexItem.CheckState == CheckState.Unchecked)
        //    {
        //        complexItem.Checked = true;
        //        engl = true;
        //        rus = true;
        //        englishItem.Checked = false;
        //        russianItem.Checked = false;
        //    }
        //    else
        //    {
        //        englishItem.Checked = false;
        //        complexItem.Checked = false;
        //        russianItem.Checked = true;
        //        rus = true;
        //    }
        //}

        private void Form1_Load(object sender, EventArgs e)
        {
            label3.Text = "Процесс...";
            progressBar1.Visible = true;
            progressBar1.Value = 1;
            progressBar1.Maximum = 1000;
        }

        /**   мои методы */ 

        // считаю слова
        private void CheckWords(string[] strings, string label, out Dictionary<string, int> words, out int cds, out int cdi)
        {
            //сообщение в прогрессбар
            this.progressBar1.Invoke(new Action(() => { this.label3.Text = "Считаю слова..."; }));

            words = new Dictionary<string, int>();
            string[] inProcess = strings;
            var groups = inProcess
                .Select((Name, Index) => new { Name, Index })
                .GroupBy(p => p.Name)
                .Select(s => new { Name = s.Key, Count = s.Count() });

            foreach (var group in groups)
            {
                string name = group.Name;
                int count = group.Count;
                words.Add(name, count);
            }
            
            cdi = words.Count;
            cds = words.Sum(y => y.Value); 
            
            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; })); // прогрессбар
        }

        // составляю семантическое ядро текста
         
        // очистить тестируемый от Воды
        private void CleanerTested(string[] strings, List<string>globalVodaKey, string label, out string[] cleanedTested)
        {
            List<string> tested = new List<string>(strings);
            List<string> gvk = new List<string>(globalVodaKey);
            cleanedTested = tested.Except(globalVodaKey).ToArray();
        }

        // проверяю массивы с одним словарем
        private void CheckOne(string[] text, string [] keys, out Dictionary<string, int> dictionary, out int cds, out int cdi, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            string[] tested = text;
            string[] key = keys;

            dictionary = new Dictionary<string, int>();
            cds = 0;
            cdi = 0;

            var d = from a1 in tested  // ключи массива   /** берем переменную в массиве1 */
                    let a2key = Array.IndexOf(key, a1)
                    let a3 = a2key == -1 ? String.Empty : key[a2key]
                    group a3 by a3 into cnt  /** группировать в какое-то cnt */
                    where cnt.Key != null
                    where cnt.Key != String.Empty/** где ключ неравен null       */
                    where cnt.Key != " "
                    where cnt.Count() > 0    /** где ключ неравен null       */
                    select new { key = (string)cnt.Key, qty = cnt.Count() }; /** сформировтаь массив */

            foreach (var y in d)
            {
                dictionary.Add(y.key, y.qty);
            }

            cdi = dictionary.Count;
            cds = dictionary.Sum(y => y.Value); 

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }

        //private void PrintSum(int cds, int cdi, string label)
        //{
        //    control.Append("Сумма слов: " + label + ": " + cds + "\n");
        //    control.Append("Уникальных cлов: " + label + ": " + cdi + "\n");
        //}

        // проверю массивы с двумя словарями
        private void CheckDouble(string[] text, string [] keys, string [] values, out Dictionary<string, int> dictionary, out int cds, out int cdi, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            string[] tested = text;
            string[] key = keys;
            string[] value = values;
            dictionary = new Dictionary<string, int>();
            cds = 0;
            cdi = 0;

            var d = from a1 in tested  // ключи массива   /** берем переменную в массиве1 */
                       let a2key = Array.IndexOf(key, a1)
                       let a3 = a2key == -1 ? String.Empty : value[a2key]
                       group a3 by a3 into cnt  /** группировать в какое-то cnt */
                       where cnt.Key != null
                       where cnt.Key != String.Empty/** где ключ неравен null       */
                       where cnt.Key != " "
                       where cnt.Count() > 0    /** где ключ неравен null       */
                       select new { key = (string)cnt.Key, qty = cnt.Count() }; /** сформировтаь массив */

            foreach (var y in d)
            {
                dictionary.Add(y.key, y.qty);
            }

            cdi = dictionary.Count;
            cds = dictionary.Sum(y => y.Value); 

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; })); // прогрессбар
        }
       
        // проверяю гиперссылки и другие - проверка на регулярные выражения
        private void CheckHyperLinks(string[] strings, string label)
        {
            //knownList.Add(y.globalKey);

            //this.progressBar1.Invoke(new Action(() => { this.label3.Text = "Searching for hyperlinks...";
            //          }));
            //timerHyperLinks = Stopwatch.StartNew();

            //            counterHyperlink++;
            //регулярка под номера телефона  

            //(^\+\d{1,2})?((\(\d{3}\))|(\-?\d{3}\-)|(\d{3}))((\d{3}\-\d{4})|(\d{3}\-\d\d\  
            //-\d\d)|(\d{7})|(\d{3}\-\d\-\d{3}))  
            //             * 
            //             * http://skillcoding.com/Default.aspx?id=203
            //             * 
            //             * (http|https):\/\/[\w\-_]+(\.[\w\-_]+)+([\w\-\.,@?^=%&amp;:/~\+#]*[\w\-\@?^=%&amp;/~\+#])?
            //             * \b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,6}\b
            //             * 
            //             * 
            //             * String e_mail = "name@gmail.";  

            //public bool isValidMail(string e_mail)  
            //{  
            //   string expr =   
            //     "[.\\-_a-z0-9]+@([a-z0-9][\\-a-z0-9]+\\.)+[a-z]{2,6}";  

            //   Match isMatch =   
            //     Regex.Match(e_mail, expr, RegexOptions.IgnoreCase);  

            //   return isMatch.Success;  
            //}

            //timerHyperLinks.Stop();

        }
     
        // проверка неизвестных
        private void CheckUnknown(string[] text, List<string> k, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            string[] tested = text;

            List<string> localkey = new List<string>(k);

            keysUnknown = tested.Where(t => !localkey.Contains(t)).Distinct().ToList<string>();

            foreach (var y in keysUnknown)
            {
                unknownObjects.Add(y, 1);
                counterSummatorunknownObjects = counterSummatorunknownObjects + 1;
            }

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }
        
        // генерируем расширенный словарь
        private void GenerateExtendDictionary(string [] keys, List<string> inkeys, List<string> invalues, string label, out List<string> extendDictionary)
        {
            //этот метод формирует расширенный словарь, на базе "кастомных" словарей:
            //негатив, позитив, стоп-слова, ключевые слова.
            //проверка по этим кастомным словарям проводится в 2 этапа - 
            //второй этап - это метод, CheckCustom
            //и распечатать

            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            List<string> key = new List<string>(keys); // ключевики            
            List<string> newinkeys = new List<string>(inkeys).ToList(); // глобальный кей
            List<string> newinvalues = new List<string>(invalues).ToList(); // глобальный вэлью
            extendDictionary = new List<string>();

            try
            {
                var d = from a1 in key  // взять элемент из массива ключевиков
                        let a2key = newinkeys.IndexOf(a1) //   найти их в ГМК
                                                          //   дать переменной значение индекса
                        let a3key = a2key == -1 ? String.Empty : newinvalues[a2key] // присвоить переменной значение элемента
                        group a3key by a3key into cnt  /** группировать в cnt */
                        where cnt.Key != null
                        where cnt.Key != String.Empty/** где ключ неравен null       */
                        where cnt.Key != " "
                        where cnt.Count() > 0    /** где ключ неравен null       */
                        select new { key = (string)cnt.Key }; /** сформировтаь массив  , qty = cnt.Count()   */
                
                // формируем Список элементов, взятых из ГМЗ
                List<string> ld = new List<string>();
                foreach (var y in d)
                {
                    ld.Add(y.key);
                }

                /** найти индексы всех равных элементам dict в newinvalues элементов.
                 * для итерируемого в newinvalues
                 * если равен итерируемому в dict
                 * в искомый массив: элеменг равен newinkyes[i]
                 */

                for (int i = 0; i < newinvalues.Count; i++)
                {
                    for (int j = 0; j < ld.Count; j++)
                    {
                        if (newinvalues[i].Trim(' ').ToLower() == ld[j].Trim(' ').ToLower())
                        {
                            extendDictionary.Add(newinkeys[i]);
                        }
                    }
                }
                ld.Clear();
            }

            catch (Exception ex)
            {
                MessageBox.Show(label + " поиск не проведен", label + " поиск не проведен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            newinkeys.Clear(); newinvalues.Clear();
            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; })); // прогрессбар
        }
        
        // проверяем кастомный словарь (негатив, позитив, стопы, ключевики)
        private void CheckCustom(string[] text, List<string> extendDictionary, string label, out Dictionary<string, int> listDictionary)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            /** узнать, какие и сколько элементов эталонного списка
             * содержатс в тестируемом */

            List<string> key = new List<string>(extendDictionary).ToList(); // эталонный список  
            List<string> tested = new List<string>(text).ToList(); // тестируемый текст
            listDictionary = new Dictionary<string, int>();

            try
            {
                var d = from a1 in key  // взять элемент из массива эталонного списка
                        let a2key = tested.IndexOf(a1) //   найти их в тестируемом массиве и дать переменной значение индекса
                        let a3key = a2key == -1 ? String.Empty : tested[a2key] // присвоить переменной значение элемента
                        group a3key by a3key into cnt  /** группировать в cnt */
                        where cnt.Key != null
                        where cnt.Key != String.Empty/** где ключ неравен null       */
                        where cnt.Key != " "
                        where cnt.Count() > 0    /** где ключ неравен null       */
                        select new { key = (string)cnt.Key, qty = cnt.Count() }; /** сформировтаь массив    */

                // сделан результирующий Диктионари                
                foreach (var pair in d)
                {
                    listDictionary.Add(pair.key, pair.qty);
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(label + " поиск не проведен", label + " поиск не проведен", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; })); // прогрессбар
        }

        // строгая проверка на ключевики
        private void CheckKeyWords(string[] text, string [] keys, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            string[] tested = text;
            string[] key = keys;

            var dict = from a1 in tested  // ключи массива   /** берем переменную в массиве1 */
                       let a2key = Array.IndexOf(key, a1)
                       let a3 = a2key == -1 ? String.Empty : key[a2key]
                       group a3 by a3 into cnt  /** группировать в какое-то cnt */
                       where cnt.Key != null
                       where cnt.Key != String.Empty/** где ключ неравен null       */
                       where cnt.Key != " "
                       where cnt.Count() > 0    /** где ключ неравен null       */
                       select new { key = (string)cnt.Key, qty = cnt.Count() }; /** сформировтаь массив */

            foreach (var y in dict)
            {
                listKeyWordsProcess.Add(y.key, y.qty);
                counterSummatorlistKeyWordsProcess = counterSummatorlistKeyWordsProcess + y.qty;
            }

            counterIteratorlistKeyWordsProcess = listKeyWordsProcess.Count;
            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; })); // прогрессбар
        }

        
        // глобальный Кей
        private void GlobalKey(string [] keys1, string [] keys2, string [] keys3, string [] keys4, string [] keys5, string [] keys6, string [] keys7, string [] keys8, string [] keys9, string [] keys10, string [] keys11, string [] keys12, string [] keys13, string [] keys14, string [] keys15, string [] keys16, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            globalKey.AddRange(keys1);
            globalKey.AddRange(keys2);
            globalKey.AddRange(keys3);
            globalKey.AddRange(keys4);
            globalKey.AddRange(keys5);
            globalKey.AddRange(keys6);
            globalKey.AddRange(keys7);
            globalKey.AddRange(keys8);
            globalKey.AddRange(keys9);
            globalKey.AddRange(keys10);
            globalKey.AddRange(keys11);
            globalKey.AddRange(keys12);
            globalKey.AddRange(keys13);
            globalKey.AddRange(keys14);
            globalKey.AddRange(keys15);
            globalKey.AddRange(keys16);

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }
        // глобальный вэлью
        private void GlobalValue(string [] values1, string [] values2, string [] values3, string [] values4, string [] values5, string [] values6, string [] values7, string [] values8, string [] values9, string [] values10, string [] values11, string [] values12, string [] values13, string [] values14, string [] values15, string [] values16, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            globalValue.AddRange(values1);
            globalValue.AddRange(values2);
            globalValue.AddRange(values3);
            globalValue.AddRange(values4);
            globalValue.AddRange(values5);
            globalValue.AddRange(values6);
            globalValue.AddRange(values7);
            globalValue.AddRange(values8);
            globalValue.AddRange(values9);
            globalValue.AddRange(values10);
            globalValue.AddRange(values11);
            globalValue.AddRange(values12);
            globalValue.AddRange(values13);
            globalValue.AddRange(values14);
            globalValue.AddRange(values15);
            globalValue.AddRange(values16);

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }

        // глобальный Кей значимых слов. Нужен для определеняи массива Ядра
        private void GlobalKeyZnach(string [] keys1, string [] keys2, string [] keys3, string [] keys4, string [] keys5, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            globalKeyZnach.AddRange(keys1);
            globalKeyZnach.AddRange(keys2);
            globalKeyZnach.AddRange(keys3);
            globalKeyZnach.AddRange(keys4);
            globalKeyZnach.AddRange(keys5);

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }

        // глобальные массивы Воды
        private void GlobalVodaKey(string [] keys1, string [] keys2, string [] keys3, string [] keys4, string [] keys5, string [] keys6, string [] keys7, string [] keys8, string [] keys9, string [] keys10, string [] keys11, string label, out List<string> globalVodaKey)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            globalVodaKey = new List<string>();

            globalVodaKey.AddRange(keys1);
            globalVodaKey.AddRange(keys2);
            globalVodaKey.AddRange(keys3);
            globalVodaKey.AddRange(keys4);
            globalVodaKey.AddRange(keys5);
            globalVodaKey.AddRange(keys6);
            globalVodaKey.AddRange(keys7);
            globalVodaKey.AddRange(keys8);
            globalVodaKey.AddRange(keys9);
            globalVodaKey.AddRange(keys10);
            globalVodaKey.AddRange(keys11);

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }
        private void GlobalVodalValue(string [] values1, string [] values2, string [] values3, string [] values4, string [] values5, string [] values6, string [] values7, string [] values8, string [] values9, string [] values10, string [] values11, string label)
        {
            this.label3.Invoke(new Action(() => { this.label3.Text = label; }));

            globalValueStop.AddRange(values1);
            globalValueStop.AddRange(values2);
            globalValueStop.AddRange(values3);
            globalValueStop.AddRange(values4);
            globalValueStop.AddRange(values5);
            globalValueStop.AddRange(values6);
            globalValueStop.AddRange(values7);
            globalValueStop.AddRange(values8);
            globalValueStop.AddRange(values9);
            globalValueStop.AddRange(values10);
            globalValueStop.AddRange(values11);

            this.label3.Invoke(new Action(() => { this.progressBar1.Value += 20; }));
        }


        private void GlobalCustomKey(string[] keywordsclipboardArray, string[] StopslovKeys, string[] PositiveKeys, string[] NegativeKeys, out List<string> globalCustomKey)
        {
            globalCustomKey = new List<string>();
            globalCustomKey.AddRange(keywordsclipboardArray);
            globalCustomKey.AddRange(StopslovKeys);
            globalCustomKey.AddRange(PositiveKeys);
            globalCustomKey.AddRange(NegativeKeys);
        }

        // работа над словами с пробелом
        // строит словарь пробельных. 
        // если основная проверка - брать в Кеях частей речи
        // если кастомная - в кеях касомных словарей
        // т.е. 2жды метод импользуется, с разным набором
        private void CollectSpace(List<string> keyMassive, out List<string> keyMassiveSpace, string label)
        {
            List<string> keyMassiveProcess = new List<string>(keyMassive);
            
            var dictionary = from a1 in keyMassiveProcess
                             where a1.Contains(" ")
                             select a1;

            keyMassiveSpace = new List<string>(dictionary); 
        }

        // заменяет в СП пробелы знаком \0
        private void OperationSpace(List<string> keyMassiveSpace, string label, out List<string> kmpNoSpace)
        {
            List<string> kmp = new List<string>(keyMassiveSpace);
            kmpNoSpace = new List<string>();
            kmpNoSpace.AddRange(kmp);

            string pattern = (" ");
            string replace = ("\0");

            for (int i = 0; i < kmpNoSpace.Count; i++) 
            { 
                kmpNoSpace[i] = kmpNoSpace[i].Replace(pattern, replace);
            }
        }

        // заменяет вхождения в тестируемом массиве
        private void ReplaseElement(string strings, List<string> p, List<string> r, string label, out string process)
        {
            /** в прараметрах - текстовая строка, Список без проблемов, Список с пробелмами */
            process = strings;
            List<string> pattern = new List<string>();
            pattern.AddRange(p);
            List<string> replace = new List<string>();
            replace.AddRange(r);

            for (int i = 0; i < pattern.Count; i++)
            {
                process = process.Replace(pattern[i], replace[i]);
            }
        }

        private void SplitArray (string process, out string [] processOut)
            {
                processOut = process.ToLower().Trim(' ').Split(splitArray, StringSplitOptions.RemoveEmptyEntries).ToArray();
            }

        // реверс замены
        private void ReplaseElementReverce(string[] strings, string label, out string[] toProcess)
        {
            toProcess = strings;

            for (int i = 0; i < toProcess.Length; i++)
            {
                toProcess[i] = toProcess[i].Replace("\0", " ");
            }
        }

        private void ToshnotaDictionary(Dictionary<string, int> d, out Dictionary<string, int> t)
        {
            t = new Dictionary<string, int>(); // выходной в метиоду тошноты

            foreach (var pair in d)
            {
                t.Add(pair.Key, pair.Value);
            }
        }

        private void Toshnota(Dictionary<string, int> t, out List<double> tc, out List<double> ta)
        {
            //toshClass.Clear();
            //toshAcad.Clear();
            tc = new List<double>();
            ta = new List<double>();

            double sp0 = toProcess.Length;
            double sp = sp0 / 100;
            
            foreach (var pair in t)
            {
                tc.Add(Math.Round(Math.Sqrt(pair.Value), 2));
                ta.Add(pair.Value / sp);
            }
        }

        private void Toshnota(Dictionary<string, int> d)
        {
            toshClass.Clear();
            toshAcad.Clear();
            Dictionary<string, int> d2 = new Dictionary<string, int>(d);
            double sp0 = d2.Sum(y => y.Value);
            double sp = sp0 / 100;

            foreach (var pair in d2)
            {
                toshClass.Add(Math.Round(Math.Sqrt(pair.Value), 2));
                toshAcad.Add(pair.Value / sp);
            }
        }

        private void PrintResume()
        {
            //Mailer.MailSend("проверка прошла", snew);
            strText.Clear();
            
            // количество Неуникальных значимых слов

            strText.Append("Время выполнения: " + processingTime.Hours + " часов " + processingTime.Minutes + " минут " + processingTime.Seconds + " секунд");

            strText.Append("\nКоличество знаков в тексте до удаления пробелов и знаков препинания: " + textcharlenght);
            strText.Append("\nКоличество знаков в тексте после удаления знаков препинания: " + toProcessLenght);

            // общее кол-во слов
            int sw = toProcess.Length;
            strText.Append("\nКоличество слов в тексте: " + sw);

            // кол-во значимых слов
            int ssum = countSumAdjectives + countSumAdverbs + counterSummatornouns + counterSummatorparticipium + counterSummatorverbaladverbs + counterSummatorverbs;
            strText.Append("\nКоличество значимых слов: " + ssum);
            strText.Append("\nCумма всех существительных, глаголов, прилагательных, наречечий, причастий и деепричастий");

            // кол-во слов-воды
            int summavoda = sw - ssum;
            strText.Append("\nКоличество слов-воды: "+ summavoda);
            strText.Append("\nCумма всех слов, кроме существительных, глаголов, прилагательных, наречечий, причастий и деепричастий, и слов для которых часть речи не была определена");
            float water = (summavoda / (counterSummatorwords / 100)); // вода доля
            strText.Append("\nДоля \"воды\": " + water.ToString("0.00") + " %");

            strText.Append("\n\nСлов частей речи всего: ");
            strText.Append("\nАббревиатур:      " + countSumAbbr + ", Доля " + (countSumAbbr / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nПрилагательных:   " + countSumAdjectives + ", Доля " + (countSumAdjectives / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nНаречий:          " + countSumAdverbs + ", Доля " + (countSumAdverbs / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nЕдиниц измерения: " + countSumEdizmerenie + ", Доля " + (countSumEdizmerenie / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nИмен:             " + countSumFirstNames + ", Доля " + (countSumFirstNames / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nГиперссылок:      " + counterSummatorhyperlinks + ", Доля " + (counterSummatorhyperlinks / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nМеждометий:       " + counterSummatorinterjection + ", Доля " + (counterSummatorinterjection / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nФамилий:          " + counterSummatorlastnames + ", Доля " + (counterSummatorlastnames / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nСуществительных:  " + counterSummatornouns + ", Доля " + (counterSummatornouns / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nЧислительных:     " + counterSummatornumerals + ", Доля " + (counterSummatornumerals / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nПричастий:        " + counterSummatorparticipium + ", Доля " + (counterSummatorparticipium / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nЧастиц:           " + counterSummatorparticles + ", Доля " + (counterSummatorparticles / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nПредлогов:        " + counterSummatorpreposition + ", Доля " + (counterSummatorpreposition / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nМестоимений:      " + counterSummatorpronounce + ", Доля " + (counterSummatorpronounce / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nСоюзов:           " + counterSummatorunions + ", Доля " + (counterSummatorunions / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nДеепричастий:     " + counterSummatorverbaladverbs + ", Доля " + (counterSummatorverbaladverbs / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nГлаголов:         " + counterSummatorverbs + ", Доля " + (counterSummatorverbs / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nНеопределено:     " + counterSummatorunknownObjects + ", Доля " + (counterSummatorunknownObjects / (counterSummatorwords / 100)).ToString("0.00") + " %");


          strText.Append("\n\nКлючевых слов всего: " + counterSummatorlistKeyWordsProcess + ", Доля " + (counterSummatorlistKeyWordsProcess / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nНегативных  всего: " + counterSummatornegative + ", Доля " + (counterSummatornegative / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nПозитивных  всего: " + counterSummatorpositive + ", Доля " + (counterSummatorpositive / (counterSummatorwords / 100)).ToString("0.00") + " %");
            strText.Append("\nСтоп-слов  всего: " + csStopslov + ", Доля " + (csStopslov / (counterSummatorwords / 100)).ToString("0.00") + " %" + Environment.NewLine);
                        

            //richTextBox2.Text += strText.ToString();
            if (richTextBox2.InvokeRequired)
                richTextBox2.Invoke(new MethodInvoker(() => { richTextBox2.Text += strText.ToString(); }));
            else
                richTextBox2.Text += strText.ToString();
        }

 


        /**   */

        /** кнопка Старт. сюда добавить всякие очистки? */
        private void button1_Click(object sender, EventArgs e)
        {
            ClearDictionary();
            ClearInterface();

            loadFileToolStripMenuItem.Enabled = false;
            ParameterizedThreadStart pst = new ParameterizedThreadStart(ProcessText);
            thr = new Thread(pst);
            thr.Priority = ThreadPriority.Highest;
            Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            thr.Start(richTextBox1.Text);  // тут содержимое Текстбокса передается в поток
            thr.IsBackground = true; // закрывает поток при закрытии программы
            keywordsclipboard = richTextBox3.Text;
            this.button1.Enabled = false;           
            this.progressBar1.Visible = true;
            this.label3.Visible = true;
        }

        // save statistics button
        private void button2_Click(object sender, EventArgs e)
        {
            saveAsTextToolStripMenuItem.PerformClick();
        }
        

        // add dictionary to listview метод
        private void AddDictionaryToListView(ListView listview, Dictionary<string, int> listWords)
        {
            ListViewItem[] lvItems = new ListViewItem[listWords.Count];

            int i = 0;

            foreach (KeyValuePair<string, int> kvp in listWords)
            {
                lvItems[i] = new ListViewItem(kvp.Key);
                lvItems[i].SubItems.Add(kvp.Value.ToString());                
                i++;
            }

            listview.Invoke(new Action(() => { listview.Items.AddRange(lvItems); }));
        }
        
        private void AddDictionaryToListView(ListView listview, Dictionary<string, int> partOfSpeech, string NameOfPartOfSpeech)
        {
            ListViewItem[] lvItems = new ListViewItem[partOfSpeech.Count];

            int i = 0;
            foreach (KeyValuePair<string, int> kvp in partOfSpeech)
            {
                lvItems[i] = new ListViewItem(kvp.Key);
                lvItems[i].SubItems.Add(kvp.Value.ToString());
                lvItems[i].SubItems.Add(NameOfPartOfSpeech);
                i++;
            }
            listview.Invoke(new Action(() => { listview.Items.AddRange(lvItems); }));
        }

        private void AddDictionaryToListView(ListView listview, Dictionary<string, int> partOfSpeech, List<double> sqrt, List<double> percent)
        {
            ListViewItem[] lvItems = new ListViewItem[partOfSpeech.Count];

            int i = 0;
            foreach (KeyValuePair<string, int> kvp in partOfSpeech)
            {
                lvItems[i] = new ListViewItem(kvp.Key);
                lvItems[i].SubItems.Add(kvp.Value.ToString());
                lvItems[i].SubItems.Add((sqrt[i]).ToString());
                lvItems[i].SubItems.Add((percent[i]).ToString("0.00"));
                i++;
            }


            listview.Invoke(new Action(() => { listview.Items.AddRange(lvItems); }));
        }
        

        // Сортировка в Listview 
        private void listView1_ColumnClick(object sender, ColumnClickEventArgs e)
       {

           if (listView1.Items.Count != 0)
           {
               if (e.Column == 0)
                   listView1.ListViewItemSorter = this.SortViaValue;
               else if (e.Column == 1)
                   listView1.ListViewItemSorter = this.SortViaFrequency;
               else 
                   listView1.ListViewItemSorter = this.SortViaPartOfSpeech;
               listView1.Sort();
               Form1.ascendOrDescend = !Form1.ascendOrDescend;
           }
       }

        private void listView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView2.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView2.ListViewItemSorter = this.SortViaValue;
                else if (e.Column == 1)
                    listView2.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView2.ListViewItemSorter = this.SortViaPartOfSpeech;
                listView2.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }

        private void listView3_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView3.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView3.ListViewItemSorter = this.SortViaValue;
                else if (e.Column == 1)
                    listView3.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView3.ListViewItemSorter = this.SortViaPartOfSpeech;
                listView3.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }

        private void listView4_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView4.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView4.ListViewItemSorter = this.SortViaValue;
                else if (e.Column == 1)
                    listView4.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView4.ListViewItemSorter = this.SortViaPartOfSpeech;
                listView4.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }

        private void listView5_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView1.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView5.ListViewItemSorter = this.SortViaValue;
                else if (e.Column == 1)
                    listView5.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView5.ListViewItemSorter = this.SortViaPartOfSpeech;
                listView5.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }

        private void listView6_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView6.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView6.ListViewItemSorter = this.SortViaValue;
                else if (e.Column == 1)
                    listView6.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView6.ListViewItemSorter = this.SortViaPartOfSpeech;
                listView6.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }

        private void listView7_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            if (listView7.Items.Count != 0)
            {
                if (e.Column == 0)
                    listView7.ListViewItemSorter = this.SortViaValue;
                //else if (e.Column == 1)
                //    listView7.ListViewItemSorter = this.SortViaFrequency;
                else
                    listView7.ListViewItemSorter = this.SortViaFrequency;
                listView7.Sort();
                Form1.ascendOrDescend = !Form1.ascendOrDescend;
            }
        }


        public class SortByValue : IComparer
       {
          public SortByValue() { }

           int IComparer.Compare(object one, object two)
           {
               if (Form1.ascendOrDescend)
                   return string.Compare(((ListViewItem)one).SubItems[0].Text, ((ListViewItem)two).SubItems[0].Text);
               else return -string.Compare(((ListViewItem)one).SubItems[0].Text, ((ListViewItem)two).SubItems[0].Text);
           }

       }
        public class SortByFrequency : IComparer
       {
           int first = 0;
           int second = 0;

           public SortByFrequency() { } 
  
        int IComparer.Compare(object one, object two)
           {

               first = Convert.ToInt32(((ListViewItem)one).SubItems[1].Text);
               second = Convert.ToInt32(((ListViewItem)two).SubItems[1].Text);
              
              if (true)
              {
		 
               if (Form1.ascendOrDescend)
               {
                   if (first > second)
                       return 1;
                   else if (first < second)
                       return -1;
                   else return 0;
               }
               else
               {
                   if (first > second)
                       return -1;
                   else if (first < second)
                       return 1;
                   else return 0;
               } 
	}

           }
       }    
        public class SortByPartOfSpeech : IComparer
           {
              public SortByPartOfSpeech() { }

              int IComparer.Compare(object one, object two)
               {
                    if (Form1.ascendOrDescend)
                       return string.Compare(((ListViewItem)one).SubItems[2].Text, ((ListViewItem)two).SubItems[2].Text);
                   else return -string.Compare(((ListViewItem)one).SubItems[2].Text, ((ListViewItem)two).SubItems[2].Text);
               }

           }

        // методы очистки
        private void ClearVarArr()
        {
            
        }

        private void ClearInterface()
        {
            listView1.Items.Clear();
            listView2.Items.Clear();
            listView3.Items.Clear();
            listView4.Items.Clear();
            listView5.Items.Clear();
            listView6.Items.Clear();

            //richTextBox1.Clear();
            richTextBox2.Clear();
            //richTextBox3.Clear();
            

            strText.Clear();
            
        }


        private void PrintControlInterface()
        {
            
        }

        private void ClearDictionary()
        {
            words.Clear();

            adverbs.Clear();
            verbs.Clear();
            nouns.Clear();
            adjectives.Clear();
            numerals.Clear();
            verbaladverbs.Clear();
            hyperlinks.Clear();
            particles.Clear();
            preposition.Clear();
            pronounce.Clear();
            unions.Clear();
            interjection.Clear();
            participium.Clear();
            lastnames.Clear();
            firstnames.Clear();
            abbreviatura.Clear();
            reduction.Clear();

            unknownObjects.Clear();

            listKeyWordsProcess.Clear();
            positive.Clear();
            negative.Clear();
            stopslovs.Clear();

            listWords.Clear();
            toshnota.Clear();
            listDictionary.Clear();
            dictionary.Clear();

            wordsAll.Clear();
            semanticCore.Clear();
            lkpxml.Clear();
            ldxml.Clear();

            globalKey.Clear();
            globalValue.Clear();
            globalKeyZnach.Clear();
            globalValueZnach.Clear();

            globalCustomKey.Clear();
            globalVodaKey.Clear();
            globalValueStop.Clear();

            toshClass.Clear();
            toshAcad.Clear();

            keyMassiveSpace.Clear();
            kmpNoSpace.Clear();

            extendDictionary.Clear();
            keysUnknown.Clear();
        }


        // info window
        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Lexiconer 1.0.7.0.\n (c) 2015\nlexiconer.ru", "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information);                              
        }

        // сохранить как текст
        private void saveAsTextToolStripMenuItem_Click(object sender, EventArgs e)
              {

                  DateTime now = DateTime.Now;         // Use current time
                  string time = now.ToString("_yyyy_MM_d_HH_mm"); // Write to consolermat);

                  //FileInfo finfo = new FileInfo(fileName);
                  //double size = finfo.Length;
                  StreamWriter sw;
                  FileStream toSave;
                  string filepath = di + @"\Lexiconer\Reports\";

                  if (Directory.Exists(filepath))
                  {
                      if (File.Exists(filepath + Path.GetFileNameWithoutExtension(fileName) + time))
                      {
                          toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".txt", FileMode.Append);
                          sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                      }
                      else
                      {
                          toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".txt", FileMode.CreateNew);
                          sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                          sw.WriteLine(new string('/', 55));
                          sw.WriteLine("File:  " + fileName);
                          sw.WriteLine(new string('/', 55));
                          sw.WriteLine("\n\n");
                      }

                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine(new string('/', 55));

                      sw.WriteLine("Часть 1");
                      for (int i = 0; i < listView1.Items.Count; i++)
                          sw.WriteLine(listView1.Items[i].Text + "\t\t\t\t" + listView1.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView1.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 2");
                      for (int i = 0; i < listView3.Items.Count; i++)
                          sw.WriteLine(listView3.Items[i].Text + "\t\t\t\t" + listView3.Items[i].SubItems[1].Text);

                      sw.WriteLine("Часть 5");
                      for (int i = 0; i < listView5.Items.Count; i++)
                          sw.WriteLine(listView5.Items[i].Text + "\t\t\t\t" + listView5.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView5.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 2");
                      for (int i = 0; i < listView2.Items.Count; i++)
                          sw.WriteLine(listView2.Items[i].Text + "\t\t\t\t" + listView2.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView2.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 4");
                      for (int i = 0; i < listView4.Items.Count; i++)
                          sw.WriteLine(listView4.Items[i].Text + "\t\t\t\t" + listView4.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView4.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 6");
                      for (int i = 0; i < listView6.Items.Count; i++)
                          sw.WriteLine(listView6.Items[i].Text + "\t\t\t\t" + listView6.Items[i].SubItems[1].Text);
                      //                        + "\t\t\t\t" + listView6.Items[i].SubItems[2].Text);

                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine("Processing time :   " + processingTime.Negate());
                      ////sw.WriteLine("Size of file: " + (size) + " Bytes.\n              " + size / 1024 + " KBytes");
                      sw.WriteLine("Date of scanning: " + DateTime.Now);
                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine("\n\n\n\n");
                      sw.Close();
                      toSave.Close();
                  }

                  else
                  {

                      Directory.CreateDirectory(filepath);
                      if (File.Exists(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".txt"))
                      {

                          toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".txt", FileMode.Append);
                          sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                      }
                      else
                      {
                          toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".txt", FileMode.CreateNew);
                          sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                          sw.WriteLine(new string('/', 55));
                          sw.WriteLine("File:  " + fileName);
                          sw.WriteLine(new string('/', 55));
                          sw.WriteLine("\n\n");
                      }

                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine(new string('/', 55));

                      for (int i = 0; i < listView1.Items.Count; i++)
                          sw.WriteLine(listView1.Items[i].Text + "\t\t\t\t" + listView1.Items[i].SubItems[1].Text
                              + "\t\t\t\t" + listView1.Items[i].SubItems[2].Text);

                      for (int i = 0; i < listView3.Items.Count; i++)
                          sw.WriteLine(listView3.Items[i].Text + "\t\t\t\t" + listView3.Items[i].SubItems[1].Text);

                      sw.WriteLine("Часть 5");
                      for (int i = 0; i < listView5.Items.Count; i++)
                          sw.WriteLine(listView5.Items[i].Text + "\t\t\t\t" + listView5.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView5.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 2");
                      for (int i = 0; i < listView2.Items.Count; i++)
                          sw.WriteLine(listView2.Items[i].Text + "\t\t\t\t" + listView2.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView2.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 4");
                      for (int i = 0; i < listView4.Items.Count; i++)
                          sw.WriteLine(listView4.Items[i].Text + "\t\t\t\t" + listView4.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView4.Items[i].SubItems[2].Text);

                      sw.WriteLine("Часть 6");
                      for (int i = 0; i < listView6.Items.Count; i++)
                          sw.WriteLine(listView6.Items[i].Text + "\t\t\t\t" + listView6.Items[i].SubItems[1].Text
                          + "\t\t\t\t" + listView6.Items[i].SubItems[2].Text);

                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine("Processing time :   " + processingTime.Negate());
                      ////sw.WriteLine("Size of file: " + (size) + " Bytes.\n              " + size / 1024 + " KBytes");
                      sw.WriteLine("Date of scanning: " + DateTime.Now);
                      sw.WriteLine(new string('/', 55));
                      sw.WriteLine('\n');
                      sw.Close();
                      toSave.Close();
                  }
                  label3.Text = "Файл сохранен";
                  Thread.CurrentThread.Join(1300);
                  label3.Text = "";
        }

        // сохранить как xml
        private void saveAsXMLToolStripMenuItem_Click(object sender, EventArgs e)
              {
                  XmlDocument doc = new XmlDocument();

                  DateTime now = DateTime.Now;         // Use current time
                  string time = now.ToString("_yyyy_MM_d_HH_mm"); // Write to consolermat);
                  string filepath = di + @"\Lexiconer\Reports\";

                  if (Directory.Exists(filepath))
                  {
                      if (File.Exists(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".xml"))
                      {
                          doc.Load(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".xml");
                      }
                      else
                      {
                          XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", Encoding.GetEncoding(0).ToString(), "yes");
                          doc.AppendChild(decl);
                          doc.LoadXml("<root></root>");                         
                          XmlElement rootIfNo = doc.DocumentElement;
                          rootIfNo.SetAttribute("FileName", fileName);
                      }

                      XmlElement stat = doc.CreateElement("Statistics");
                      stat.SetAttribute("DateTime", DateTime.Now.ToString());
                      stat.SetAttribute("SizeOfFile",fileName.Length.ToString());
                      stat.SetAttribute("TimeOfProcessing", processingTime.Negate().ToString());

                      XmlElement nounss = doc.CreateElement("Nouns");
                      nounss.SetAttribute("Count", nouns.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in nouns)
                      {
                          XmlElement noun = doc.CreateElement("noun");
                          noun.SetAttribute("Valuable", kvp.Value.ToString());
                          noun.InnerText = kvp.Key;
                          nounss.AppendChild(noun);
                      }
                      stat.AppendChild(nounss);

                      XmlElement adjectivess = doc.CreateElement("Adjectives");
                      adjectivess.SetAttribute("Count", adjectives.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in adjectives)
                      {
                          XmlElement adjective = doc.CreateElement("adjective");
                          adjective.SetAttribute("Valuable", kvp.Value.ToString());
                          adjective.InnerText = kvp.Key;
                          adjectivess.AppendChild(adjective);
                      }
                      stat.AppendChild(adjectivess);

                      XmlElement verbss = doc.CreateElement("Verbs");
                      verbss.SetAttribute("Count", verbs.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in verbs)
                      {
                          XmlElement verb = doc.CreateElement("verb");
                          verb.SetAttribute("Valuable", kvp.Value.ToString());
                          verb.InnerText = kvp.Key;
                          verbss.AppendChild(verb);
                      }
                      stat.AppendChild(verbss);

                      XmlElement numeralss = doc.CreateElement("Numerals");
                      numeralss.SetAttribute("Count", numerals.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in numerals)
                      {
                          XmlElement numeral = doc.CreateElement("numeral");
                          numeral.SetAttribute("Valuable", kvp.Value.ToString());
                          numeral.InnerText = kvp.Key;
                          numeralss.AppendChild(numeral);
                      }
                      stat.AppendChild(numeralss);


                      XmlElement verbaladverbss = doc.CreateElement("Verbaladverbs");
                      verbaladverbss.SetAttribute("Count", verbaladverbs.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in verbaladverbs)
                      {
                          XmlElement verbaladverb = doc.CreateElement("verbaladverb");
                          verbaladverb.SetAttribute("Valuable", kvp.Value.ToString());
                          verbaladverb.InnerText = kvp.Key;
                          verbaladverbss.AppendChild(verbaladverb);
                      }
                      stat.AppendChild(verbaladverbss);


                      //XmlElement hyperlinkss = doc.CreateElement("Hyperlinks");
                      //hyperlinkss.SetAttribute("Count", hyperlinks.Count.ToString());
                      //foreach (KeyValuePair<string, int> kvp in hyperlinks)
                      //{
                      //    XmlElement hyperlink = doc.CreateElement("hyperlink");
                      //    hyperlink.SetAttribute("Valuable", kvp.Value.ToString());
                      //    hyperlink.InnerText = kvp.Key;
                      //    hyperlinkss.AppendChild(hyperlink);
                      //}
                      //stat.AppendChild(hyperlinkss);

                      XmlElement particless = doc.CreateElement("Particles");
                      particless.SetAttribute("Count", particles.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in particles)
                      {
                          XmlElement particle = doc.CreateElement("particle");
                          particle.SetAttribute("Valuable", kvp.Value.ToString());
                          particle.InnerText = kvp.Key;
                          particless.AppendChild(particle);
                      }
                      stat.AppendChild(particless);


                      XmlElement prepositions = doc.CreateElement("Preposition");
                      prepositions.SetAttribute("Count", preposition.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in preposition)
                      {
                          XmlElement prepositio = doc.CreateElement("preposition");
                          prepositio.SetAttribute("Valuable", kvp.Value.ToString());
                          prepositio.InnerText = kvp.Key;
                          prepositions.AppendChild(prepositio);
                      }
                      stat.AppendChild(prepositions);


                      XmlElement pronounces = doc.CreateElement("Pronounces");
                      pronounces.SetAttribute("Count", pronounce.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in pronounce)
                      {
                          XmlElement pronounc = doc.CreateElement("pronounce");
                          pronounc.SetAttribute("Valuable", kvp.Value.ToString());
                          pronounc.InnerText = kvp.Key;
                          pronounces.AppendChild(pronounc);
                      }
                      stat.AppendChild(pronounces);


                      XmlElement unionss = doc.CreateElement("Unions");
                      unionss.SetAttribute("Count", unions.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in unions)
                      {
                          XmlElement union = doc.CreateElement("union");
                          union.SetAttribute("Valuable", kvp.Value.ToString());
                          union.InnerText = kvp.Key;
                          unionss.AppendChild(union);
                      }
                      stat.AppendChild(unionss);


                      XmlElement interjections = doc.CreateElement("Interjection");
                      interjections.SetAttribute("Count", interjection.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in interjection)
                      {
                          XmlElement interjectio = doc.CreateElement("interjection");
                          interjectio.SetAttribute("Valuable", kvp.Value.ToString());
                          interjectio.InnerText = kvp.Key;
                          interjections.AppendChild(interjectio);
                      }
                      stat.AppendChild(interjections);

                      XmlElement participiums = doc.CreateElement("Participiums");
                      participiums.SetAttribute("Count", participium.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in participium)
                      {
                          XmlElement participiu = doc.CreateElement("participium");
                          participiu.SetAttribute("Valuable", kvp.Value.ToString());
                          participiu.InnerText = kvp.Key;
                          participiums.AppendChild(participiu);
                      }
                      stat.AppendChild(participiums);

                      XmlElement lastnamess = doc.CreateElement("Lastnames");
                      lastnamess.SetAttribute("Count", lastnames.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in lastnames)
                      {
                          XmlElement lastname = doc.CreateElement("lastname");
                          lastname.SetAttribute("Valuable", kvp.Value.ToString());
                          lastname.InnerText = kvp.Key;
                          lastnamess.AppendChild(lastname);
                      }
                      stat.AppendChild(lastnamess);


                      XmlElement firstnamess = doc.CreateElement("Firstnames");
                      firstnamess.SetAttribute("Count", firstnames.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in firstnames)
                      {
                          XmlElement firstname = doc.CreateElement("firstname");
                          firstname.SetAttribute("Valuable", kvp.Value.ToString());
                          firstname.InnerText = kvp.Key;
                          firstnamess.AppendChild(firstname);
                      }
                      stat.AppendChild(firstnamess);


                      XmlElement abbreviaturas = doc.CreateElement("Abbreviaturas");
                      abbreviaturas.SetAttribute("Count", abbreviatura.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in abbreviatura)
                      {
                          XmlElement abbreviatur = doc.CreateElement("abbreviatura");
                          abbreviatur.SetAttribute("Valuable", kvp.Value.ToString());
                          abbreviatur.InnerText = kvp.Key;
                          abbreviaturas.AppendChild(abbreviatur);
                      }
                      stat.AppendChild(abbreviaturas);


                      XmlElement reductions = doc.CreateElement("Reductions");
                      reductions.SetAttribute("Count", reduction.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in reduction)
                      {
                          XmlElement reductio = doc.CreateElement("reduction");
                          reductio.SetAttribute("Valuable", kvp.Value.ToString());
                          reductio.InnerText = kvp.Key;
                          reductions.AppendChild(reductio);
                      }
                      stat.AppendChild(reductions);


                      XmlElement unknownObjectss = doc.CreateElement("UnknownObjects");
                      unknownObjectss.SetAttribute("Count", unknownObjects.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in unknownObjects)
                      {
                          XmlElement unknownObject = doc.CreateElement("unknownObject");
                          unknownObject.SetAttribute("Valuable", kvp.Value.ToString());
                          unknownObject.InnerText = kvp.Key;
                          unknownObjectss.AppendChild(unknownObject);
                      }
                      stat.AppendChild(unknownObjectss);

                      XmlElement positives = doc.CreateElement("Positive");
                      positives.SetAttribute("Count", positive.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in positive)
                      {
                          XmlElement positiv = doc.CreateElement("positive");
                          positiv.SetAttribute("Valuable", kvp.Value.ToString());
                          positiv.InnerText = kvp.Key;
                          positives.AppendChild(positiv);
                      }
                      stat.AppendChild(positives);


                      XmlElement negatives = doc.CreateElement("Negative");
                      negatives.SetAttribute("Count", negative.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in negative)
                      {
                          XmlElement negativ = doc.CreateElement("negative");
                          negativ.SetAttribute("Valuable", kvp.Value.ToString());
                          negativ.InnerText = kvp.Key;
                          negatives.AppendChild(negativ);
                      }
                      stat.AppendChild(negatives);


                      XmlElement stopslovss = doc.CreateElement("Stopslovs");
                      stopslovss.SetAttribute("Count", stopslovs.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in stopslovs)
                      {
                          XmlElement stopslov = doc.CreateElement("stopslovo");
                          stopslov.SetAttribute("Valuable", kvp.Value.ToString());
                          stopslov.InnerText = kvp.Key;
                          stopslovss.AppendChild(stopslov);
                      }
                      stat.AppendChild(stopslovss);



                      //XmlElement lkpxmls = doc.CreateElement("Keywords");
                      //lkpxmls.SetAttribute("Count", lkpxml.Count.ToString());
                      //foreach (KeyValuePair<string, int> kvp in lkpxml)
                      //{
                      //    XmlElement lkpxml = doc.CreateElement("keywords");
                      //    lkpxml.SetAttribute("Valuable", kvp.Value.ToString());
                      //    lkpxml.InnerText = kvp.Key;
                      //    lkpxmls.AppendChild(lkpxml);
                      //}
                      //stat.AppendChild(lkpxmls);



                      //XmlElement ldxmls = doc.CreateElement("KeywordsExtention");
                      //ldxmls.SetAttribute("Count", ldxml.Count.ToString());
                      //foreach (KeyValuePair<string, int> kvp in ldxml)
                      //{
                      //    XmlElement lxml = doc.CreateElement("keywordextetion");
                      //    lxml.SetAttribute("Valuable", kvp.Value.ToString());
                      //    lxml.InnerText = kvp.Key;
                      //    ldxml.AppendChild(lxml);
                      //}
                      //stat.AppendChild(ldxmls);


//


                      //XmlElement wordsAlls = doc.CreateElement("Allwords");
                      //wordsAlls.SetAttribute("Count", wordsAll.Count.ToString());
                      //foreach (KeyValuePair<string, int> kvp in wordsAll)
                      //{
                      //    XmlElement wordsAl = doc.CreateElement("Allword");
                      //    wordsAl.SetAttribute("Valuable", kvp.Value.ToString());
                      //    wordsAl.InnerText = kvp.Key;
                      //    wordsAlls.AppendChild(lxml);
                      //}
                      //stat.AppendChild(wordsAlls);




                      //XmlElement semanticCores = doc.CreateElement("semanticCore");
                      //semanticCores.SetAttribute("Count", semanticCore.Count.ToString());
                      //foreach (KeyValuePair<string, int> kvp in semanticCore)
                      //{
                      //    XmlElement semanticCor = doc.CreateElement("semanticCore");
                      //    semanticCor.SetAttribute("Valuable", kvp.Value.ToString());
                      //    semanticCor.InnerText = kvp.Key;
                      //    semanticCore.AppendChild(semanticCor);
                      //}
                      //stat.AppendChild(semanticCores);

                      // toshnota, toshClass, toshAcad);





                      ////////////

                      XmlElement adverbss = doc.CreateElement("Adverbs");
                      adverbss.SetAttribute("Count", adverbs.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in adverbs)
                      {
                          XmlElement adverb = doc.CreateElement("adverbs");
                          adverb.SetAttribute("Valuable", kvp.Value.ToString());
                          adverb.InnerText = kvp.Key;
                          adverbss.AppendChild(adverb);
                      }


                      ////////////////

                      stat.AppendChild(numeralss);
                      doc.DocumentElement.AppendChild(stat);
                      doc.Save(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".xml");
                      this.label3.Text = "S A V E D";
                      Thread.CurrentThread.Join(1300);
                      this.label3.Text = "";

                  }
                  else
                  {
                      Directory.CreateDirectory(filepath);
                      XmlDeclaration decl = doc.CreateXmlDeclaration("1.0", Encoding.GetEncoding(0).ToString(), "yes");
                      doc.AppendChild(decl);
                      doc.LoadXml("<root></root>");
                      doc.DocumentElement.SetAttribute("FileName", fileName);

                      XmlElement stat = doc.CreateElement("Statistics");
                      stat.SetAttribute("DateTime", DateTime.Now.ToString());
                      stat.SetAttribute("SizeOfFile",fileName.Length.ToString());
                      stat.SetAttribute("TimeOfProcessing", processingTime.Negate().ToString());

                      XmlElement nounss = doc.CreateElement("Nouns");
                      nounss.SetAttribute("Count", nouns.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in nouns)
                      {
                          XmlElement noun = doc.CreateElement("noun");
                          noun.SetAttribute("Valuable", kvp.Value.ToString());
                          noun.InnerText = kvp.Key;
                          nounss.AppendChild(noun);
                      }
                      stat.AppendChild(nounss);

                      XmlElement adjectivess = doc.CreateElement("Adjectives");
                      adjectivess.SetAttribute("Count", adjectives.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in adjectives)
                      {
                          XmlElement adjective = doc.CreateElement("adjective");
                          adjective.SetAttribute("Valuable", kvp.Value.ToString());
                          adjective.InnerText = kvp.Key;
                          adjectivess.AppendChild(adjective);
                      }
                      stat.AppendChild(adjectivess);

                      XmlElement verbss = doc.CreateElement("Verbs");
                      verbss.SetAttribute("Count", verbs.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in verbs)
                      {
                          XmlElement verb = doc.CreateElement("verb");
                          verb.SetAttribute("Valuable", kvp.Value.ToString());
                          verb.InnerText = kvp.Key;
                          verbss.AppendChild(verb);
                      }
                      stat.AppendChild(verbss);

                      XmlElement numeralss = doc.CreateElement("Numerals");
                      numeralss.SetAttribute("Count", numerals.Count.ToString());
                      foreach (KeyValuePair<string, int> kvp in numerals)
                      {
                          XmlElement numeral = doc.CreateElement("numeral");
                          numeral.SetAttribute("Valuable", kvp.Value.ToString());
                          numeral.InnerText = kvp.Key;
                          numeralss.AppendChild(numeral);
                      }

                      stat.AppendChild(numeralss);
                      doc.DocumentElement.AppendChild(stat);
                      doc.Save(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".xml");
                      this.label3.Text = "S A V E D";
                      Thread.CurrentThread.Join(1300);
                      this.label3.Text = "";
                  }
            
              }

        // сохранить как csv
        private void SaveAsCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DateTime now = DateTime.Now;         // Use current time
            string time = now.ToString("_yyyy_MM_d_HH_mm"); // Write to consolermat);

            FileInfo finfo = new FileInfo(fileName);
            double size = finfo.Length;
            StreamWriter sw;
            FileStream toSave;
            string filepath = di + @"\Lexiconer\Reports\";

            if (Directory.Exists(filepath))
            {
                if (File.Exists(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".csv"))
                {
                    toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time + ".csv", FileMode.Append);
                    sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                }
                else
                {
                    toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time  + ".csv", FileMode.CreateNew);
                    sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                    sw.WriteLine("File:  " + fileName);
                }

                sw.WriteLine("Часть 1");
                for (int i = 0; i < listView1.Items.Count; i++)
                    sw.WriteLine(listView1.Items[i].Text + ";" + listView1.Items[i].SubItems[1].Text
                    + ";" + listView1.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 2");
                for (int i = 0; i < listView3.Items.Count; i++)
                    sw.WriteLine(listView3.Items[i].Text + ";" + listView3.Items[i].SubItems[1].Text);

                sw.WriteLine("Часть 5");
                for (int i = 0; i < listView5.Items.Count; i++)
                    sw.WriteLine(listView5.Items[i].Text + ";" + listView5.Items[i].SubItems[1].Text
                    + ";" + listView5.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 2");
                for (int i = 0; i < listView2.Items.Count; i++)
                    sw.WriteLine(listView2.Items[i].Text + ";" + listView2.Items[i].SubItems[1].Text
                    + ";" + listView2.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 4");
                for (int i = 0; i < listView4.Items.Count; i++)
                    sw.WriteLine(listView4.Items[i].Text + ";" + listView4.Items[i].SubItems[1].Text
                    + ";" + listView4.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 6");
                for (int i = 0; i < listView6.Items.Count; i++)
                    sw.WriteLine(listView6.Items[i].Text + ";" + listView6.Items[i].SubItems[1].Text);
                //                        + ";" + listView6.Items[i].SubItems[2].Text);

                sw.WriteLine(new string('/', 55));
                sw.WriteLine("Processing time :   " + processingTime.Negate());
                sw.WriteLine("Size of file: " + (size) + " Bytes.\n              " + size / 1024 + " KBytes");
                sw.WriteLine("Date of scanning: " + DateTime.Now);
                sw.WriteLine(new string('/', 55));
                sw.WriteLine("\n\n\n\n");
                sw.Close();
                toSave.Close();
            }

            else
            {

                Directory.CreateDirectory(filepath);
                if (File.Exists(filepath + Path.GetFileNameWithoutExtension(fileName) + time  + ".csv"))
                {

                    toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time  + ".csv", FileMode.Append);
                    sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                }
                else
                {
                    toSave = new FileStream(filepath + Path.GetFileNameWithoutExtension(fileName) + time  + ".csv", FileMode.CreateNew);
                    sw = new StreamWriter(toSave, Encoding.GetEncoding(0));
                    sw.WriteLine(new string('/', 55));
                    sw.WriteLine("File:  " + fileName);
                    sw.WriteLine(new string('/', 55));
                    sw.WriteLine("\n\n");
                }

                sw.WriteLine(new string('/', 55));
                sw.WriteLine(new string('/', 55));
                for (int i = 0; i < listView1.Items.Count; i++)
                    sw.WriteLine(listView1.Items[i].Text + ";" + listView1.Items[i].SubItems[1].Text
                        + ";" + listView1.Items[i].SubItems[2].Text);

                for (int i = 0; i < listView3.Items.Count; i++)
                    sw.WriteLine(listView3.Items[i].Text + ";" + listView3.Items[i].SubItems[1].Text);

                sw.WriteLine("Часть 5");
                for (int i = 0; i < listView5.Items.Count; i++)
                    sw.WriteLine(listView5.Items[i].Text + ";" + listView5.Items[i].SubItems[1].Text
                    + ";" + listView5.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 2");
                for (int i = 0; i < listView2.Items.Count; i++)
                    sw.WriteLine(listView2.Items[i].Text + ";" + listView2.Items[i].SubItems[1].Text
                    + ";" + listView2.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 4");
                for (int i = 0; i < listView4.Items.Count; i++)
                    sw.WriteLine(listView4.Items[i].Text + ";" + listView4.Items[i].SubItems[1].Text
                    + ";" + listView4.Items[i].SubItems[2].Text);

                sw.WriteLine("Часть 6");
                for (int i = 0; i < listView6.Items.Count; i++)
                    sw.WriteLine(listView6.Items[i].Text + ";" + listView6.Items[i].SubItems[1].Text
                    + ";" + listView6.Items[i].SubItems[2].Text);

                sw.WriteLine(new string('/', 55));
                sw.WriteLine("Processing time :   " + processingTime.Negate());
                sw.WriteLine("Size of file: " + (size) + " Bytes.\n              " + size / 1024 + " KBytes");
                sw.WriteLine("Date of scanning: " + DateTime.Now);
                sw.WriteLine(new string('/', 55));
                sw.WriteLine('\n');
                sw.Close();
                toSave.Close();
            }
            label3.Text = "Файл сохранен";
            Thread.CurrentThread.Join(1300);
            label3.Text = "";

        }

        // выход        
        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
              {
                  Application.Exit();
              }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e) { }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView2_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView3_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView4_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView5_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView6_SelectedIndexChanged(object sender, EventArgs e) { }
        private void listView7_SelectedIndexChanged(object sender, EventArgs e) { }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e) { }

        // буфер ключевых слов
        private void button3_Click(object sender, EventArgs e)
        {
            // Объявление объекта-получателя из БО
            var adapter = Clipboard.GetDataObject();
            // Если данные в БО представлены в текстовом формате...
            if (adapter.GetDataPresent(DataFormats.Text) == true)
            {
                // то записать их в Text тоже в текстовом формате,
                richTextBox3.Text = adapter.GetData(DataFormats.Text).ToString();
                keywordsclipboard = richTextBox3.Text; //adapter.GetData(DataFormats.Text).ToString();
                
            }
            else // иначе
                richTextBox3.Text = "Запишите что-либо в буфер обмена";
        }
        
        // очистить ключевые слова
        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.Clear();
        }

        // буфер тестируемого текста
        private void button5_Click(object sender, EventArgs e)
        {
            var adapter1 = Clipboard.GetDataObject(); // Объявление объекта-получателя из БО

            if (adapter1.GetDataPresent(DataFormats.Text) == true) // Если данные в БО представлены в текстовом формате...
            {
                richTextBox1.Text = adapter1.GetData(DataFormats.Text).ToString(); // то записать их в Text тоже в текстовом формате
                this.button1.Enabled = true;
            }
            else 
                richTextBox1.Text = "Запишите что-либо в буфер обмена";
}

        // чистить тестируемый текст
        private void button6_Click(object sender, EventArgs e)
        {
            richTextBox3.Clear();
        }

        // очистить все
        private void button8_Click(object sender, EventArgs e)
        {
            ClearInterface();
            ClearDictionary();
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
        }
        private void tabPage4_Click(object sender, EventArgs e)
        {
        }

        // очистить буфер
        private void button10_Click(object sender, EventArgs e)
        {
            Clipboard.Clear();
        }

        // скопировать результаты в буфер
        private void button11_Click(object sender, EventArgs e)
        {

            string copyText = string.Empty;

            if (richTextBox1.SelectionLength == 0 &&
                richTextBox2.SelectionLength == 0 &&
                richTextBox3.SelectionLength == 0 &&
                listView1.SelectedItems.Count == 0 && 
                listView2.SelectedItems.Count == 0 && 
                listView3.SelectedItems.Count == 0 && 
                listView4.SelectedItems.Count == 0 && 
                listView5.SelectedItems.Count == 0 &&
                listView6.SelectedItems.Count == 0 &&
                listView7.SelectedItems.Count == 0)
            {
                copyText += "Резюме" + Environment.NewLine;
                copyText += richTextBox2.Text + Environment.NewLine;

                copyText += "Словоформы" + Environment.NewLine; 
                for (int i = 0; i < listView1.Items.Count; i++)
                    copyText += listView1.Items[i].Text + "\t" + listView1.Items[i].SubItems[1].Text
                        + "\t" + listView1.Items[i].SubItems[2].Text + Environment.NewLine;

                copyText += "Слова" + Environment.NewLine; 
                for (int i = 0; i < listView3.Items.Count; i++)
                    copyText += listView3.Items[i].Text + "\t" + listView3.Items[i].SubItems[1].Text + Environment.NewLine;

                copyText += "Словари" + Environment.NewLine;
                for (int i = 0; i < listView5.Items.Count; i++)
                    copyText += listView5.Items[i].Text + "\t" + listView5.Items[i].SubItems[1].Text
                    + "\t" + listView5.Items[i].SubItems[2].Text + Environment.NewLine;

                copyText += "Ключевики" + Environment.NewLine;
                for (int i = 0; i < listView2.Items.Count; i++)
                    copyText += listView2.Items[i].Text + "\t" + listView2.Items[i].SubItems[1].Text
                    + "\t" + listView2.Items[i].SubItems[2].Text + Environment.NewLine;

                copyText += "Ядро" + Environment.NewLine;
                for (int i = 0; i < listView4.Items.Count; i++)
                    copyText += listView4.Items[i].Text + "\t" + listView4.Items[i].SubItems[1].Text
                    + "\t" + listView4.Items[i].SubItems[2].Text + Environment.NewLine;

                copyText += "Тошнота" + Environment.NewLine;
                for (int i = 0; i < listView6.Items.Count; i++)
                    copyText += listView6.Items[i].Text + "\t" + listView6.Items[i].SubItems[1].Text
                    + "\t" + listView6.Items[i].SubItems[2].Text + "\t" + listView6.Items[i].SubItems[3].Text + Environment.NewLine;
                }

            else {

                if (richTextBox1.SelectionLength != 0) {
                copyText += richTextBox1.SelectedText + Environment.NewLine;
                }

                if (richTextBox2.SelectionLength != 0) {
                copyText += "Резюме" + Environment.NewLine;
                copyText += richTextBox2.Text + Environment.NewLine;
                }

                if (richTextBox3.SelectionLength != 0) {
                copyText += richTextBox3.SelectedText + Environment.NewLine;
                }

                if (listView1.SelectedItems.Count != 0)
                {
                    copyText += "Словоформы" + Environment.NewLine;
                    for (int i = 0; i < listView1.SelectedItems.Count; i++)
                        copyText += listView1.Items[i].Text + "\t" + listView1.Items[i].SubItems[1].Text
                            + "\t" + listView1.Items[i].SubItems[2].Text + Environment.NewLine;
                }

                if (listView3.SelectedItems.Count != 0)
                {
                    copyText += "Слова" + Environment.NewLine;
                    for (int i = 0; i < listView3.SelectedItems.Count; i++)
                        copyText += listView3.Items[i].Text + "\t" + listView3.Items[i].SubItems[1].Text + Environment.NewLine;
                }

                if (listView5.SelectedItems.Count != 0)
                {
                    copyText += "Словари" + Environment.NewLine;
                    for (int i = 0; i < listView5.SelectedItems.Count; i++)
                        copyText += listView5.Items[i].Text + "\t" + listView5.Items[i].SubItems[1].Text
                        + "\t" + listView5.Items[i].SubItems[2].Text + Environment.NewLine;
                }

                if (listView2.SelectedItems.Count != 0)
                {
                    copyText += "Ключевики" + Environment.NewLine;
                    for (int i = 0; i < listView2.SelectedItems.Count; i++)
                        copyText += listView2.Items[i].Text + "\t" + listView2.Items[i].SubItems[1].Text
                        + "\t" + listView2.Items[i].SubItems[2].Text + Environment.NewLine;
                }

                if (listView4.SelectedItems.Count != 0)
                {
                    copyText += "Ядро" + Environment.NewLine;
                    for (int i = 0; i < listView4.SelectedItems.Count; i++)
                        copyText += listView4.Items[i].Text + "\t" + listView4.Items[i].SubItems[1].Text
                        + "\t" + listView4.Items[i].SubItems[2].Text + Environment.NewLine;
                }

                if (listView6.SelectedItems.Count != 0)
                {
                    copyText += "Тошнота" + Environment.NewLine;
                    for (int i = 0; i < listView6.SelectedItems.Count; i++)
                        copyText += listView6.Items[i].Text + "\t" + listView6.Items[i].SubItems[1].Text
                        + "\t" + listView6.Items[i].SubItems[2].Text + "\t" + listView6.Items[i].SubItems[3].Text + Environment.NewLine;
                }
            }

            try
            {
                Clipboard.SetDataObject(copyText, true, 3, 400);
            }

            catch (System.Runtime.InteropServices.ExternalException)
            {
                MessageBox.Show(this, "Не удалось очистить буфер обмена. Возможно буфер обмена используется другим процессом.",
                    "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            var data = new DataObject();
            data.SetData(DataFormats.UnicodeText, true, copyText);
            var thread = new Thread(() => Clipboard.SetDataObject(data, true));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {}

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Properties.Resources.helper, "Справка по программе");
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }
       
    }
}