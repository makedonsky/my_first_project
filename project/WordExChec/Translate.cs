using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace Translate
{
    class TranslateData
    {
        internal TranslateData()
        {

        }
        public string TranslateStr(string str)
        {
           
            string result="";
            int pos = 0, count = 1; //stage = 0,
            //string cur = getcurrencyname(str);
            str=str_reverse(str);
            while(pos<str.Length)
            {
              // stage = getstage(pos);
                if (str.Length - pos >= 3)
                {
                    result += translate(m_strcopy(str, pos, 3),count);
                    pos=pos+3;
                }
                else
                {
                    if (str.Length - pos < 3)
                    {
                        result +=translate(m_strcopy(str,pos,str.Length-pos),count);
                        pos=pos+(str.Length-pos);
                    }
                }
                count++;
            }
            result = strword_reverse(result);
            return result;// +" " + cur; 
        }
        private string strword_reverse(string str)
        {
            string[] strarrst;
            string result;
            string s=str.Trim(' ');
            strarrst = str.Split(' ');
            string[] strarren = new string[strarrst.Length];
            for (int i = 0; i < strarrst.Length;i++)
            {
                strarren[strarrst.Length - i-1] = strarrst[i];
            }
            result = String.Join(" ", strarren);
            result = result.Trim(' ');
            return result;
        }
        public string word_reverse(string str)
        {
            string result="";
            //string[] tempstr = new string(str);
            for (int i = 0; i < str.Length;i++)
            {
                result += str.Substring(str.Length - i - 1, 1);
            }
            
            return result;
        }
        private string m_strcopy(string str,int pos,int num)
        {
            string result="";
            for (int i=0;i<num;i++)
            {
                result +=str[pos+i];
            }
            return result;
        }
        private int getstage(int pos)
        {
            int lvl=0;
            if (pos < 3) { lvl = 1; }
            if ((pos < 3)&&(pos<6)) { lvl = 2; }
            if ((pos < 6) && (pos < 9)) { lvl = 3; }
            if ((pos < 9) && (pos < 12)) { lvl = 4; }
            if ((pos < 12) && (pos < 15)) { lvl = 5; }
            if ((pos < 15) && (pos < 18)) { lvl = 6; }
            return lvl;
        }
        private string str_reverse(string str)
        {
            string result = "";
            for (int i = str.Length-1; i > -1;i--)
            {
                result += str[i];
            }

            return result;
        }
        private string translate(string str,int stage)
        {
            string result="";
           // const char one = '1',two = '2',three = '3',four = '4',five = '5',six = '6',seven = '7',eight = '8',nine = '9',zero = '0';
            for (int i=0;i<str.Length;i++)
            {
                if (i == 0)
                {
                    if ((str.Length>=2)&&(str[1] != '1'))
                    {
                        if (stage == 2)
                        {
                            switch (str[i])
                            {
                                case '1': result += "одна"; break;
                                case '2': result += "две"; break;
                                case '3': result += "три"; break;
                                case '4': result += "четыре"; break;
                                case '5': result += "п€ть"; break;
                                case '6': result += "шесть"; break;
                                case '7': result += "семь"; break;
                                case '8': result += "восемь"; break;
                                case '9': result += "дев€ть"; break;
                            }
                        }
                        else
                        {
                            switch (str[i])
                            {
                                case '1': result += "один"; break;
                                case '2': result += "два"; break;
                                case '3': result += "три"; break;
                                case '4': result += "четыре"; break;
                                case '5': result += "п€ть"; break;
                                case '6': result += "шесть"; break;
                                case '7': result += "семь"; break;
                                case '8': result += "восемь"; break;
                                case '9': result += "дев€ть"; break;
                            }
                        }
                    }
                    else if (str.Length < 2)
                    {
                        if (stage == 2)
                        {
                            switch (str[i])
                            {
                                case '1': result += "одна"; break;
                                case '2': result += "две"; break;
                                case '3': result += "три"; break;
                                case '4': result += "четыре"; break;
                                case '5': result += "п€ть"; break;
                                case '6': result += "шесть"; break;
                                case '7': result += "семь"; break;
                                case '8': result += "восемь"; break;
                                case '9': result += "дев€ть"; break;
                            }
                        }
                        else
                        {
                            switch (str[i])
                            {
                                case '1': result += "один"; break;
                                case '2': result += "два"; break;
                                case '3': result += "три"; break;
                                case '4': result += "четыре"; break;
                                case '5': result += "п€ть"; break;
                                case '6': result += "шесть"; break;
                                case '7': result += "семь"; break;
                                case '8': result += "восемь"; break;
                                case '9': result += "дев€ть"; break;
                            }
                        }
                    }
                }
                if (i == 1)
                {
                    if ((str.Length >= 2) && (str[1] != '1'))
                    {
                        switch (str[i])
                        {
                            case '2': result += "двадцать"; break;
                            case '3': result += "тридцать"; break;
                            case '4': result += "сорок"; break;
                            case '5': result += "п€тьдес€т"; break;
                            case '6': result += "шестьдес€т"; break;
                            case '7': result += "семьдес€т"; break;
                            case '8': result += "восемьдес€т"; break;
                            case '9': result += "дев€носто"; break;
                        }
                    }
                    if ((str.Length >= 2) && (str[1] == '1'))
                    {
                        switch (str[i-1])
                        {
                            case '1': result += "одиннадцать"; break;
                            case '2': result += "двенадцать"; break;
                            case '3': result += "тринадцать"; break;
                            case '4': result += "четырнадцать"; break;
                            case '5': result += "п€тнадцать"; break;
                            case '6': result += "шестнадцать"; break;
                            case '7': result += "семнадцать"; break;
                            case '8': result += "восемнадцать"; break;
                            case '9': result += "дев€тнадцать"; break;
                            case '0': result += "дес€ть"; break;
                        }
                    }
                }
                if (i == 2)
                {
                        switch (str[i])
                        {
                            case '1': result += "сто"; break;
                            case '2': result += "двести"; break;
                            case '3': result += "тристо"; break;
                            case '4': result += "четыресто"; break;
                            case '5': result += "п€тьсот"; break;
                            case '6': result += "шестьсот"; break;
                            case '7': result += "семьсот"; break;
                            case '8': result += "восемьсот"; break;
                            case '9': result += "дев€тьсот"; break;
                        }
                    
                }
                if (str[i] != '0') { result += " "; }
            }
            string stagename = getstagename(stage,str);
            return stagename+" " + result;
        }
        private string get3tname(char str)
        {
            string result = "";
            if (str == '1') { result = "тыс€ча"; }
            if ((str == '2') || (str == '3') || (str == '4')) { result = "тыс€чи"; }
            if ((str != '1') && (str != '2') && (str != '3') && (str != '4')) { result = "тыс€ч"; }
            return result;
        }
        private string get6mname(char str)
        {
            string result = "";
            if (str == '1') { result = "миллион"; }
            if ((str == '2') || (str == '3') || (str == '4')) { result = "миллиона"; }
            if ((str != '1') && (str != '2') && (str != '3') && (str != '4')) { result = "миллионов"; }
            return result;

        }

    private string getstagename(int stage, string str)
    {
        string result = "";
        Regex exp = new Regex(@"[0-9]|(?=([0-9]))[0-9]|(?=([0-9][0-9]))[0-9]|1[0-9]|(?=([0-9]))1[0-9]");
        Match m = exp.Match(str);

        switch(stage)
        {
            case 2: /*if (m.Value.ToString() == "1") { result = "тыс€ча"; }
                if ((m.Value.ToString() == "2") || (m.Value.ToString() == "3") || (m.Value.ToString() == "4")) { result = "тыс€чи"; }
                if ((m.Value.ToString() != "1") && (m.Value.ToString() != "2") && (m.Value.ToString() != "3") && (m.Value.ToString() != "4")) { result = "тыс€ч"; }*/
                if (Convert.ToInt32(str) > 0)
                {
                    if (str.Length < 2)
                    {
                        result = get3tname(str[0]);
                    }
                    else
                    {
                        if (str.Length == 2)
                        {
                            if (str[0] == '0')
                            {
                                result = "тыс€ч";
                            }
                            else
                            {
                                result = get3tname(str[0]);
                            }
                        }
                        if (str.Length == 3)
                        {
                            if ((str[1] == '0') && (str[2] == '0'))
                            {
                                result = "тыс€ч";
                            }
                            else
                            {
                                if (str[0] == '0')
                                {
                                    result = "тыс€ч";
                                }
                                else
                                {
                                    result = get3tname(str[0]);
                                }
                            }
                        }
                    }
                }
                break;
            case 3: /*if (m.Value.ToString() == "1") { result = "миллион"; }
                if ((m.Value.ToString() == "2") || (m.Value.ToString() == "3") || (m.Value.ToString() == "4")) { result = "миллиона"; }
                if ((m.Value.ToString() != "1") && (m.Value.ToString() != "2") && (m.Value.ToString() != "3") && (m.Value.ToString() != "4")) { result = "миллионов"; }*/
                if (Convert.ToInt32(str) > 0)
                {
                    if (str.Length < 2)
                    {
                        result = get6mname(str[0]);
                    }
                    else
                    {
                        if (str.Length == 2)
                        {
                            if (str[1] == '1')
                            {
                                result = "миллионов";
                            }
                            else
                            {
                                result = get6mname(str[0]);
                            }
                        }
                        if (str.Length == 3)
                        {
                            if ((str[1] == '0') && (str[2] == '0'))
                            {
                                result = "миллионов";
                            }
                            else
                            {
                                if (str[1] == '1')
                                {
                                    result = "миллионов";
                                }
                                else
                                {
                                    result = get6mname(str[1]);
                                }
                            }
                        }
                    }
                }
                break;
            case 4: if (m.Value.ToString() == "1") { result = "миллиард"; }
                if ((m.Value.ToString() == "2") || (m.Value.ToString() == "3") || (m.Value.ToString() == "4")) { result = "миллиарда"; }
                if ((m.Value.ToString() != "1") && (m.Value.ToString() != "2") && (m.Value.ToString() != "3") && (m.Value.ToString() != "4")) { result = "миллиарда"; }
                break;
            case 5: if (m.Value.ToString() == "1") { result = "триллион"; }
                if ((m.Value.ToString() == "2") || (m.Value.ToString() == "3") || (m.Value.ToString() == "4")) { result = "триллиона"; }
                if ((m.Value.ToString() != "1") && (m.Value.ToString() != "2") && (m.Value.ToString() != "3") && (m.Value.ToString() != "4")) { result = "триллионов"; }
                break;
        }
        return result;
    }
       /* private string getcurrencyname(string str)
        {
            string result = "";
            int currency = 0;
            if (this.checkBox1.Checked) { currency = 1; }
            if (this.checkBox2.Checked) { currency = 2; }
            if (this.checkBox3.Checked) { currency = 3; }
           //Regex exp = new Regex(@"[0-9]|(?=([0-9]))[0-9]|(?=([0-9][0-9]))[0-9]|1[0-9]|(?=([0-9]))1[0-9]");
            //Match m = exp.Match(str);
            switch (currency)
            {
                case 1: if (str[str.Length - 1].ToString() == "1") { result = "рубль"; }
                    if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "рубл€"; }
                    if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "рублей"; }
                    break;
                case 2: if (str[str.Length - 1].ToString() == "1") { result = "евро"; }
                    if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "евро"; }
                    if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "евро"; }
                    break;
                case 3: if (str[str.Length - 1].ToString() == "1") { result = "доллар"; }
                    if ((str[str.Length - 1].ToString() == "2") || (str[str.Length - 1].ToString() == "3") || (str[str.Length - 1].ToString() == "4")) { result = "доллара"; }
                    if ((str[str.Length - 1].ToString() != "1") && (str[str.Length - 1].ToString() != "2") && (str[str.Length - 1].ToString() != "3") && (str[str.Length - 1].ToString() != "4")) { result = "долларов"; }
                    break;
            }
            return result;
        }*/

    }
}
