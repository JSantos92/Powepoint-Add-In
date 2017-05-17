using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyRibbonAddIn
{
    public class Question
    {
        public string Title { get; set; }
        public List<string> Answers { get; set; }
        public string Type { get; set; }
        public List<Choice> Choices { get; set; }

        public Question()
        {
            this.Choices = new List<Choice>();
            this.Answers = new List<string>();
        }
    }

    public class Choice
    {
        public string option { get; set; }
        public int count { get; set; }
        public string row { get; set; }

        public Choice(string gridRow, string option, int count)
        {
            this.option = option;
            this.count = count;
            row = gridRow;
        }
    }


}
