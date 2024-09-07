using System;
using System.Collections.Generic;
using System.Data;

namespace appquanlysothu
{
    public class Creature
    {
        public string Name { get; set; }
        public string Type { get; set; }
        public string Id { get; set; }
        public string Barn { get; set; }
        public string Note { get; set; }
        public int Age { get; set; }
        public int Sex { get; set; }
        public int Condition { get; set; }
        public int Carn_herbivore { get; set; }
        public float Weight { get; set; }
        public DateTime Entry { get; set; }
        public DateTime Birth { get; set; }

        //Default Constructor
        public Creature()
        {
        }

        //Parameterized Constructor
        public Creature(string id, string name, string type, string barn, int sex, int condition, int carn_herbivore, float weight, DateTime entry, DateTime birth, string note)
        {
            Name = name;
            Type = type;
            Id = id;
            Barn = barn;
            Sex = sex;
            Condition = condition;
            Carn_herbivore = carn_herbivore;
            Weight = weight;
            Entry = entry;
            Birth = birth;
            Note = note;
            var today = DateTime.Today;
            Age = today.Year - birth.Year;
            if (birth.Date > today.AddYears(-Age)) Age--;
        }

        private DateTime CreateDate(int d, int m, int y)
        {
            DateTime Date = new DateTime(y, m, d);
            return Date.Date;
        }
    }
}