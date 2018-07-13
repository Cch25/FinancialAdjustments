using System;
using System.Collections.Generic;

namespace GenerareData
{
    public class GeneratorDate
    {
        public List<DateTime> GenereazaData(DateTime dataStart, DateTime dataSfarsit)
        {
            if (dataStart > dataSfarsit)
                return null;
            var rv = new List<DateTime>();
            var temp = dataStart;
            do
            {
                var dayOfWeek = temp.DayOfWeek;
                if (dayOfWeek != DayOfWeek.Saturday && dayOfWeek != DayOfWeek.Sunday)
                    rv.Add(temp);
                temp = temp.AddDays(1);
            } while (temp <= dataSfarsit);
            return rv;
        }
        public List<DateTime> GenereazaLuna(DateTime start, DateTime final)
        {
            var ln = new List<DateTime>();
            int an = start.Year, luna = start.Month;
            start = new DateTime(an, luna, 1).AddMonths(1).AddDays(-1);
            do
            {
                if (start.DayOfWeek != DayOfWeek.Saturday && start.DayOfWeek != DayOfWeek.Sunday)
                    ln.Add(start);
                else if (start.DayOfWeek == DayOfWeek.Sunday)
                {
                    start = start.AddDays(-2);
                    ln.Add(start);
                }
                else if (start.DayOfWeek == DayOfWeek.Saturday)
                {
                    start = start.AddDays(-1);
                    ln.Add(start);
                }
                luna++;
                if (luna > 12)
                {
                    luna = 1;
                    an++;
                }
                start = new DateTime(an, luna, 1).AddDays(-1);
            } while (start < final);
            return ln;
        }
        public List<DateTime> GenereazaPrimeleZile(DateTime inceputA, DateTime sfarsitA)
        {
            List<DateTime> la = new List<DateTime>();
            int an; int numara = 1;
            an = inceputA.Year;
            int x = 1;
            inceputA = new DateTime(an, 12, 31).AddDays(x);
            do
            {
                if (inceputA.DayOfWeek != DayOfWeek.Saturday && inceputA.DayOfWeek != DayOfWeek.Sunday)
                {
                    la.Add(inceputA);
                    numara++;
                }
                x++;
                if (numara > 5)
                {
                    x = 1;
                    an++;
                    numara = 1;
                }
                inceputA = new DateTime(an, 12, 31).AddDays(x);
            } while (inceputA < sfarsitA);
            return la;
        }
    }
}
