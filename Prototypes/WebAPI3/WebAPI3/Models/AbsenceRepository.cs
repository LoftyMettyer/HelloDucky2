using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebAPI3.Models
{
    public class AbsenceRepository
    {

        private List<Absence> absences = new List<Absence>();
        private int _nextId = 1;

        public AbsenceRepository()
        {
        }

        public IEnumerable<Absence> GetAll()
        {
            return absences;
        }

        public Absence Get(int id)
        {
            return absences.Find(p => p.Id == id);
        }

        public Absence Add(Absence item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }
            item.Id = _nextId++;
            absences.Add(item);
            return item;
        }

        public void Remove(int id)
        {
            absences.RemoveAll(p => p.Id == id);
        }

        public bool Update(Absence item)
        {
            if (item == null)
            {
                throw new ArgumentNullException("item");
            }
            int index = absences.FindIndex(p => p.Id == item.Id);
            if (index == -1)
            {
                return false;
            }
            absences.RemoveAt(index);
            absences.Add(item);
            return true;
        }
    }

}