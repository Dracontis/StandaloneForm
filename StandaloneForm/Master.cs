using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StandaloneForm
{
    public class Master
    {
        // Какой вуз окончил
        public String University { private set; get; }
        // Диплом
        public String Diploma { private set; get; }
        //Специальность
        public Specialization[] Specs { private set; get; }

        public Master(String University,
            String Diploma,
            Specialization[] Specs)
        {
            this.University = University;
            this.Diploma = Diploma;
            this.Specs = Specs;
        }
        public Master()
        {

        }
    }
}