using Nexus.Common.Models;
using OpenHRNexus.Common.Enums;
using System.Diagnostics;

namespace Nexus.Common.Classes
{
    public class ProcessElement
    {
        public int Id { get; set; }
        public ProcessElementType Type { get; set; }
        public int Sequence { get; set; }
        public void Validate() {
            // Perform validation?
            //probably make an abstract class and inheirt
            Debug.Print("codestub here");
        }

        public ProcessFormElement WebForm { get; set; }

    }
}
