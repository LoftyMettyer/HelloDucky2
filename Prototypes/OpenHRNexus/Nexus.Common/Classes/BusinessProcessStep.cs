using OpenHRNexus.Common.Enums;
using System.Diagnostics;

namespace Nexus.Common.Classes
{
    public class BusinessProcessStep
    {
        public int Id { get; set; }
        public BusinessProcessStepType Type { get; set; }

        public void Validate() {
            // Perform validation?
            //probably make an abstract class and inheirt
            Debug.Print("codestub here");
        }


    }
}
