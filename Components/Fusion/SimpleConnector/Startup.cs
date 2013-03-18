using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NServiceBus;
using Fusion.Messages.Example;

namespace Subscriber1
{
    public class Startup : IWantToRunAtStartup
    {
            public IBus Bus { get; set; }

            public void Run()
            {
                
                Console.WriteLine("This will send Service User Updates to the fusion publisher");

                for (; ; )
                {
                    Console.Write("Enter guid, or enter for new guid: ");
                    string r = Console.ReadLine();

                    if (String.IsNullOrWhiteSpace(r))
                    {
                        r = Guid.NewGuid().ToString();
                    }

                    Console.Write("Forename: ");
                    string forename = Console.ReadLine();

                    Console.Write("Surname: ");
                    string surname = Console.ReadLine();

                    //OnHoldMessage onHoldMessage = new OnHoldMessage
                    //{
                    //    Message = new ServiceUserUpdateMessage
                    //    {
                    //         Id = Guid.NewGuid(),
                    //    CreatedUtc = DateTime.Now,
                    //    Xml = String.Format("<serviceUserUpdate><ref>{0}</ref><forename>{1}</forename><surname>{2}</surname></serviceUserUpdate>",
                    //        r, forename, surname),
                    //    Originator = "Subscriber1",
                    //    EntityRef = new Guid(r)
                    //    }
                    //};

                    //Bus.SendLocal(onHoldMessage);

                    Bus.Send<ServiceUserUpdateRequest>(m =>
                    {
                        m.Id = Guid.NewGuid();
                        m.CreatedUtc = DateTime.Now;
                        m.Xml = String.Format("<serviceUserUpdate><ref>{0}</ref><forename>{1}</forename><surname>{2}</surname></serviceUserUpdate>",
                            r, forename, surname);
                        m.Originator = "Subscriber1";
                        m.EntityRef = new Guid(r);
                    });

                    Console.WriteLine("Published!");

                }
            }

            public void Stop()
            {

            }
    }
    
}
