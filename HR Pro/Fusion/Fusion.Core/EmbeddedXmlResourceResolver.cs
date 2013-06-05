using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml;

namespace Fusion.Core
{

    public class EmbeddedXmlResourceString
    {
        public string Name
        {
            get;
            set;
        }
        public string Data
        {
            get;
            set;
        }
    }

    public class EmbeddedXmlResourceResolver : XmlUrlResolver
    {

        public EmbeddedXmlResourceResolver()
        {
        }

        IEnumerable<EmbeddedXmlResourceString> resources = null;

        public EmbeddedXmlResourceResolver(IEnumerable<EmbeddedXmlResourceString> resources)
        {
            this.resources = resources;
        }

        List<Uri> _resolvedUris = new List<Uri>();

        public override object GetEntity(Uri absoluteUri, string role, Type ofObjectToReturn)
        {

            Stream stream = null;


            switch (absoluteUri.Scheme)
            {
                case "string":
                    if (resources == null)
                        return null;

                    string path = absoluteUri.AbsolutePath;
                    if (path.StartsWith("/"))
                    {
                        path = path.Substring(1);
                    }

                    var resource = resources.FirstOrDefault(a => a.Name == path);                    if (resource == null) return null;

                    return new MemoryStream(Encoding.ASCII.GetBytes(resource.Data));

                case "res":
                    // Handled res:// scheme requests against 
                    // a named assembly with embedded resources

                    Assembly assembly = null;
                    string assemblyfilename = absoluteUri.Host;

                    try
                    {
                        //if (string.Compare(Assembly.GetEntryAssembly().GetName().Name, assemblyfilename, true) == 0)
                        //{       

                        if (string.Compare(Assembly.GetExecutingAssembly().GetName().Name, assemblyfilename, true) == 0)
                        {
                            assembly = Assembly.GetExecutingAssembly();
                        }
                        else
                        {
                            Assembly entryAssembly = Assembly.GetEntryAssembly();

                            if (entryAssembly != null && string.Compare(Assembly.GetEntryAssembly().GetName().Name, assemblyfilename, true) == 0)
                            {
                                assembly = Assembly.GetEntryAssembly();
                            }
                            else
                            {
                                assembly = Assembly.Load(AssemblyName.GetAssemblyName(assemblyfilename + ".dll"));
                            }
                        }

                        string resourceName = absoluteUri.AbsolutePath.Replace('/', '.');
                        if (resourceName.StartsWith("."))
                            resourceName = resourceName.Substring(1);

                        stream = assembly.GetManifestResourceStream(resourceName);
                        if (stream == null)
                        {
                            Trace.WriteLine("Could not find resource {0}", resourceName);
                        }
                    }

                    catch (Exception e)
                    {
                        Trace.WriteLine(e.Message);
                    }
                    return stream;

                default:
                    // Handle file:// and http:// 
                    // requests from the XmlUrlResolver base class
                    stream = (Stream)base.GetEntity(absoluteUri, role, ofObjectToReturn);

                    return stream;
            }
        }
    }
}