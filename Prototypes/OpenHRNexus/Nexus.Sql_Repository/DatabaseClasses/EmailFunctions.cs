using Nexus.Common.Interfaces;

namespace Nexus.Sql_Repository.DatabaseClasses
{
    public static class EmailFunctions
    {
        public static string FormatBody(string template, object[] data, string language)
        {
            string body;

            // Test template is valid

            // Translate for regional stuff

            // Translate dictionary text

            IDictionary dictionary;

            try
            {
                body = string.Format(template, data);
            }

            catch
            {
                body = template;
            }

            return body;

        }


    }
}
