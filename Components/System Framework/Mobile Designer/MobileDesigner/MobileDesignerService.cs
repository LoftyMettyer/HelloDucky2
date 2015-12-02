using System.Windows.Forms;
using NHibernate;
using System.Runtime.InteropServices;

namespace MobileDesigner
{
    [ComVisible(true)]
    public class MobileDesignerSerivce
    {
        private static bool _initialised;
        private static string _databasePath;
        internal static ISessionFactory SessionFactory { get; private set; }

        public static void Initialise(string databasePath)
        {
            if (!_initialised) {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                _initialised = true;
            }

            if(_databasePath == null || _databasePath != databasePath)
            {
                SessionFactory = DataManager.BuildSessionFactory("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + databasePath);
                _databasePath = databasePath;
            }
                
        }

        /// <summary>
        /// VB6 specific method cos it cant call static methods
        /// </summary>
        public void InitialiseForVB6(string databasePath)
        {
            Initialise(databasePath);
        }
    }
}
