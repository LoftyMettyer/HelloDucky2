using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Entity;
using DynamicDb;

namespace AsWorkflows
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        
        private void button6_Click(object sender, EventArgs e)
        {

            var dcf = new DynamicClassFactory();

            var context = new DynamicEntities();

//            //var bigBlah = (from t in context.DynamicTemplates
//            //               select t);

////            bigBlah.Load();
//            foreach (var blah in context.DynamicTemplates)
//            {
//                blah.Name = blah.Name + "ddd";
//                Debug.Print(blah.Name);
//            }


        //    Database.SetInitializer<MyContext>(null);
            var templates = (from t in context.DynamicTemplates.Include("DynamicTemplateAttributes").Include("DynamicTemplateAttributes.DynamicAttribute")
                             select t);

            foreach (var dynamicTemplate in templates)
            {
                var type = CreateType(dcf, dynamicTemplate.Name, dynamicTemplate.DynamicTemplateAttributes);
               // result.AddTable(type);
            }

            //      return result;
           // return null;

        }

        private Type CreateType(DynamicClassFactory dcf, string name, ICollection<DynamicTemplateAttribute> dynamicAttributes)
        {
            var props = dynamicAttributes.ToDictionary(da => da.DynamicAttribute.Name, da => typeof(string));
            var t = dcf.CreateDynamicType<BaseDynamicEntity>(name, props);
            return t;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var sysMan = new SystemManager();
            sysMan.CreateTables();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var dataLibrary = new DataLibrary();
            dataLibrary.GetAllCustomerTables();
        }
    }
}
