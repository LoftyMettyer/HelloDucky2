namespace Prototype.NHibernateTypeSerialization.Domain.Maps
{
    using System;
    using System.Data;
    using System.Data.Common;
    using System.IO;
    using System.Text;
    using System.Xml;
    using System.Xml.Serialization;
    using NHibernate.SqlTypes;
    using NHibernate.UserTypes;
    using Prototype.NHibernateTypeSerialization.Persistance;

    public class XmlClassType<T> : IUserType where T : ICloneable
    {
        public virtual new bool Equals(object x, object y)
        {
            if (x != null) return x.Equals(y);
            else if (y != null) return y.Equals(x);
            else return true;
        }

        public object NullSafeGet(IDataReader rs, string[] names, object owner)
        {
            if (names.Length != 1)
                throw new InvalidOperationException("names array has more than one element. can't handle this!");

            XmlDocument document = new XmlDocument();
            string val = rs[names[0]] as string;
            if (val != null)
            {
                document.LoadXml(val);
                return Load(document);
            }
            return null;
        }

        protected virtual XmlDocument Save(T data)
        {
            return SaveValue(data);
        }

        protected virtual XmlDocument SaveValue(object data)
        {
            using (MemoryStream stream = new MemoryStream())
            {
                Type type = data.GetType();
                XmlWriter writer = new XmlTextWriter(stream, Encoding.Unicode);

                XmlSerializer serializer = new XmlSerializer(type);
                serializer.Serialize(writer, data);

                stream.Seek(0, SeekOrigin.Begin);

                XmlDocument doc = new XmlDocument();
                doc.Load(stream);

                XmlAttribute typeAttribute = doc.CreateAttribute("type");
                typeAttribute.Value = GetFullName(type);
                doc.DocumentElement.Attributes.Append(typeAttribute);

                return doc;
            }
        }

        protected virtual T Load(XmlDocument value)
        {
            return (T)LoadValue(value.DocumentElement);
        }

        protected virtual object LoadValue(XmlNode value)
        {
            if (value != null)
            {
                string typeName = value.Attributes["type"].Value;
                Type type = Type.GetType(typeName);

                if (type == null)
                    throw new Exception(string.Format("No implementing type could be found for the type: {0}", typeName));

                using (MemoryStream stream = new MemoryStream())
                {
                    XmlDocument document = new XmlDocument();
                    document.AppendChild(document.ImportNode(value, true));

                    XmlSerializer serializer = new XmlSerializer(type);
                    document.Save(stream);
                    stream.Seek(0, SeekOrigin.Begin);

                    return serializer.Deserialize(stream);
                }
            }

            return null;
        }

        public void NullSafeSet(IDbCommand cmd, object value, int index)
        {
            DbParameter parameter = (DbParameter)cmd.Parameters[index];

            if (value == null)
            {
                parameter.Value = DBNull.Value;
                return;
            }

            parameter.Value = Save((T)value).OuterXml;
        }

        public object DeepCopy(object value)
        {
            if (value == null) return null;
            T other = (T)value;
            return other.Clone();
        }

        public SqlType[] SqlTypes
        {
            get
            {
                return new SqlType[] { new SqlXmlType() };
            }
        }

        public Type ReturnedType
        {
            get { return typeof(XmlDocument); }
        }

        public bool IsMutable
        {
            get { return true; }
        }

        #region IUserType Members

        public object Assemble(object cached, object owner)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public object Disassemble(object value)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public int GetHashCode(object x)
        {
            throw new Exception("The method or operation is not implemented.");
        }

        public object Replace(object original, object target, object owner)
        {
            return original;
        }

        #endregion

        private static string GetFullName(Type type)
        {
            return string.Format("{0},{1}", type.FullName, type.Assembly.GetName().Name);
        }
    }
}