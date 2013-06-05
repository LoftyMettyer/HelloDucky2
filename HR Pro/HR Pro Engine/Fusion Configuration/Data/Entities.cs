using System;
using System.Collections.Generic;

namespace Fusion
{
	public enum DataType
	{
		Logic = -7,
		Unknown1 = -4,
		Unknown2 = -3,
		Ole = -2,
		Unknown3 = -1,
		Numeric = 2,
		Integer = 4,
		Date = 11,
		Character = 12,
		Guid = 15
	}

	public class Entity
	{
		public virtual int Id { get; set; }
	}

	public class Table : Entity
	{
		public virtual string Name { get; set; }
		public virtual IList<Column> Columns { get; protected set; }
		public override string ToString() { return Name; } }

	public class Column : Entity
	{
		public virtual Table Table { get; set; }
		public virtual string Name { get; set; }
		public virtual DataType DataType { get; set; }
		public virtual int Size { get; set; }
		public virtual int Decimals { get; set; }
		public virtual int LookupTableId { get; set; }
		public override string ToString() { return Name; }
	}

	public class FusionCategory : Entity
	{
		public virtual string Name { get; set; }
		public virtual Table Table { get; set; } 
		public virtual IList<FusionElement> Elements { get; protected set; }
	}

	public class FusionElement : Entity
	{
		public virtual FusionCategory Category { get; set; }
		public virtual string Name { get; set; }
		public virtual string Description { get; set; }
		public virtual DataType DataType { get; set; }
		public virtual int? MinSize { get; set; }
		public virtual int? MaxSize { get; set; }
		public virtual Column Column { get; set; }
		public virtual bool Lookup { get; set; }
	}

	public class FusionLog : Entity
	{
		public virtual string MessageType { get; set; }
		public virtual Guid BusRef { get; set; }
		public virtual DateTime? LastGeneratedDate { get; set; }
		public virtual DateTime? LastProcessedDate { get; set; }
		public virtual string LastGeneratedXml { get; set; }
		public virtual string UserName { get; set; }
	}

	public class FusionMessage : Entity
	{

		public FusionMessage()
		{
			Elements = new List<FusionMessageElement>();
		}
		public virtual string Name { get; set; }
		public virtual string Description { get; set; }
		public virtual int Version { get; set; }
		public virtual bool AllowPublish { get; set; }
		public virtual bool AllowSubscribe { get; set; }
		public virtual bool Publish { get; set; }
		public virtual bool Subscribe { get; set; }
		public virtual IList<FusionMessageElement> Elements { get; protected set; }
	}

	public class FusionMessageElement : Entity
	{
		public virtual FusionMessage Message { get; set; }
		public virtual string NodeKey { get; set; }
		public virtual int Position { get; set; }
		public virtual DataType DataType { get; set; }
		public virtual bool Nillable { get; set; }
		public virtual int MinOccurs { get; set; }
		public virtual int MaxOccurs { get; set; }
		public virtual int? MinSize { get; set; }
		public virtual int? MaxSize { get; set; }
		public virtual bool Lookup { get; set; }
		public virtual FusionElement Element { get; set; }
	}
}