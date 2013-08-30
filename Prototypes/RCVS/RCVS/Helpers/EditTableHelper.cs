using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Web.Mvc;

namespace RCVS.Helpers
{
	public static class EditTableHelper
	{

		public static string EditTable(this HtmlHelper helper, string name, IList items, IDictionary<string, object> attributes)
		{
			if (items == null || items.Count == 0 || string.IsNullOrEmpty(name))
			{
				return string.Empty;
			}

			return BuildTable(name, items, attributes);
		}

		private static string BuildTable(string name, IList items, IDictionary<string, object> attributes)
		{
			StringBuilder sb = new StringBuilder();
			BuildTableHeader(sb, items[0].GetType());

			var i = 0;
			foreach (var item in items)
			{
				BuildTableRow(name, sb, item, i);
				i++;
			}

			TagBuilder builder = new TagBuilder("table");
			builder.MergeAttributes(attributes);
			builder.MergeAttribute("name", name);
			builder.MergeAttribute("id", name);

			foreach (var item in items)
			{
				BuildTableModelBinder(sb, item);
			}

			builder.InnerHtml = sb.ToString();
			return builder.ToString(TagRenderMode.Normal);
		}

		private static void BuildTableRow(string name, StringBuilder sb, object obj, int RowNo)
		{
			Type objType = obj.GetType();
			sb.AppendLine("\t<tr>");

			foreach (var property in objType.GetProperties())
			{
				sb.AppendFormat("\t\t<td>{0}</td>\n", property.GetValue(obj, null));
			}
			sb.AppendFormat("\t\t<td><input class='editbutton' type={0}button{0} onclick='editActivity({0}{1}{0},{2})' value={0}...{0} id={0}button1{0}></td>\n", '"', name, RowNo);
			sb.AppendFormat("\t\t<td><input class='deletebutton' type={0}button{0} onclick='deleteActivity({0}{1}{0},{2})' value={0}X{0} id={0}button1{0}></td>\n", '"', name, RowNo);
			sb.AppendLine("\t</tr>");
		}

		private static void BuildTableHeader(StringBuilder sb, Type p)
		{
			sb.AppendLine("\t<tr>");
			foreach (var property in p.GetProperties())
			{
				sb.AppendFormat("\t\t<th>{0}</th>\n", property.Name);
			}
			sb.AppendFormat("\t\t<th>Edit</th>\n");
			sb.AppendFormat("\t\t<th>Delete</th>\n");
			sb.AppendLine("\t</tr>");
		}

		private static void BuildTableModelBinder(StringBuilder sb, object obj)
		{

//			Type objType = obj.GetType();
//			var i = 0;
//			foreach (var property in objType.GetProperties())
//			{
//				sb.AppendFormat("\t\t<td>{0}</td>\n", property.GetValue(obj, null));
////				<input id="EmploymentHistory_0__City" name="EmploymentHistory[0].City" type="hidden" value="Aberdare" />

//			}

		}


	}
}