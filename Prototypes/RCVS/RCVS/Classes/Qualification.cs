using System;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RCVS.Classes
{
	[ModelBinder(typeof(DefaultModelBinder))]
	public class Qualification : BaseClass
	{	
		public string Name { get; set; }
		public DateTime ObtainedDate { get; set; }
		public string AwardingBody { get; set; }

		//public override object CreateModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
		//{
		//	HttpRequestBase request = controllerContext.HttpContext.Request;

		//	return new Qualification
		//	{
		//		Name = ConvertEmptyStringToNull(request.Form.Get("Name")),
		//		ObtainedDate = Convert.ToDateTime(request.Form.Get("ObtainedDate")),
		//		AwardingBody = ConvertEmptyStringToNull(request.Form.Get("AwardingBody")),	
		//	};
		//}

		//public override object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
		//{
		//	if (bindingContext.ModelType.IsInterface)
		//	{
		//		Type desiredType = Type.GetType(
		//				EncryptionService.Decrypt(
		//						(string)bindingContext.ValueProvider.GetValue("AssemblyQualifiedName").ConvertTo(typeof(string))));
		//		bindingContext.ModelMetadata = ModelMetadataProviders.Current.GetMetadataForType(null, desiredType);
		//	}

		//	return base.BindModel(controllerContext, bindingContext);
		//}

	//	public object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
	//	{
	//		throw new NotImplementedException();
	//	}
	}

}