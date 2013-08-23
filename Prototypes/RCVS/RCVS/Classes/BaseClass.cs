using System;
using System.Web.Mvc;

namespace RCVS.Classes
{
	//public abstract class BaseClass : IModelBinder
	public abstract class BaseClass : DefaultModelBinder
	{
		public long ID { get; set; }
		//// Abstract methods
		//public abstract object CreateModel(ControllerContext controllerContext, ModelBindingContext bindingContext);

		//// Methods
		//protected virtual object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
		//{
		//	// Create model
		//	object model = this.CreateModel(controllerContext, bindingContext);

		//	// Iterate through model properties
		//	foreach (ModelMetadata property in bindingContext.PropertyMetadata.Values)
		//	{
		//		// Get property value
		//		property.Model = bindingContext.ModelType.GetProperty(property.PropertyName).GetValue(model, null);
		//		// Get property validator
		//		foreach (ModelValidator validator in property.GetValidators(controllerContext))
		//		{
		//			// Validate property
		//			foreach (ModelValidationResult result in validator.Validate(model))
		//			{
		//				// Add error message into model state
		//				bindingContext.ModelState.AddModelError(property.PropertyName + "." + result.MemberName, result.Message);
		//			}
		//		}
		//	}

		//	return model;
		//}

		//// IModelBinder members
		//object IModelBinder.BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
		//{
		//	return BindModel(controllerContext, bindingContext);
		//}

		//protected string ConvertEmptyStringToNull(string value)
		//{
		//	return (string.IsNullOrEmpty(value) ? null : value);
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
	}
}