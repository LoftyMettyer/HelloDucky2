using System;
using System.Web.Mvc;
using RCVS.Classes;

public class MBTestBinder : IModelBinder
 {
  

	public object BindModel(ControllerContext controllerContext, ModelBindingContext bindingContext)
	{
		Qualification instance = new Qualification();
		instance.Name = controllerContext.HttpContext.Request["Name"];
		return instance;
		//throw new NotImplementedException();
	}
 }