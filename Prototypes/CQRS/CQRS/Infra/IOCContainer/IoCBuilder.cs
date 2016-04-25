using Autofac;

namespace Infra.IOCContainer
{
	public static class IoCBuilder
	{
		private static ContainerBuilder containerBuilder;
		private static IContainer container;

		public static ContainerBuilder ContainerBuilder
		{
			get
			{
				if (containerBuilder == null)
				{
					containerBuilder = new ContainerBuilder();
				}
				return containerBuilder;
			}
		}

		public static IContainer Container
		{
			get
			{
				if (container == null)
				{
					container = ContainerBuilder.Build();
				}
				return container;
			}
		}
	}
}