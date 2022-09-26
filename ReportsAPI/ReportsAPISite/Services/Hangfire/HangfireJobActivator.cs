using Hangfire;
using StructureMap;
using System;

namespace ReportsAPISite.Services.Hangfire
{
    public class HangfireJobActivator : JobActivator
    {
        private readonly IContainer _container;

        public HangfireJobActivator(IContainer container)
        {
            _container = container;
        }

        public override object ActivateJob(Type type)
        {
            return _container.GetInstance(type);
        }
    }
}