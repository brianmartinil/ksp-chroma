using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace KspChromaControl.DataDrains
{
    class DrainManager
    {
        private readonly List<IDataDrain> dataDrains = new List<IDataDrain>();
        private bool initialized = false;

        private static DrainManager instance;

        private DrainManager() { }

        public static DrainManager Instance
        {
            get {
                if (instance == null)
                {
                    instance = new DrainManager();
                }

                return instance;
            }
        }

        public void Init()
        {
            if (initialized)
            {
                return;
            }

            UnityEngine.Debug.Log("[CHROMA] Looking for valid drains");

            // This sort of convulted process is necessary because Unity's loader
            // doesn't let us just try to instantiate a drain and catch the exception
            // if it fails because the class' dependencies aren't available.

            dataDrains.AddRange(
                 Assembly.GetExecutingAssembly().GetTypes()
                    .Where(typeof(IDataDrain).IsAssignableFrom)
                    .Where(t => typeof(IDataDrain) != t)
                    .Select(t => (IDataDrain)Activator.CreateInstance(t))
                    .Where(d => d.Available())
            );

            UnityEngine.Debug.Log("[CHROMA] Found " + dataDrains.Count + " drains");


            initialized = true;
        }

        public void Send(ColorSchemes.ColorScheme scheme)
        {
            this.dataDrains.ForEach(drain => drain.Send(scheme));
        }
    }
}
