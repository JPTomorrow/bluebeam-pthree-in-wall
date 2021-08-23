

namespace JPMorrow.Pdf.Bluebeam.P3
{
    public class BluebeamP3BoxConfig
    {
    }

    public class BluebeamP3Box
    {
        public BluebeamP3BoxConfig Config { get; private set; }
        public BluebeamP3Box(BluebeamP3BoxConfig config)
        {
            Config = config;
        }
    }

    public class BluebeamConduitPackage
    {

    } 
}