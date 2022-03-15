
using System;
using JPMorrow.Revit.Labor;

namespace JPMorrow.Bluebeam.FlexConduit
{
    public class BluebeamFlexConduitTotal
    {
        public static string OneHalfFlexPipeName { get => "Steel FMC - 6 Ft. Whip - 1/2\""; }
        public static string ThreeQuarterFlexPipeName { get => "Steel FMC - 6 Ft. Whip - 3/4\""; }
        public static string OneHalfFlexConnectorName { get => "Connector - Steel FMC - 1/2\""; }
        public static string ThreeQuarterFlexConnectorName { get => "Connector - Steel FMC - 3/4\""; }
        public int PipeQty { get; private set; } = 0;
        public int ConnectorQty { get; private set; } = 0;

        public BluebeamFlexConduitTotal(int box_cnt)
        {
            // get 70% of box count
            int adjusted_box_cnt = (int)Math.Round(box_cnt * 0.70);

            PipeQty = adjusted_box_cnt;
            ConnectorQty = adjusted_box_cnt * 2;
        }
    }

    public static class BluebeamFlexConduit
    {
        public static BluebeamFlexConduitTotal GetTotalFromBoxCount(int box_count)
        {
            return new BluebeamFlexConduitTotal(box_count);
        }
    }
}