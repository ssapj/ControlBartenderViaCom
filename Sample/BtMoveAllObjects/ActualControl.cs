using BarTender;
using Ssapj.ControlBartenderViaCom;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace BtMoveAllObjects
{
    class ActualControl : ControlBartenderBase
    {
        private double _x = 1.5;
        private double _y = 1.5;
        private BtUnits _btUnits = BtUnits.btUnitsCentimeters;

        public void ControlBtwFile(string btwfilepath)
        {
            try
            {
                using (var hProcess = Process.GetProcessById(this.BartenderProcessId))
                {
                }
            }
            catch (ArgumentException)
            {
                Console.WriteLine("Restart BarTender");
                this.RestartBartender();
            }

            var btFormat = this.BartenderApplication.Formats.Open(btwfilepath);

            btFormat.MeasurementUnits = this._btUnits;

            var targetDesignObjects = Enumerable.Range(1, btFormat.Objects.Count)
                .Select(i => btFormat.Objects.Item(i))
                .Where(xs => xs.Type != BarTender.BtObjectType.btObjectError
                    && xs.Type != BarTender.BtObjectType.btObjectRFID
                    && xs.Type != BarTender.BtObjectType.btObjectMagneticStripe
                    && xs.Type != BarTender.BtObjectType.btObjectShape)
                    ;

            foreach (var item in targetDesignObjects)
            {
                item.X += this._x;
                item.Y += this._y;
            }

            btFormat.Close(BarTender.BtSaveOptions.btSaveChanges);

            Marshal.FinalReleaseComObject(btFormat);
            btFormat = null;
        }

    }


}
