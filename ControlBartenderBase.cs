﻿using BarTender;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace Ssapj.ControlBartenderViaCom
{
    public class ControlBartenderBase : IDisposable
    {
        private Application _bartenderApplication;
        private int _batenderProcessId = -1;

        protected Application BartenderApplication
        {
            get { return this._bartenderApplication; }
        }

        protected int BartenderProcessId
        {
            get { return this._batenderProcessId; }
        }

        public async Task StartBartenderAsync()
        {
            await Task.Run(() =>
            {
                this._bartenderApplication = new BarTender.Application()
                {
                    VisibleWindows = BarTender.BtVisibleWindows.btNone
                };

                this._batenderProcessId = this._bartenderApplication.ProcessId;
            });
        }

        protected void StartBartender()
        {
            //reset
            Marshal.FinalReleaseComObject(this._bartenderApplication);
            this._bartenderApplication = null;
            this._batenderProcessId = -1;

            this._bartenderApplication = new Application()
            {
                VisibleWindows = BarTender.BtVisibleWindows.btNone
            };

            this._batenderProcessId = this._bartenderApplication.ProcessId;
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            //for managed resources
            if (disposing)
            { }

            //for unmanaged resources
            if (this._bartenderApplication != null)
            {
                try
                {
                    using (var hProcess = Process.GetProcessById(this._batenderProcessId))
                    {
                        this._bartenderApplication.Quit(BtSaveOptions.btDoNotSaveChanges);
                    }
                }
                finally
                {
                    Marshal.FinalReleaseComObject(this._bartenderApplication);
                }
            }
        }

        ~ControlBartenderBase()
        {
            this.Dispose(false);
        }

    }

}


