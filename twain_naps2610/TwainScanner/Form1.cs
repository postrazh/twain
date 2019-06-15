using NAPS2.Config;
using NAPS2.Scan;
using NAPS2.Scan.Images;
using NAPS2.Util;
using NAPS2.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace TwainScanner
{
    public partial class Form1 : FormBase
    {
        #region Dependencies

        private readonly IScanDriverFactory driverFactory;

        private readonly AppConfigManager appConfigManager;
        private readonly IProfileManager profileManager;
        private readonly IScanPerformer scanPerformer;
        private readonly WinFormsExportHelper exportHelper;
        private readonly NotificationManager notify;

        #endregion

        #region State Fields

        private readonly ScannedImageList imageList = new ScannedImageList();

        #endregion

        public Form1(IScanDriverFactory driverFactory, AppConfigManager appConfigManager, IProfileManager profileManager, IScanPerformer scanPerformer,
            WinFormsExportHelper exportHelper, NotificationManager notify)
        {
            this.driverFactory = driverFactory;

            this.appConfigManager = appConfigManager;
            this.profileManager = profileManager;
            this.scanPerformer = scanPerformer;
            this.exportHelper = exportHelper;
            this.notify = notify;

            InitializeComponent();

            notify.ParentForm = this;
        }

        /// <summary>
        /// Constructs a receiver for scanned images.
        /// This keeps images from the same source together, even if multiple sources are providing images at the same time.
        /// </summary>
        /// <returns></returns>
        public Action<ScannedImage> ReceiveScannedImage()
        {
            ScannedImage last = null;
            return scannedImage =>
            {
                SafeInvoke(() =>
                {
                    lock (imageList)
                    {
                        // Default to the end of the list
                        int index = imageList.Images.Count;
                        // Use the index after the last image from the same source (if it exists)
                        if (last != null)
                        {
                            int lastIndex = imageList.Images.IndexOf(last);
                            if (lastIndex != -1)
                            {
                                index = lastIndex + 1;
                            }
                        }
                        imageList.Images.Insert(index, scannedImage);
                        scannedImage.MovedTo(index);
                        last = scannedImage;
                    }
                });
            };
        }

        private async void SavePDF(List<ScannedImage> images)
        {
            if (await exportHelper.SavePDF(images, notify))
            {
                if (appConfigManager.Config.DeleteAfterSaving)
                {
                    SafeInvoke(() =>
                    {
                        imageList.Delete(imageList.Images.IndiciesOf(images));
                    });
                }
            }
        }

        ScanProfile defaultScanProfile = null;
        ScanParams defaultScanParams = null;

        private async void btnScan_Click(object sender, EventArgs e)
        {
            // prepare the scan profile
            if (defaultScanProfile == null)
            {
                defaultScanProfile = new ScanProfile { Version = ScanProfile.CURRENT_VERSION };
                defaultScanProfile.DriverName = "twain";

                defaultScanParams = new ScanParams();

                var driver = driverFactory.Create(defaultScanProfile.DriverName);
                driver.ScanProfile = defaultScanProfile;
                driver.ScanParams = defaultScanParams;
                var deviceList = driver.GetDeviceList();
                if (!deviceList.Any())
                {
                    MessageBox.Show("There is no connected device!");
                    return;
                }
                defaultScanProfile.Device = deviceList[0];
            }

            // perfor scan
            do
            {
                await scanPerformer.PerformScan(defaultScanProfile, defaultScanParams, this, notify, ReceiveScannedImage());
            } while (MessageBox.Show("Would you like to continue?", "Question", MessageBoxButtons.YesNo) == DialogResult.Yes);

            SavePDF(imageList.Images);
            imageList.Delete(Enumerable.Range(0, imageList.Images.Count));
        }
    }
}
