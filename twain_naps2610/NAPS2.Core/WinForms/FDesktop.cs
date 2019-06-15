#region Usings

using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NAPS2.Config;
using NAPS2.ImportExport;
using NAPS2.Lang;
using NAPS2.Lang.Resources;
using NAPS2.Logging;
using NAPS2.Ocr;
using NAPS2.Operation;
using NAPS2.Platform;
using NAPS2.Recovery;
using NAPS2.Scan;
using NAPS2.Scan.Exceptions;
using NAPS2.Scan.Images;
using NAPS2.Scan.Wia;
using NAPS2.Scan.Wia.Native;
using NAPS2.Update;
using NAPS2.Util;
using NAPS2.Worker;

#endregion

namespace NAPS2.WinForms
{
    public partial class FDesktop : FormBase
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

        public FDesktop(IScanDriverFactory driverFactory, AppConfigManager appConfigManager, IProfileManager profileManager, IScanPerformer scanPerformer, 
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


        private async Task ScanDefault()
        {
            if (profileManager.DefaultProfile != null)
            {
                await scanPerformer.PerformScan(profileManager.DefaultProfile, new ScanParams(), this, notify, ReceiveScannedImage());
                Activate();
            }
            else if (profileManager.Profiles.Count == 0)
            {
                await ScanWithNewProfile();
            }
            else
            {
                ShowProfilesForm();
            }
        }

        private async Task ScanWithNewProfile()
        {
            var editSettingsForm = FormFactory.Create<FEditProfile>();
            editSettingsForm.ScanProfile = appConfigManager.Config.DefaultProfileSettings ?? new ScanProfile { Version = ScanProfile.CURRENT_VERSION };
            editSettingsForm.ShowDialog();
            if (!editSettingsForm.Result)
            {
                return;
            }
            profileManager.Profiles.Add(editSettingsForm.ScanProfile);
            profileManager.DefaultProfile = editSettingsForm.ScanProfile;
            profileManager.Save();

            await scanPerformer.PerformScan(editSettingsForm.ScanProfile, new ScanParams(), this, notify, ReceiveScannedImage());
            Activate();
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

        private void ShowProfilesForm()
        {
            var form = FormFactory.Create<FProfiles>();
            form.ImageCallback = ReceiveScannedImage();
            form.ShowDialog();
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


        private async void tsScan_ButtonClick(object sender, EventArgs e)
        {
            await ScanDefault();
        }

        private async void tsNewProfile_Click(object sender, EventArgs e)
        {
            await ScanWithNewProfile();
        }


        private void tsProfiles_Click(object sender, EventArgs e)
        {
            ShowProfilesForm();
        }

        
        private void tsdSavePDF_ButtonClick(object sender, EventArgs e)
        {
           SavePDF(imageList.Images);
        }

        ScanProfile defaultScanProfile = null;
        ScanParams defaultScanParams = null;

        private async void btnConnect_Click(object sender, EventArgs e)
        {
            
     
        }

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
