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

        private readonly StringWrapper stringWrapper;
        private readonly AppConfigManager appConfigManager;
        private readonly RecoveryManager recoveryManager;
        private readonly OcrManager ocrManager;
        private readonly IProfileManager profileManager;
        private readonly IScanPerformer scanPerformer;
        private readonly IScannedImagePrinter scannedImagePrinter;
        private readonly ChangeTracker changeTracker;
        private readonly StillImage stillImage;
        private readonly IOperationFactory operationFactory;
        private readonly IUserConfigManager userConfigManager;
        private readonly KeyboardShortcutManager ksm;
        private readonly ThumbnailRenderer thumbnailRenderer;
        private readonly WinFormsExportHelper exportHelper;
        private readonly ScannedImageRenderer scannedImageRenderer;
        private readonly NotificationManager notify;
        private readonly CultureInitializer cultureInitializer;
        private readonly IWorkerServiceFactory workerServiceFactory;
        private readonly IOperationProgress operationProgress;
        private readonly UpdateChecker updateChecker;

        #endregion

        #region State Fields

        private readonly ScannedImageList imageList = new ScannedImageList();
        private readonly AutoResetEvent renderThumbnailsWaitHandle = new AutoResetEvent(false);

        #endregion

        public FDesktop(StringWrapper stringWrapper, AppConfigManager appConfigManager, RecoveryManager recoveryManager, OcrManager ocrManager, IProfileManager profileManager, IScanPerformer scanPerformer, IScannedImagePrinter scannedImagePrinter, ChangeTracker changeTracker, StillImage stillImage, IOperationFactory operationFactory, IUserConfigManager userConfigManager, KeyboardShortcutManager ksm, ThumbnailRenderer thumbnailRenderer, WinFormsExportHelper exportHelper, ScannedImageRenderer scannedImageRenderer, NotificationManager notify, CultureInitializer cultureInitializer, IWorkerServiceFactory workerServiceFactory, IOperationProgress operationProgress, UpdateChecker updateChecker)
        {
            this.stringWrapper = stringWrapper;
            this.appConfigManager = appConfigManager;
            this.recoveryManager = recoveryManager;
            this.ocrManager = ocrManager;
            this.profileManager = profileManager;
            this.scanPerformer = scanPerformer;
            this.scannedImagePrinter = scannedImagePrinter;
            this.changeTracker = changeTracker;
            this.stillImage = stillImage;
            this.operationFactory = operationFactory;
            this.userConfigManager = userConfigManager;
            this.ksm = ksm;
            this.thumbnailRenderer = thumbnailRenderer;
            this.exportHelper = exportHelper;
            this.scannedImageRenderer = scannedImageRenderer;
            this.notify = notify;
            this.cultureInitializer = cultureInitializer;
            this.workerServiceFactory = workerServiceFactory;
            this.operationProgress = operationProgress;
            this.updateChecker = updateChecker;
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

            UpdateScanButton();

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


        private void UpdateScanButton()
        {
            const int staticButtonCount = 2;

            // Clean up the dropdown
            while (tsScan.DropDownItems.Count > staticButtonCount)
            {
                tsScan.DropDownItems.RemoveAt(0);
            }

            // Populate the dropdown
            var defaultProfile = profileManager.DefaultProfile;
            int i = 1;
            foreach (var profile in profileManager.Profiles)
            {
                var item = new ToolStripMenuItem
                {
                    Text = profile.DisplayName.Replace("&", "&&"),
                    Image = profile == defaultProfile ? Icons.accept_small : null,
                    ImageScaling = ToolStripItemImageScaling.None
                };
                item.Click += async (sender, args) =>
                {
                    profileManager.DefaultProfile = profile;
                    profileManager.Save();

                    UpdateScanButton();

                    await scanPerformer.PerformScan(profile, new ScanParams(), this, notify, ReceiveScannedImage());
                    Activate();
                };
                tsScan.DropDownItems.Insert(tsScan.DropDownItems.Count - staticButtonCount, item);

                i++;
            }

            if (profileManager.Profiles.Any())
            {
                tsScan.DropDownItems.Insert(tsScan.DropDownItems.Count - staticButtonCount, new ToolStripSeparator());
            }
        }

        private void ShowProfilesForm()
        {
            var form = FormFactory.Create<FProfiles>();
            form.ImageCallback = ReceiveScannedImage();
            form.ShowDialog();
            UpdateScanButton();
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

    }
}
