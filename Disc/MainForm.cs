using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Windows.Forms;
using IMAPI2.Interop;
using IMAPI2.MediaItem;
using BurnMedia.Model;

namespace BurnMedia
{

    public partial class MainForm : Form
    {
        private const string ClientName = "BurnMedia";

        private double? freeSpaceOnDisk;
        private double? totalSpaceOnDisk;

        private bool _isBurning;
        private bool _isFormatting;
        private IMAPI_BURN_VERIFICATION_LEVEL _verificationLevel = IMAPI_BURN_VERIFICATION_LEVEL.IMAPI_BURN_VERIFICATION_NONE;
        private bool _closeMedia;
        private bool _ejectMedia;

        private Dictionares _dictionares;
        private BurnData _burnData = new BurnData();

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            _dictionares = new Dictionares();
            MsftDiscMaster2 discMaster = null;
            try
            {
                discMaster = new MsftDiscMaster2();
                if (!discMaster.IsSupportedEnvironment)
                {
                    return;
                }
                foreach (string uniqueRecorderId in discMaster)
                {
                    var tempDiscRecorder = new MsftDiscRecorder2();
                    tempDiscRecorder.InitializeDiscRecorder(uniqueRecorderId);
                    devicesComboBox.Items.Add(tempDiscRecorder);
                }
                if (devicesComboBox.Items.Count > 0)
                {
                    devicesComboBox.SelectedIndex = 0;
                }
            }
            catch (COMException ex)
            {
                MessageBox.Show(ex.Message,
                    string.Format("Error:{0} - Please install IMAPI2", ex.ErrorCode),
                    MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            finally
            {
                if (discMaster != null)
                {
                    Marshal.ReleaseComObject(discMaster);
                }
            }
            InitVolumeLabel();
        }

        private void InitVolumeLabel()
        {
            var now = DateTime.Now;
            textBoxLabel.Text = now.Year + "_" + now.Month + "_" + now.Day;
            labelStatusText.Text = string.Empty;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (MsftDiscRecorder2 tempDiscRecorder in devicesComboBox.Items)
            {
                if (tempDiscRecorder != null)
                {
                    Marshal.ReleaseComObject(tempDiscRecorder);
                }
            }
        }

        private void DevicesComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (devicesComboBox.SelectedIndex == -1)
            {
                return;
            }

            var discRecorder = (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];
            IDiscFormat2Data discFormatData = null;
            try
            {
                discFormatData = new MsftDiscFormat2Data();
                if (!discFormatData.IsRecorderSupported(discRecorder))
                {
                    MessageBox.Show("Recorder not supported", 
                        ClientName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                totalSpaceOnDisk = (2048 * discFormatData.TotalSectorsOnMedia) / 1048576;
                TotalSpace.Text = "Total Space: " + Convert.ToInt32(totalSpaceOnDisk).ToString() + " MB";
                freeSpaceOnDisk = freeSpaceOnDisk == null ? 
                    (2048 * discFormatData.FreeSectorsOnMedia) / 1048576 
                    : freeSpaceOnDisk + (2048 * discFormatData.FreeSectorsOnMedia) / 1048576;
                FreeSpace.Text = "Free Space: " + Convert.ToInt32(freeSpaceOnDisk) + " MB";

                var supportedMediaTypes = new StringBuilder();
                foreach (IMAPI_PROFILE_TYPE profileType in discRecorder.SupportedProfiles)
                {
                    var profileName = _dictionares.ProfileTypeDictionary[profileType];
                    if (string.IsNullOrEmpty(profileName))
                    {
                        continue;
                    }
                    if (supportedMediaTypes.Length > 0)
                    {
                        supportedMediaTypes.Append(", ");
                    }
                    supportedMediaTypes.Append(profileName);
                }

            }
            catch (COMException)
            { }
            finally
            {
                if (discFormatData != null)
                {
                    Marshal.ReleaseComObject(discFormatData);
                }
            }
        }

        private void DevicesComboBox_Format(object sender, ListControlConvertEventArgs e)
        {
            var discRecorder2 = (IDiscRecorder2)e.ListItem;
            var devicePaths = string.Empty;
            var volumePath = (string)discRecorder2.VolumePathNames.GetValue(0);
            foreach (string volPath in discRecorder2.VolumePathNames)
            {
                if (!string.IsNullOrEmpty(devicePaths))
                {
                    devicePaths += ",";
                }
                devicePaths += volumePath;
            }
            e.Value = string.Format("{0} [{1}]", devicePaths, discRecorder2.ProductId);
        }

        private void ButtonDetectMedia_Click(object sender, EventArgs e)
        {
            if (devicesComboBox.SelectedIndex == -1)
            {
                return;
            }
            var discRecorder = (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];
            MsftFileSystemImage fileSystemImage = null;
            MsftDiscFormat2Data discFormatData = null;
            try
            {
                discFormatData = new MsftDiscFormat2Data();
                if (!discFormatData.IsCurrentMediaSupported(discRecorder))
                {
                    labelMediaType.Text = "Media not supported!";
                    return;
                }
                else
                {
                    discFormatData.Recorder = discRecorder;
                    var mediaType = discFormatData.CurrentPhysicalMediaType;

                    freeSpaceOnDisk = (2048 * discFormatData.FreeSectorsOnMedia) / 1048576;
                    FreeSpace.Text = "Free Space: " + Convert.ToInt32(freeSpaceOnDisk).ToString()+" MB";
                    TotalSpace.Text = "Total Space: " + Convert.ToInt32((2048 * discFormatData.TotalSectorsOnMedia) / 1048576).ToString() +" MB";

                    labelMediaType.Text = _dictionares.TypeDictionary[mediaType];
                    fileSystemImage = new MsftFileSystemImage();
                    fileSystemImage.ChooseImageDefaultsForMediaType(mediaType);
                    if (!discFormatData.MediaHeuristicallyBlank)
                    {
                        fileSystemImage.MultisessionInterfaces = discFormatData.MultisessionInterfaces;
                        fileSystemImage.ImportFileSystem();
                    }
                }
            }
            catch (COMException exception)
            {
                MessageBox.Show(this, exception.Message, "Detect Media Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (discFormatData != null)
                {
                    Marshal.ReleaseComObject(discFormatData);
                }

                if (fileSystemImage != null)
                {
                    Marshal.ReleaseComObject(fileSystemImage);
                }
            }
        }

        private void ButtonBurn_Click(object sender, EventArgs e)
        {
            if (devicesComboBox.SelectedIndex == -1 || freeSpaceOnDisk < 0)
            {
                return;
            }

            if (_isBurning)
            {
                buttonBurn.Enabled = false;
                backgroundBurnWorker.CancelAsync();
            }
            else
            {
                _isBurning = true;
                _closeMedia = true;
                _ejectMedia = checkBoxEject.Checked;

                EnableBurnUi(false);

                var discRecorder = (IDiscRecorder2)devicesComboBox.Items[devicesComboBox.SelectedIndex];
                _burnData.UniqueRecorderId = discRecorder.ActiveDiscRecorder;

                backgroundBurnWorker.RunWorkerAsync(_burnData);
                BurnerNotifyIcon.BalloonTipText = "Burning start";
                BurnerNotifyIcon.ShowBalloonTip(30000);
            }
        }

        private void BackgroundBurnWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            MsftDiscRecorder2 discRecorder = null;
            MsftDiscFormat2Data discFormatData = null;
            try
            {
                //
                // Create and initialize the IDiscRecorder2 object
                //
                discRecorder = new MsftDiscRecorder2();
                var burnData = (BurnData)e.Argument;
                discRecorder.InitializeDiscRecorder(burnData.UniqueRecorderId);
                //
                // Create and initialize the IDiscFormat2Data
                //
                discFormatData = new MsftDiscFormat2Data
                    {
                        Recorder = discRecorder,
                        ClientName = ClientName,
                        ForceMediaToBeClosed = _closeMedia
                    };
                //
                // Set the verification level
                //
                var burnVerification = (IBurnVerification)discFormatData;
                burnVerification.BurnVerificationLevel = _verificationLevel;
                //
                // Check if media is blank, (for RW media)
                //    
                object[] multisessionInterfaces = null;
                if (!discFormatData.MediaHeuristicallyBlank)
                {
                    multisessionInterfaces = discFormatData.MultisessionInterfaces;
                }
                //
                // Create the file system
                //
                if (!CreateMediaFileSystem(discRecorder, multisessionInterfaces, out IStream fileSystem))
                {
                    e.Result = -1;
                    return;
                }
                //
                // add the Update event handler
                //
                discFormatData.Update += DiscFormatData_Update;
                //
                // Write the data here
                //
                try
                {
                    discFormatData.Write(fileSystem);
                    e.Result = 0;
                }
                catch (COMException ex)
                {
                    e.Result = ex.ErrorCode;
                    MessageBox.Show(ex.Message, "IDiscFormat2Data.Write failed",
                        MessageBoxButtons.OK, MessageBoxIcon.Stop);
                }
                finally
                {
                    if (fileSystem != null)
                    {
                        Marshal.FinalReleaseComObject(fileSystem);
                    }
                }
                //
                // remove the Update event handler
                //
                discFormatData.Update -= DiscFormatData_Update;

                if (_ejectMedia)
                {
                    discRecorder.EjectMedia();
                }
            }
            catch (COMException exception)
            {
                //
                // If anything happens during the format, show the message
                //
                MessageBox.Show(exception.Message);
                e.Result = exception.ErrorCode;
            }
            finally
            {
                if (discRecorder != null)
                {
                    Marshal.ReleaseComObject(discRecorder);
                }

                if (discFormatData != null)
                {
                    Marshal.ReleaseComObject(discFormatData);
                }
            }
        }

        private void DiscFormatData_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender,
            [In, MarshalAs(UnmanagedType.IDispatch)] object progress)
        {
            if (backgroundBurnWorker.CancellationPending)
            {
                var format2Data = (IDiscFormat2Data)sender;
                format2Data.CancelWrite();
                return;
            }

            var eventArgs = (IDiscFormat2DataEventArgs)progress;
            _burnData.Task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_WRITING;
            // IDiscFormat2DataEventArgs Interface
            _burnData.ElapsedTime = eventArgs.ElapsedTime;
            _burnData.RemainingTime = eventArgs.RemainingTime;
            _burnData.TotalTime = eventArgs.TotalTime;

            // IWriteEngine2EventArgs Interface
            _burnData.CurrentAction = eventArgs.CurrentAction;
            _burnData.StartLba = eventArgs.StartLba;
            _burnData.SectorCount = eventArgs.SectorCount;
            _burnData.LastReadLba = eventArgs.LastReadLba;
            _burnData.LastWrittenLba = eventArgs.LastWrittenLba;
            _burnData.TotalSystemBuffer = eventArgs.TotalSystemBuffer;
            _burnData.UsedSystemBuffer = eventArgs.UsedSystemBuffer;
            _burnData.FreeSystemBuffer = eventArgs.FreeSystemBuffer;

            backgroundBurnWorker.ReportProgress(0, _burnData);
        }

        private void BackgroundBurnWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            labelStatusText.Text = (int)e.Result == 0 ? "Finished Burning Disc!" : "Error Burning Disc!";
            BurnerNotifyIcon.BalloonTipText = "Burning end";
            BurnerNotifyIcon.ShowBalloonTip(30000);
            statusProgressBar.Value = 0;
            _isBurning = false;
            EnableBurnUi(true);
            buttonBurn.Enabled = true;
        }

        private void EnableBurnUi(bool enable)
        {
            buttonBurn.Text = enable ? "&Burn" : "&Cancel";
            buttonDetectMedia.Enabled = enable;

            devicesComboBox.Enabled = enable;
            listBoxFiles.Enabled = enable;

            buttonAddFiles.Enabled = enable;
            buttonAddFolders.Enabled = enable;
            buttonRemoveFiles.Enabled = enable;
            checkBoxEject.Enabled = enable;
            textBoxLabel.Enabled = enable;
        }

        private void BackgroundBurnWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            var burnData = (BurnData)e.UserState;

            if (burnData.Task == BURN_MEDIA_TASK.BURN_MEDIA_TASK_FILE_SYSTEM)
            {
                labelStatusText.Text = burnData.StatusMessage;
            }
            else if (burnData.Task == BURN_MEDIA_TASK.BURN_MEDIA_TASK_WRITING)
            {
                labelStatusText.Text = burnData.CurrentAction == IMAPI_FORMAT2_DATA_WRITE_ACTION.IMAPI_FORMAT2_DATA_WRITE_ACTION_WRITING_DATA ?
                    StatusWriting(burnData)
                    : _dictionares.StatusDictionary[burnData.CurrentAction];               
            }
        }

        private string StatusWriting(BurnData burnData)
        {
            long writtenSectors = burnData.LastWrittenLba - burnData.StartLba;

            if (writtenSectors > 0 && burnData.SectorCount > 0)
            {
                var percent = (int)((100 * writtenSectors) / burnData.SectorCount);
                statusProgressBar.Value = percent;
                return string.Format("Progress: {0}%", percent);
            }
            else
            {
                statusProgressBar.Value = 0;
                return "Progress 0%";
            }
        }

        private void EnableBurnButton()
        {
            buttonBurn.Enabled = (listBoxFiles.Items.Count > 0);
        }

        private bool CreateMediaFileSystem(IDiscRecorder2 discRecorder, object[] multisessionInterfaces, out IStream dataStream)
        {
            MsftFileSystemImage fileSystemImage = null;
            try
            {
                fileSystemImage = new MsftFileSystemImage();
                fileSystemImage.ChooseImageDefaults(discRecorder);
                fileSystemImage.FileSystemsToCreate = FsiFileSystems.FsiFileSystemJoliet | FsiFileSystems.FsiFileSystemISO9660;
                fileSystemImage.VolumeName = textBoxLabel.Text;

                fileSystemImage.Update += FileSystemImage_Update;

                if (multisessionInterfaces != null)
                {
                    fileSystemImage.MultisessionInterfaces = multisessionInterfaces;
                    fileSystemImage.ImportFileSystem();
                }

                var rootItem = fileSystemImage.Root;

                foreach (IMediaItem mediaItem in listBoxFiles.Items)
                {
                    if (backgroundBurnWorker.CancellationPending)
                        break;

                    mediaItem.AddToFileSystem(rootItem);
                }

                fileSystemImage.Update -= FileSystemImage_Update;

                if (backgroundBurnWorker.CancellationPending)
                {
                    dataStream = null;
                    return false;
                }

                dataStream = fileSystemImage.CreateResultImage().ImageStream;
            }
            catch (COMException exception)
            {
                MessageBox.Show(this, exception.Message, "Create File System Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                dataStream = null;
                return false;
            }
            finally
            {
                if (fileSystemImage != null)
                {
                    Marshal.ReleaseComObject(fileSystemImage);
                }
            }
	        return true;
        }

        private void FileSystemImage_Update([In, MarshalAs(UnmanagedType.IDispatch)] object sender,
            [In, MarshalAs(UnmanagedType.BStr)]string currentFile, [In] int copiedSectors, [In] int totalSectors)
        {
            var percentProgress = 0;
            if (copiedSectors > 0 && totalSectors > 0)
            {
                percentProgress = (copiedSectors * 100) / totalSectors;
            }

            if (!string.IsNullOrEmpty(currentFile))
            {
                var fileInfo = new FileInfo(currentFile);
                _burnData.StatusMessage = "Adding \"" + fileInfo.Name + "\" to image...";

                _burnData.Task = BURN_MEDIA_TASK.BURN_MEDIA_TASK_FILE_SYSTEM;
                backgroundBurnWorker.ReportProgress(percentProgress, _burnData);
            }

        }

        private void ButtonAddFiles_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                var fileItem = new FileItem(openFileDialog.FileName);
                if (freeSpaceOnDisk < 0)
                {
                    return;
                }
                freeSpaceOnDisk = freeSpaceOnDisk == null ? -fileItem.SizeOnDisc / 1048576 : freeSpaceOnDisk - fileItem.SizeOnDisc / 1048576;
                FreeSpace.Text = "Free Space: " + Convert.ToInt32(freeSpaceOnDisk) + " MB";

                listBoxFiles.Items.Add(fileItem);
                EnableBurnButton();
            }
        }

        private void ButtonAddFolders_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog.ShowDialog(this) != DialogResult.OK)
            {
                return;
            }
            if (freeSpaceOnDisk < 0)
            {
                return;
            }
            var directoryItem = new DirectoryItem(folderBrowserDialog.SelectedPath);

            freeSpaceOnDisk = freeSpaceOnDisk == null ? directoryItem.SizeOnDisc / 1048576 : freeSpaceOnDisk - directoryItem.SizeOnDisc / 1048576;
            FreeSpace.Text = "Free Space: " + Convert.ToInt32(freeSpaceOnDisk) + " MB";

            listBoxFiles.Items.Add(directoryItem);
            EnableBurnButton();
        }

        private void ButtonRemoveFiles_Click(object sender, EventArgs e)
        {
            var mediaItem = (IMediaItem)listBoxFiles.SelectedItem;
            if (mediaItem == null)
            {
                return;
            }

            if (MessageBox.Show("Are you sure you want to remove \"" + mediaItem + "\"?",
                "Remove item", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                freeSpaceOnDisk = freeSpaceOnDisk + mediaItem.SizeOnDisc / 1048576;
                FreeSpace.Text = "Free Space: " + Convert.ToInt32(freeSpaceOnDisk) + " MB";
                listBoxFiles.Items.Remove(mediaItem);
                EnableBurnButton();
            }
        }

        private void ListBoxFiles_SelectedIndexChanged(object sender, EventArgs e)
        {
            buttonRemoveFiles.Enabled = (listBoxFiles.SelectedIndex != -1);
        }

        private void ListBoxFiles_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index == -1) return;
            var mediaItem = (IMediaItem)listBoxFiles.Items[e.Index];
            if (mediaItem == null)
                return;

            e.DrawBackground();

            if ((e.State & DrawItemState.Focus) != 0)
                e.DrawFocusRectangle();

            if (mediaItem.FileIconImage != null)
                e.Graphics.DrawImage(mediaItem.FileIconImage, new Rectangle(4, e.Bounds.Y + 4, 16, 16));

            var rectF = new RectangleF(e.Bounds.X + 24, e.Bounds.Y,
                e.Bounds.Width - 24, e.Bounds.Height);

            var font = new Font(FontFamily.GenericSansSerif, 11);

            var stringFormat = new StringFormat
                {
                    LineAlignment = StringAlignment.Center,
                    Alignment = StringAlignment.Near,
                    Trimming = StringTrimming.EllipsisCharacter
                };

            e.Graphics.DrawString(mediaItem.ToString(), font, new SolidBrush(e.ForeColor),
                rectF, stringFormat);
        }

        private void TabControl_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (_isBurning || _isFormatting)
            {
                e.Cancel = true;
            }
        }

        private void MainForm_Resize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                this.ShowInTaskbar = false;
            }
        }

        private void BurnerNotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            this.ShowInTaskbar = true;
            WindowState = FormWindowState.Normal;
        }
    }
}
