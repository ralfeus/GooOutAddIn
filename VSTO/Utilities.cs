using System;
using System.Drawing;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Regexp = System.Text.RegularExpressions;
using Microsoft.Win32;
using System.Reflection;
using System.Windows.Forms;

namespace R.GoogleOutlookSync
{
    internal static class Utilities
    {
        //private static string tempPhotoPath = AppDomain.CurrentDomain.BaseDirectory + "\\TempOutlookContactPhoto.jpg";
        //private static string tempPhotoPath = Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData) + @"\" + System.Windows.Forms.Application.ProductName + @"\TempOutlookContactPhoto.jpg";
        private static string tempPhotoPath = Path.GetTempPath() + @"\TempOutlookContactPhoto.jpg";
        private static NotifyIcon notifyIcon = new NotifyIcon {
            Icon = SystemIcons.Application,
            Visible = true
        };

        public static byte[] BitmapToBytes(Bitmap bitmap)
        {
            //bitmap
            MemoryStream stream = new MemoryStream();
            bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Bmp);
            return stream.ToArray();
        }

        //public static bool HasPhoto(Contact googleContact)
        //{
        //    if (googleContact.PhotoEtag == null)
        //        return false;
        //    return true;
        //}
        public static bool HasPhoto(Outlook.ContactItem outlookContact)
        {
            return outlookContact.HasPicture;
        }

        //public static bool SaveGooglePhoto(Syncronizer sync, Contact googleContact, Image image)
        //{
        //    if (googleContact.ContactEntry.PhotoUri == null)
        //        throw new Exception("Must reload contact from google.");

        //    try
        //    {
        //        WebClient client = new WebClient();
        //        client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.ContactsRequest.Service.QueryClientLoginToken());
        //        client.Headers.Add(HttpRequestHeader.ContentType, "image/*");
        //        Bitmap pic = new Bitmap(image);
        //        Stream s = client.OpenWrite(googleContact.ContactEntry.PhotoUri.AbsoluteUri, "PUT");
        //        byte[] bytes = BitmapToBytes(pic);

        //        s.Write(bytes, 0, bytes.Length);
        //        s.Flush();
        //        s.Close();
        //        s.Dispose();
        //        client.Dispose();
        //        pic.Dispose();
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //    return true;
        //}
        //public static Image GetGooglePhoto(Syncronizer sync, Contact googleContact)
        //{
        //    if (!HasPhoto(googleContact))
        //        return null;

        //    try
        //    {
        //        WebClient client = new WebClient();
        //        client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.ContactsRequest.Service.QueryClientLoginToken());
        //        Stream stream = client.OpenRead(googleContact.PhotoUri.AbsoluteUri);
        //        BinaryReader reader = new BinaryReader(stream);
        //        Image image = Image.FromStream(stream);
        //        reader.Close();
        //        stream.Close();
        //        stream.Dispose();
        //        client.Dispose();

        //        return image;
        //    }
        //    catch
        //    {
        //        return null;
        //    }
        //}

        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, string fullImagePath)
        {
            try
            {
                outlookContact.AddPicture(fullImagePath);
                //outlookContact.Save();
                return true;
            }
            catch
            {
                return false;
            }
        }
        public static bool SetOutlookPhoto(Outlook.ContactItem outlookContact, Image image)
        {
            try
            {
                image.Save(tempPhotoPath);
                return SetOutlookPhoto(outlookContact, tempPhotoPath);
            }
            catch
            {
                return false;
            }
        }
        public static Image GetOutlookPhoto(Outlook.ContactItem outlookContact)
        {
            if (!HasPhoto(outlookContact))
                return null;

            try
            {
                foreach (Outlook.Attachment a in outlookContact.Attachments)
                {
                    // CH Fixed this to Contains, due to outlook picture that looks like "ContactPicture_138382.jpg"
                    if (a.DisplayName.ToUpper().Contains("CONTACTPICTURE") || a.DisplayName.ToUpper().Contains("CONTACTPHOTO"))
                    {

                        //TODO: Check why always the first added picture is returned
                        //If you add another picture, still the old picture is saved to tempPhotoPath
                        a.SaveAsFile(tempPhotoPath); 

                        using (Image img = Image.FromFile(tempPhotoPath))
                        {
                            return new Bitmap(img);
                        }
                    }
                }
                return null;
            }
            catch
            {
                // There's an error here... If Outlook says it has a contact photo, and we can't get it, Something's broken.

                return null;
            }
        }

        public static Image CropImageGoogleFormat(Image original)
        {
            // crop image to a square in the center

            int width, height, diff;
            Point p;
            Rectangle r;

            if (original.Height == original.Width)
                return original;
            if (original.Height > original.Width)
            {
                // tall image
                width = original.Width;
                height = width;

                diff = original.Height - height;
                p = new Point(0, diff / 2);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
            else
            {
                // flat image
                height = original.Height;
                width = height;

                diff = original.Width - width;
                p = new Point(diff / 2, 0);
                r = new Rectangle(p, new Size(width, height));

                return CropImage(original, r);
            }
        }
        public static Image CropImage(Image original, Rectangle cropArea)
        {
            Bitmap bmpImage = new Bitmap(original);
            Bitmap bmpCrop = bmpImage.Clone(cropArea, bmpImage.PixelFormat);
            return (Image)(bmpCrop);
        }

        public static void DeleteTempPhoto()
        {
            try
            {
                if (File.Exists(tempPhotoPath))
                    File.Delete(tempPhotoPath);
            }
            catch { }
        }

        //public static bool ContainsGroup(Syncronizer sync, Contact googleContact, string groupName)
        //{
        //    Group group = sync.GetGoogleGroupByName(groupName);
        //    if (group == null)
        //        return false;
        //    return ContainsGroup(googleContact, group);
        //}
        //public static bool ContainsGroup(Contact googleContact, Group group)
        //{
        //    foreach (GroupMembership m in googleContact.GroupMembership)
        //    {
        //        if (m.HRef == group.GroupEntry.Id.AbsoluteUri)
        //            return true;
        //    }
        //    return false;
        //}
        //public static bool ContainsGroup(Outlook.ContactItem outlookContact, string group)
        //{
        //    if (outlookContact.Categories == null)
        //        return false;

        //    return outlookContact.Categories.Contains(group);
        //}

        //public static Collection<Group> GetGoogleGroups(Syncronizer sync, Contact googleContact)
        //{
        //    int c = googleContact.GroupMembership.Count;
        //    Collection<Group> groups = new Collection<Group>();
        //    string id;
        //    Group group;
        //    for (int i = 0; i < c; i++)
        //    {
        //        id = googleContact.GroupMembership[i].HRef;
        //        group = sync.GetGoogleGroupById(id);

        //        groups.Add(group);
        //    }
        //    return groups;
        //}
        //public static void AddGoogleGroup(Contact googleContact, Group group)
        //{
        //    if (ContainsGroup(googleContact, group))
        //        return;

        //    GroupMembership m = new GroupMembership();
        //    m.HRef = group.GroupEntry.Id.AbsoluteUri;
        //    googleContact.GroupMembership.Add(m);
        //}
        //public static void RemoveGoogleGroup(Contact googleContact, Group group)
        //{
        //    if (!ContainsGroup(googleContact, group))
        //        return;

        //    // TODO: broken. removes group membership but does not remove contact
        //    // from group in the end.

        //    // look for id
        //    GroupMembership mem;
        //    for (int i = 0; i < googleContact.GroupMembership.Count; i++)
        //    {
        //        mem = googleContact.GroupMembership[i];
        //        if (mem.HRef == group.GroupEntry.Id.AbsoluteUri)
        //        {
        //            googleContact.GroupMembership.Remove(mem);
        //            return;
        //        }
        //    }
        //    throw new Exception("Did not find group");
        //}

        //public static string[] GetOutlookGroups(string outlookContactCategories)
        //{
        //    if (outlookContactCategories == null)
        //        return new string[] { };

        //    char[] listseparator = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator.ToCharArray();
        //    string[] categories = outlookContactCategories.Split(listseparator);
        //    for (int i = 0; i < categories.Length; i++)
        //    {
        //        categories[i] = categories[i].Trim();
        //    }
        //    return categories;
        //}
        //public static void AddOutlookGroup(Outlook.ContactItem outlookContact, string group)
        //{
        //    if (ContainsGroup(outlookContact, group))
        //        return;

        //    // append
        //    if (outlookContact.Categories == null)
        //        outlookContact.Categories = "";
        //    if (outlookContact.Categories != "")
        //        outlookContact.Categories += ", " + group;
        //    else
        //        outlookContact.Categories += group;
        //}
        //public static void RemoveOutlookGroup(Outlook.ContactItem outlookContact, string group)
        //{
        //    if (!ContainsGroup(outlookContact, group))
        //        return;

        //    outlookContact.Categories = outlookContact.Categories.Replace(", " + group, "");
        //    outlookContact.Categories = outlookContact.Categories.Replace(group, "");
        //}

        //ToDo: Workaround to save google Content is also not working, beause of error when closing the StreamWriter
        //public static bool SaveGoogleNoteContent(Syncronizer sync, Google.Documents.Document updated, Google.Documents.Document googleNote)
        //{

        //    if (updated.DocumentEntry.EditUri == null || googleNote.MediaSource == null)
        //        throw new Exception("Must reload note from google.");

        //    StreamWriter writer = null;
        //    StreamReader reader = null;
        //    WebClient client = null;
        //    try
        //    {
        //        client = new WebClient();
        //        client.Headers.Add(HttpRequestHeader.Authorization, "GoogleLogin auth=" + sync.DocumentsRequest.Service.QueryClientLoginToken());
        //        client.Headers.Add(HttpRequestHeader.ContentType, googleNote.MediaSource.ContentType);
        //        Stream s = client.OpenWrite(updated.DocumentEntry.EditUri.ToString(), "PUT");
        //        writer = new StreamWriter(s);
        //        reader = new StreamReader(googleNote.MediaSource.GetDataStream());
        //        string body = reader.ReadToEnd();
        //        writer.Write(body);
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //    finally
        //    {
        //        if (client != null)
        //            client.Dispose();
        //        if (writer != null)
        //            writer.Close(); //This throws an exception 400 (Ung�ltige Anforderung)
        //        if (reader != null)
        //            reader.Close();

        //    }

        //    return true;
        //}

        internal static Regexp.Regex SMTPAddressPattern = new Regexp.Regex(@"[\w_\.]+@([\w-]+\.?)+");

        //internal void InitializeFormControls(Form form)
        //{
        //    ResourceManager temp = new ResourceManager("R.GoogleOutlookSync.Properties.Resources", typeof(Properties.Resources).Assembly);
        //    var resourceSet = temp.GetResourceSet(Thread.CurrentThread.CurrentUICulture, false, false);
        //    Dictionary<string, string> controlsValues = new Dictionary<string, string>();
        //    List<string> controlsNames = new List<string>();
        //    foreach (DictionaryEntry resource in resourceSet)
        //    {
        //        if (resource.Key.ToString().Contains(form.Name))
        //        {
        //            var tokens = resource.Value.ToString().Split('_');
        //            controlsValues.Add(tokens[1], tokens[2]);
        //            controlsNames.Add(tokens[1]);
        //        }
        //    }
        //    if (controlsValues.Count != 0)
        //    {
        //        foreach (Control control in form.Controls)
        //        {
        //            if (controlsNames.Contains(control.Name))
        //                control.GetType().InvokeMember(controlsValues[control.Name], System.Reflection.BindingFlags.SetProperty, null, control, new object[] {
        //        }
        //    }
        //}

        public static void TryDeleteFile(string filePath)
        {
            try
            {
                File.Delete(filePath);
            }
            catch (Exception)
            {
                File.Move(filePath, Path.GetDirectoryName(filePath) + "\\" + Path.GetFileName(filePath) + ".toDelete");
                var runOnceRegKey = Registry.CurrentUser.OpenSubKey(VSTO.Properties.Settings.Default.Registry_Key_Autorun, true);
                runOnceRegKey.SetValue("Delete file " + filePath.Replace('\\', ' '), "cmd.exe /c del " + filePath);
            }
        }

        public static ProcessorArchitecture GetAssemblyArchitecture()
        {
            return Assembly.GetExecutingAssembly().GetName().ProcessorArchitecture;
        }

        internal static string GetOSArchitecture()
        {
            if (Environment.GetEnvironmentVariable("ProgramFiles(x86)") != null)
                return "AMD64";
            else
                return "X86";
        }

        internal static string GetStringResource(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(resourceName))
                {
                    TextReader tr = new StreamReader(asm.GetManifestResourceStream(name));
                    //Debug.Assert(tr != null); 
                    string resource = tr.ReadToEnd();

                    tr.Close();
                    return resource;
                }
            }
            return null;
        }

        /// <summary>
        /// Returns a value from application's registry key
        /// </summary>
        /// <param name="valueName">A name of the value</param>
        /// <returns>Value. Type is object so consumer should cast it afterwards</returns>
        internal static object GetRegistryValue(string valueName)
        {
            var regKey = Registry.CurrentUser.OpenSubKey(VSTO.Properties.Settings.Default.ApplicationRegistryKey);
            return regKey.GetValue(valueName);
        }

        internal static void Notify(string message, ToolTipIcon toolTipIcon) {
            notifyIcon.ShowBalloonTip(20000, "OutlookGoogleSync notification", message, toolTipIcon);
        }

        /// <summary>
        /// Sets a value in application's registry key
        /// Type of value is inferred from the value itself
        /// </summary>
        /// <param name="valueName">Name of the value</param>
        /// <param name="value">Value to be set</param>
        internal static void SetRegistryValue(string valueName, object value)
        {
            var regKey = Registry.CurrentUser.OpenSubKey(VSTO.Properties.Settings.Default.ApplicationRegistryKey, true);
            try {
                regKey.SetValue(valueName, value);
            } catch (NullReferenceException e) {
                throw new Exception(String.Format("Couldn't open registry key {0}", VSTO.Properties.Settings.Default.ApplicationRegistryKey), e);
            }
        }
    }
}
