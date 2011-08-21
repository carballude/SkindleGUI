using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml;
using System.Reflection;
using System.Threading;
using System.Linq;
using SkindleGUI.Model;

namespace SkindleGUI
{
    public partial class frmMain : Form
    {
        /************************************************
        *                                               *
        *   Instance Variables                          *
        *                                               *
        ************************************************/

        private Regex fileRegex;
        private Regex bookRegex;
        private WebClient wb; // used to fetch book images via HTTP
        private int lastIndex = -1;
        private string skindle_path = Path.GetTempPath() + "skindle.exe"; // store skindle.exe in temporary directory

        /************************************************
        *                                               *
        *   Constructor                                 *
        *                                               *
        ************************************************/
        public frmMain()
        {
            InitializeComponent();

            // get assembly resource for skindle.exe -- credit to http://www.cs.nyu.edu/~vs667/articles/embed_executable_tutorial/
            Assembly currentAssembly = Assembly.GetExecutingAssembly();
            Stream skindleEXEStream = currentAssembly.GetManifestResourceStream("SkindleGUI.skindle.exe");

            // open output file for skindle.exe
            FileInfo fileInfoOutputFile = new FileInfo(skindle_path);
            FileStream outputStream = fileInfoOutputFile.OpenWrite();

            // read from embedded resource and write to output file
            const int size = 4096;
            byte[] bytes = new byte[4096];
            int numBytes;
            while ((numBytes = skindleEXEStream.Read(bytes, 0, size)) > 0)
            {
                outputStream.Write(bytes, 0, numBytes);
            }
            outputStream.Close();
            skindleEXEStream.Close();

            // setup regular expression object for matching filenames
            fileRegex = new Regex(@"(?<folder>[a-z]:\\(?:[^\\/:*?""<>|\r\n]+\\)*)(?<file>[^\\/:*?""<>|\r\n.]*)\.(?<extension>.+)$", RegexOptions.IgnoreCase);
            bookRegex = new Regex(@"(?<drive>[a-z]:)\\(?<folder>(?:[^\\/:*?""<>|\r\n]+\\)*)(?<book>[^\\/:*?""<>|\r\n.]*)_EBOK\.(?<extension>.+)$", RegexOptions.IgnoreCase);

            // initialize webclient object
            wb = new WebClient();

            // determine path to "My Kindle Content" within My Documents folder and set it as default input directory
            string myDocsPath = Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments);
            string kindlePath = myDocsPath + "\\My Kindle Content\\";
            txtInput.Text = kindlePath;            
        }

        private void UpdateText(Book book)
        {
            int index = lstBooks.Items.IndexOf(book);
            lstBooks.BeginUpdate();
            lstBooks.Items[index] = book;
            lstBooks.EndUpdate();
        }

        private void SetBookName(object book)
        {
            Book b = (Book)book;
            this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Looking title for book: " + b.FileName); });
            try
            {
                var content = new WebClient().DownloadString("http://www.amazon.com/asd/dp/" + b.FileName.Split(new char[] { '_' })[0] + "/");
                var title = content.Split(new string[] { "<meta name=\"title\" content=\"Amazon.com: " }, StringSplitOptions.RemoveEmptyEntries)[1].Split(new string[] { ": Kindle Store\" />" }, StringSplitOptions.RemoveEmptyEntries)[0];
                b.Name = title;
                this.Invoke((MethodInvoker)delegate() { UpdateText(b); });
                this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Title found for book: " + b.FileName); });
            }
            catch (WebException) { this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Title not found for book " + b.FileName); }); }
        }

        private void SetBookInfo(Book book)
        {
            Match match = fileRegex.Match(book.FilePath);
            string extension = match.Groups["extension"].Value;
            if (extension == "azw" || extension == "tpz")
            {
                book.FileName = match.Groups["file"] + "." + extension;
                lstBooks.Items.Add(book);
                new Thread(new ParameterizedThreadStart(SetBookName)).Start(book);
            }
        }

        private void loadBookNames()
        {
            //lstBooks.Items.Clear();
            Directory.GetFiles(txtInput.Text).Select(x => new Book() { FilePath = x }).ToList().ForEach(x => SetBookInfo(x));
        }

        /************************************************
        *                                               *
        *   event handlers for Main tab                 *
        *                                               *
        ************************************************/

        private void lstBooks_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lastIndex != lstBooks.SelectedIndex)
            {                
                new Thread(new ParameterizedThreadStart(getCoverImage)).Start(lstBooks.SelectedItem);                
                lastIndex = lstBooks.SelectedIndex;
            }
        }

        private void getCoverImage(object bookObject)
        {
            if (bookObject == null) return;
            // get the book ID (in form like "B003O86FMW") from the filename (books are stored as "B003O86FMW_EBOK.azw")
            //string book = bookRegex.Match((string)bookFileName).Groups["book"].Value;
            //GetBookNames(book);
            Book book = (Book)bookObject;
            string bookId = book.FileName.Split(new char[] { '_' })[0];

            // construct URL where book is located
            // this is used both to read from local cache or to get via HTTP
            string url = "http://ecx.images-amazon.com/images/P/" + bookId + ".jpg";
            
            // define variable for holding the book's image
            Image bookImage = null;

            bookImage = GetImageFromLocalCache(bookId);
            // if book image still null at this point then it means the local cache failed to return image
            // so we will try to read from web URL via HTTP
            if (bookImage == null && chkUseInternet.Checked)
                    bookImage = GetImageFromInternet(bookId);             
            SetBookCover(bookImage);
        }

        private Image GetImageFromLocalCache(string book)
        {
            Image bookImage = null;
            string url = "http://ecx.images-amazon.com/images/P/" + book + ".jpg";
            // first try getting the book image from the local disk cache stored in the Kindle for PC local application data
            string kindleCachePath = Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData) + "\\Amazon\\Kindle For PC\\Cache\\res\\";
            if (File.Exists(kindleCachePath + "cache.xml"))
            {
                // have to parse XML file which stores the mapping between book image URL -> local cache filename
                XmlTextReader reader = new XmlTextReader(kindleCachePath + "cache.xml");
                bool inResource = false;
                bool inID = false;
                bool inPath = false;
                bool breakLoop = false;
                string filename = null;
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            if (inResource && reader.Name == "path")
                                inPath = true;
                            else if (inResource && reader.Name == "id")
                                inID = true;
                            else if (reader.Name == "resource")
                                inResource = true;
                            break;
                        case XmlNodeType.Text:
                            if (inID && reader.Value != url)
                                inResource = false; // don't read filename if not the right book
                            else if (inPath)
                            {
                                filename = reader.Value;
                                breakLoop = true;
                            }
                            break;
                        case XmlNodeType.EndElement:
                            if (reader.Name == "resource")
                                inResource = false;
                            else if (reader.Name == "id")
                                inID = false;
                            else if (reader.Name == "path")
                                inPath = false;
                            break;
                    }
                    if (breakLoop)
                        break;
                }
                if (File.Exists(filename))
                    bookImage = Image.FromFile(filename);
            }
            return bookImage;
        }

        private void SetStatusBarText(string text)
        {
            statusBarLabel.Text = text;
        }

        private Image GetImageFromInternet(string bookid)
        {
            this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Looking cover for book: " + bookid); });
            Image bookImage = null;
            string url = "http://ecx.images-amazon.com/images/P/" + bookid + ".jpg";
            try
            {
                bookImage = Image.FromStream(new WebClient().OpenRead(url));
            }
            catch (WebException ex)
            {
                MessageBox.Show(ex.ToString(), "WebException Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Impossible to retreive cover for book: " + bookid); });
                picCover.Image = null;
                return null;
            }
            return bookImage;
        }

        private void SetBookCover(Image bookImage)
        {
            if (bookImage == null) return;
            Image resizedImage = bookImage.GetThumbnailImage(picCover.Width, picCover.Height, null, IntPtr.Zero);
            picCover.Image = resizedImage;
            this.Invoke((MethodInvoker)delegate() { SetStatusBarText("Cover found!"); });
        }

        private void btnBrowseOut_Click(object sender, EventArgs e)
        {
            string InputExt = fileRegex.Match(txtInput.Text + lstBooks.SelectedItem).Groups["extension"].Value;
            SaveFileDialog sd = new SaveFileDialog();
            if (InputExt.Equals("azw"))
                sd.Filter = "MOBI Files|*.mobi|All Files|*.*";
            else if (InputExt.Equals("tpz"))
                sd.Filter = "Amazon Topaz Files|*.tpz|All Files|*.*";
            else
                sd.Filter = "All Files|*.*";
           
            if (sd.ShowDialog() == DialogResult.OK)
                txtOutput.Text = sd.FileName;
        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            // clear results textbox since we're starting new job
            txtResults.Text = "";

            // confirm there is a selected book
            if (lstBooks.SelectedIndex == -1)
            { // -1 signifies no item selected
                MessageBox.Show("Must select a book filename from the list of books.", "Required Fields Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // confirm the selected book exists
            string input = txtInput.Text + ((Book)lstBooks.SelectedItem).FileName;
            if (!File.Exists(input))
            {
                MessageBox.Show("The selected book filename does not exist. Please restart the program.", "Required Fields Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // confirm they chose a destination output path
            if (txtOutput.Text == "")
            {
                MessageBox.Show("Must specify an output file path.", "Required Fields Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            
            // confirm that "skindle.exe" exists in the current directory            
            /*txtResults.Text += "Looking for skindle.exe at " + skindle_path + Environment.NewLine;*/
            if (!File.Exists(skindle_path))
            {
                MessageBox.Show("\"skindle.exe\" file not found. It must be in the same directory as this program's executable.", "Missing Skindle Program", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Now, all error checking is done.

            // Construct the run string for skindle.exe (that is, the command line parameters)
            string RunString = "";
            if (chkDecompress.Checked)
                RunString += "-d ";
            if (chkDump.Checked)
                RunString += "-v ";
            RunString += "-i \"" + input + "\" ";
            RunString += "-o \"" + txtOutput.Text + "\" ";
            if (txtInfo.Text != "")
                RunString += "-k \"" + txtInfo.Text + "\" ";
            if (txtPID.Text != "")
                RunString += "-p " + txtPID.Text + " ";

            txtResults.Text += skindle_path + " " + RunString;

            // make a Process object to launch skindle.exe
            Process process = new Process();
            process.StartInfo.UseShellExecute = false;
            process.StartInfo.RedirectStandardError = true;
            process.StartInfo.CreateNoWindow = true;
            process.StartInfo.FileName = skindle_path;
            process.StartInfo.Arguments = RunString;
            // begin the launch, catching any exceptions
            try
            {
                process.Start();
                // collect and output results
                string output = process.StandardError.ReadToEnd();                
                txtResults.Text += Environment.NewLine + Environment.NewLine + output;
                process.WaitForExit();
                txtResults.Text += Environment.NewLine + "Done.";
                process.Close();
            }
            catch (Exception ex)
            {
                txtResults.Text += Environment.NewLine + Environment.NewLine + "Error: " + ex.ToString();
            }
        }

         /************************************************
         *                                               *
         *   event handlers for Optional Settings tab    *
         *                                               *
         ************************************************/

        private void btnBrowseInputFolder_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Amazon Kindle Files|*.azw|Amazon Topaz Files|*.tpz|All Files|*.*";
            if (od.ShowDialog() == DialogResult.OK)
            {
                string folder = fileRegex.Match(od.FileName).Groups["folder"].Value;
                txtInput.Text = folder;
            }
            loadBookNames();
        }

        private void btnBrowseInfo_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Amazon Kindle.info Files|*.info|All Files|*.*";
            if (od.ShowDialog() == DialogResult.OK)
                txtInfo.Text = od.FileName;
        }

         /****************************************************
         *                                                   *
         *   FormClosing handler to clean up temporary EXE   *
         *                                                   *
         ****************************************************/

        private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
        {
            // delete temporary "skindle.exe" file
            File.Delete(skindle_path);
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            // load listbox with files in default input directory
            loadBookNames();
        }

        

        
    }
}
