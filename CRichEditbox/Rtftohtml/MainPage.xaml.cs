using RtfPipe;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.Storage;
using Windows.Storage.Streams;
using Windows.UI.Text;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace Rtftohtml
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
        }
        
        private void btngetcontent_Click(object sender, RoutedEventArgs e)
        {
            string rtf = "";
            Richbox.Document.GetText(TextGetOptions.FormatRtf, out rtf);
            string html = Rtf.ToHtml(rtf);
            webshow.NavigateToString(html);
        }

        private async void Page_Loaded(object sender, RoutedEventArgs e)
        {
            StorageFolder folder = ApplicationData.Current.LocalFolder;
            StorageFile file = await StorageFile.GetFileFromApplicationUriAsync(new Uri("ms-appx:///Assets/test1.rtf"));
            IRandomAccessStream randAccStream = await file.OpenAsync(FileAccessMode.Read);
            Richbox.Document.LoadFromStream(TextSetOptions.FormatRtf, randAccStream);
        }
    }
}
