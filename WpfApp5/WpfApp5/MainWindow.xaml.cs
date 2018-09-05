using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Exchange.WebServices;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Exchange.WebServices.Autodiscover;


namespace WpfApp5
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;
            Uri redirectionUri = new Uri(redirectionUrl);
            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials. 
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void TextBox_TextChanged_1(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //https://docs.microsoft.com/ru-ru/exchange/client-developer/exchange-web-services/get-started-with-ews-managed-api-client-applications

            ExchangeService ExServ = new ExchangeService(ExchangeVersion.Exchange2013_SP1);

            ExServ.UseDefaultCredentials = true;

            ExServ.Credentials = new WebCredentials(TextLoginBox.Text,PassBox.Password);
            ExServ.AutodiscoverUrl(TextLoginBox.Text, RedirectionUrlValidationCallback);

            EmailMessage EmailMessage = new EmailMessage(ExServ);
            EmailMessage.Subject = TextTitleBox.Text;
            EmailMessage.Body = TextMsgBox.Text;
            EmailMessage.ToRecipients.Add(TextReciveBox.Text);
            EmailMessage.Save();




            try
            {
                EmailMessage.SendAndSaveCopy();
                MessageBox.Show("Message send OK");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + ex.ToString());
            }

        }
    }
}

