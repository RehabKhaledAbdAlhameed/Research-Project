using Res1.Views.Sectors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Res1.Views
{
	[XamlCompilation(XamlCompilationOptions.Compile)]
	public partial class Login : ContentPage
	{
		public Login ()
		{
			InitializeComponent ();
		}
        private async  void Login_Clicked(object sender, EventArgs e)
        {
            if (Name.Text.Equals("Ahmed") && Pass.Text.Equals("123"))
            {

                await Navigation.PushAsync(new SectorsList());
            }
        }
        protected override bool OnBackButtonPressed()
        {
            return true;
        }
    }
}