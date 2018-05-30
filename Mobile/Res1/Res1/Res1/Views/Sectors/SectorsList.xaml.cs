using Res1.Models;
using Res1.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Res1.Views.Sectors
{
	[XamlCompilation(XamlCompilationOptions.Compile)]
	public partial class SectorsList : ContentPage
	{
		public SectorsList ()
		{
			InitializeComponent ();
            MyListView.ItemsSource = SectorServices.GetSectorss();
            MyListView.ItemSelected += SelectedItem_Method;
        }

        protected override bool OnBackButtonPressed()
        {
            return true;
        }


        public async void SelectedItem_Method(object sender,SelectedItemChangedEventArgs e)
        {
            if (e.SelectedItem == null)
                return;

            Sector S = e.SelectedItem as Sector;
            await Navigation.PushAsync(new SectorDetailed(S));

            MyListView.SelectedItem = null;

        }
    }
}