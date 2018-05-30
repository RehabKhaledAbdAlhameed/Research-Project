using Res1.Models;
using Res1.Services;
using Res1.Views.Industories;
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
	public partial class SectorDetailed : ContentPage
	{
		public SectorDetailed (Sector sector)
		{
			InitializeComponent ();

            BindingContext = sector;
            List<Industry> industries = IndustryServices.GetIndustries();
            industries = (from i in industries
                          where i.Sec_id == sector.Sec_id
                          select i).ToList();
            MyIndustries.ItemsSource = industries;
            MyIndustries.ItemSelected += SelectedItem_Method;

        }

        public async void SelectedItem_Method(object sender, SelectedItemChangedEventArgs e)
        {
            if (e.SelectedItem == null)
                return;

            Industry I = e.SelectedItem as Industry;
            await Navigation.PushAsync(new IndustryDetailed(I));

            MyIndustries.SelectedItem = null;

        }
    }
}