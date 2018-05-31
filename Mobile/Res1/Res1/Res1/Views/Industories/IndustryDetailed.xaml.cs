using Res1.Models;
using Res1.Services;
using Res1.Views.Companies;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Res1.Views.Industories
{
	[XamlCompilation(XamlCompilationOptions.Compile)]
	public partial class IndustryDetailed : ContentPage
	{
		public IndustryDetailed (Industry Industry)
		{
			InitializeComponent ();
            BindingContext = Industry;

            List<Company> Companies = CompanyServices.GetCompanies();
            Companies = (from c in Companies
                          where c.Ind_id== Industry.Ind_id
                          select c).ToList();
            MyCompanies.ItemsSource = Companies;
            MyCompanies.ItemSelected += SelectedItem_Method;
        }
        public async void SelectedItem_Method(object sender, SelectedItemChangedEventArgs e)
        {
            if (e.SelectedItem == null)
                return;

            Company C = e.SelectedItem as Company;
            await Navigation.PushAsync(new CompanyDetailed(C));

            MyCompanies.SelectedItem = null;

        }
    }
}