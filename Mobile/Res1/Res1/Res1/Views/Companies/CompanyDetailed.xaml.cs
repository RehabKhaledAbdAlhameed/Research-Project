using Res1.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;
using Xamarin.Forms.Xaml;

namespace Res1.Views.Companies
{
	[XamlCompilation(XamlCompilationOptions.Compile)]
	public partial class CompanyDetailed : ContentPage
	{
		public CompanyDetailed (Company company)
		{
			InitializeComponent ();
            BindingContext = company;

		}
	}
}