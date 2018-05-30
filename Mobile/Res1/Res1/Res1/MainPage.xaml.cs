using Res1.Views;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xamarin.Forms;

namespace Res1
{
	public partial class MainPage : ContentPage
	{
		public MainPage()
		{
			InitializeComponent();
            Thread.Sleep(1000);
            Navigation.PushAsync(new Login());
		}

       
    }
}
