using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Data;

namespace SMLetsExchangeConnectorSettingsUI.Validation
{
    //combobox validation rule
    //https://stackoverflow.com/questions/15342143/how-can-i-not-allow-wpf-combobox-empty
    public class SelectionValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, System.Globalization.CultureInfo cultureInfo)
        {
            return value == null
                ? new ValidationResult(false, "Please select a template")
                : new ValidationResult(true, null);
        }
    }

    //change visibility based on SOME kind of selection
    //http://geekswithblogs.net/thibbard/archive/2007/12/10/wpf---showhide-element-based-on-checkbox.checked.aspx
    public class BooleanToHiddenVisibility : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            Visibility rv = Visibility.Visible;
            try
            {
                var x = bool.Parse(value.ToString());
                if (x)
                {
                    rv = Visibility.Visible;
                }
                else
                {
                    rv = Visibility.Collapsed;
                }
            }
            catch (Exception)
            {
            }
            return rv;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return value;
        }
    }
}
