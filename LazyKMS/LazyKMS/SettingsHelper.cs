using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LazyKMS
{
    class SettingsHelper
    {
        public static Settings settings
        {
            get
            {
                if (_settings == null)
                {
                    _settings = new Settings();
                }
                return _settings;
            }
        }
        private static Settings _settings;
    }
}
