using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LazyKMS
{
    class FormProvider
    {
        public static MainForm mainForm
        {
            get
            {
                if (_mainFrom == null)
                {
                    _mainFrom = new MainForm();
                }
                return _mainFrom;
            }
        }
        public static ProcessForm procForm
        {
            get
            {
                if (_procFrom == null)
                {
                    _procFrom = new ProcessForm();
                }
                return _procFrom;
            }
        }
        private static MainForm _mainFrom;
        private static ProcessForm _procFrom;
    }
}
