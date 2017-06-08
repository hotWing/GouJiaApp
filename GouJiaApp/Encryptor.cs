using System;

namespace GouJiaApp
{
    class Encryptor
    {
        public static bool isValid() 
        {
            //DateTime now = DateTime.Today;
            //DateTime then = new DateTime(2015,5,15);

            //return DateTime.Compare(now,then)<=0;
            return true;
        }

        public static bool isExpired()
        {
            //DateTime now = DateTime.Today;
            //DateTime then = new DateTime(2016, 6, 20);

            //return DateTime.Compare(now, then) > 0;
            return false;
        }
    }
}
