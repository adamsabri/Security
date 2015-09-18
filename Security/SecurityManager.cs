using System.Security;
using System.Security.Principal;
using System.Threading;

namespace Security
{
    public class SecurityManager
    {
        public static CustomIdentity CurrentIdentity
        {
            get
            {
                var cp = Thread.CurrentPrincipal;
                if (cp == null) throw new SecurityException("Non Authenticated User");
                return CurrentPrincipal.Identity as CustomIdentity;
            }
        }

        public static GenericPrincipal CurrentPrincipal
        {
            get
            {
                var cp = Thread.CurrentPrincipal as GenericPrincipal;
                if (cp == null) throw new SecurityException("Non Authenticated User");
                return cp;
            }
            private set { Thread.CurrentPrincipal = value; }
        }

        public static void SetCurrentPrincipal(IIdentity identity, string[] roles)
        {
            CurrentPrincipal = new GenericPrincipal(identity, roles);
        }
    }
}
