using System;
using System.Security.Principal;

namespace Security
{
	/// <summary>
	/// Represents a Custom Principal
	/// </summary>
	[Serializable]
	public class CustomPrincipal: GenericPrincipal
	{
		# region Constructor
		/// <summary>
		/// Initializes a new instance of the GenericPrincipal class 
        /// from a CustomIdentity and an an array of role names 
		/// to which the user represented by that CustomIdentity belongs
		/// </summary>
        /// <param name="identity">An implementation of System.Security.Principal.IIdentity
        /// that represents any user.</param>
        /// <param name="roles">An array of role names to which the user represented by the identity parameter
        ///     belongs.</param>
        /// <exception cref="System.ArgumentNullException">The identity parameter is null.</exception>
        public CustomPrincipal(IIdentity identity, string[] roles)
            : base(identity, roles)
		{
		}
		#endregion
	}
}
