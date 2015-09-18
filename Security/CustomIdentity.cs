using System;
using System.Security.Principal;

namespace Security
{
    /// <summary>
    /// Represents the Identity of a User. 
    /// Stores the details of a User. 
    ///  </summary>
    [Serializable]
    public class CustomIdentity : GenericIdentity
    {
        #region Properties

        /// <summary>
        /// Gets or sets the Id of the User
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// Gets or sets the First Name of the User
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// Gets or sets the Last Name of the User
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// Gets or sets the Email of the User
        /// </summary>
        public string Email { get; set; }

        /// <summary>
        /// Indicates whether the User is an Administrator
        /// </summary>
        public bool IsSuperUser { get; set; }

        #endregion

        # region Constructors

        /// <summary>
        /// Initializes a new instance of the Security.CustomIdentity
        /// class with the passed parameter.
        /// </summary>
        /// <param name="identity">The object from which to construct the new instance of Security.CustomIdentity.</param>
        public CustomIdentity(GenericIdentity identity)
            : base(identity)
        {
        }

        /// <summary>
        /// Initializes a new instance of the Security.CustomIdentity
        /// class with the passed parameter.
        /// </summary>
        /// <param name="name">The name of the user on whose behalf the code is running.</param>
        /// <exception cref="System.ArgumentNullException">The name parameter is null.</exception>
        public CustomIdentity(string name)
            : base(name)
        {
        }

        /// <summary>
        /// Initializes a new instance of the Security.CustomIdentity
        /// class with the passed parameters.
        /// </summary>
        /// <param name="name">The name of the user on whose behalf the code is running.</param>
        /// <param name="type">The type of authentication used to identify the user</param>
        /// <exception cref="System.ArgumentNullException">The name parameter is null.-or- The type parameter is null.</exception>
        public CustomIdentity(string name, string type)
            : base(name, type)
        {
        }

        #endregion
    }
}
