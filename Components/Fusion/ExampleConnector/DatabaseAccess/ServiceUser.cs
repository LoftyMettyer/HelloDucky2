// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ServiceUser.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the POCO for Service Users
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Connector1.DatabaseAccess
{
    /// <summary>
    /// A service user. 
    /// </summary>
    public class ServiceUser
    {
        /// <summary>
        /// Gets or sets the forename.
        /// </summary>
        /// <value>
        /// The forename.
        /// </value>
        public string Forename
        {
            get;
            set;
        }

        /// <summary>
        /// Gets or sets the surname.
        /// </summary>
        /// <value>
        /// The surname.
        /// </value>
        public string Surname
        {
            get;
            set;
        }    
    }
}
