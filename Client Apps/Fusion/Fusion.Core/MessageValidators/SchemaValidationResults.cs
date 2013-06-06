// --------------------------------------------------------------------------------------------------------------------
// <copyright file="SchemaValidationResults.cs" company="Advanced Health and Care Limited">
//   Copyright © 2011 Advanced Health and Care Limited - All Rights Reserved.
// </copyright>
// <summary>
//   Implements the schema validation results class
// </summary>
// --------------------------------------------------------------------------------------------------------------------

namespace Fusion.Core.MessageValidators
{
    using System.Text;
    using System;
    
    /// <summary>
    /// Schema validation results. 
    /// </summary>
    public class SchemaValidationResults
    {
        /// <summary>
        /// Gets a value indicating whether the validation found errors.
        /// </summary>
        /// <value>
        /// true if the validation found errors, false if not.
        /// </value>
        public bool HasErrors
        {
            get
            {
                return this.ValidationErrors == null ? false : ValidationErrors.Length > 0;
            }          
        }

        /// <summary>
        /// Gets or sets the validation errors.
        /// </summary>
        /// <value>
        /// The validation errors.
        /// </value>
        public string[] ValidationErrors
        {
            get;
            set;
        }

        /// <summary>
        /// Gets the validation error string.  This is a concatenated list of all the errors present in the ValidationErrors list
        /// </summary>
        /// <value>
        /// The validation error string.
        /// </value>
        public string ValidationErrorString
        {
            get
            {
                if (ValidationErrors == null)
                {
                    return String.Empty;
                }

                StringBuilder sb = new StringBuilder();
                foreach (string s in ValidationErrors)
                {
                    sb.AppendLine(s);
                }

                return sb.ToString();
            }
        }
    }
}
