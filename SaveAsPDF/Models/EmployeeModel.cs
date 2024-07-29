namespace SaveAsPDF.Models
{
    public class EmployeeModel
    {
        /// <summary>
        /// Employee ID
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// First name of the employee
        /// </summary>
        public string FirstName { get; set; }

        /// <summary>
        /// Last name of the employee
        /// </summary>
        public string LastName { get; set; }

        /// <summary>
        /// Email address of the employee
        /// </summary>
        public string EmailAddress { get; set; }

        /// <summary>
        /// Phone number of the employee
        /// </summary>
        public string PhoneNumber { get; set; }

        /// <summary>
        /// Indicates whether the employee is a project leader or not
        /// </summary>
        public bool IsLeader { get; set; }

        /// <summary>
        /// First name and last name combined with email address
        /// </summary>
        public string FullName => $"{FirstName} {LastName} <{EmailAddress}>";

        public EmployeeModel()
        {
            Id = 0;
            FirstName = string.Empty;
            LastName = string.Empty;
            EmailAddress = string.Empty;
            PhoneNumber = string.Empty;
            IsLeader = false;
        }
    }

}
