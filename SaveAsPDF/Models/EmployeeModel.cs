namespace SaveAsPDF.Models
{
    public class EmployeeModel
    {
        /// <summary>
        /// Employee ID
        /// </summary>
        public int Id { get; set; }

        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string EmailAddress { get; set; }
        /// <summary>
        /// First name and last name combined
        /// </summary>
        public string FullName => $"{FirstName} {LastName} <{EmailAddress}>";
    }
}
