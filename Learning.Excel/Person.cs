namespace Learning.Excel
{
    public class Person
    {
        public double ID { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Email { get; set; }

        public string Gender { get; set; }
    }

    public enum Gender
    {
        Male,
        Female
    }

    public static class Constants
    {
        public static string MALE { get; } = "Male";
        public static string FEMALE { get; } = "Female";
    }
}



    
