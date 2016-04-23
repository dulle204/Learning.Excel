namespace Learning.Excel
{
    public class Person
    {
        public int ID { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string Email { get; set; }

        public Gender Genfeder { get; set; }
    }
}



    public enum Gender
    {
        Male, 
        Female
    }
