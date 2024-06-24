namespace PR50.Models
{
    public class Owner
    {
        public string Surname { get; set; }
        public string Name { get; set; }
        public string Lastname { get; set; }
        public int NumberRoom { get; set; }
        public string Image { get; set; }
        public Owner(string surname, string name, string lastname, int numberRoom, string image)
        {
            Surname = surname;
            Name = name;
            Lastname = lastname;
            NumberRoom = numberRoom;
            Image = image;
        }
    }
}
