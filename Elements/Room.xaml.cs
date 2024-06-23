using PR50.Contexts;
using System.Collections.Generic;
using System.Windows.Controls;

namespace PR50.Elements
{
    /// <summary>
    /// Логика взаимодействия для Room.xaml
    /// </summary>
    public partial class Room : UserControl
    {
        public Room(int Room)
        {
            InitializeComponent();
            NameRoom.Content = "Квартира №" + Room;
            LoadOwner(Room);
        }
        public void LoadOwner(int Room)
        {
            List<OwnerContext> roomOwners = OwnerContext.AllOwners().FindAll(x => x.NumberRoom == Room);
            foreach (OwnerContext roomOwner in roomOwners)
                Parent.Children.Add(new Elements.Ower(roomOwner));
        }
    }
}
