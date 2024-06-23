using PR50.Contexts;
using System.Windows.Controls;

namespace PR50.Elements
{
    /// <summary>
    /// Логика взаимодействия для Owner.xaml
    /// </summary>
    public partial class Owner : UserControl
    {
        public Owner(OwnerContext roomOwner)
        {
            InitializeComponent();
            NameOwner.Content = $"{roomOwner.Surname} {roomOwner.Name} {roomOwner.Lastname}";
        }
    }
}
