namespace WpfAppGetLoginByNameFromAD.Repository.Entities
{
    public class UserEntity : BaseEntity
    {
        public string Login { get; set; }

        public string Email { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public string DisplayName { get; set; }
    }
}
