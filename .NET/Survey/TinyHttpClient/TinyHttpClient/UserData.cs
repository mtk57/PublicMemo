using System.Runtime.Serialization;

namespace MyData
{
    [DataContract]
    public class UserData
    {
        [DataMember(Name = "id")]
        public int Id { get; set; }

        [DataMember(Name = "name")]
        public string Name { get; set; }
    }
}
