

using Newtonsoft.Json;

namespace Model
{
    public class RuModel : BaseModel
    {
         [JsonProperty(Order = 2)]
        public string RuValue { get; set; }
    }
}
