

using Newtonsoft.Json;

namespace Model
{
    public class EnModel : BaseModel
    {
        [JsonProperty(Order = 2)]
        public string EnValue { get; set; }
    }
}
