using System.Text.Json.Serialization;

namespace JitsiMeetOutlook.Entities
{
    /// <summary>
    /// Conference Mapper Endpoint Response
    /// Derived from https://github.com/jitsi/jitsi-meet/blob/master/resources/cloud-api.swagger
    /// </summary>
    internal class ConferenceMapperResponse
    {
        /// <summary>
        /// Message from the server.
        /// </summary>
        /// <value></value>
        [JsonPropertyName("message")]
        public string Message { get; set; }

        /// <summary>
        /// Conference Name.
        /// </summary>
        /// <value></value>
        [JsonPropertyName("conference")]
        public string conference { get; set; }

        /// <summary>
        /// Id of the Conference (also the PIN).
        /// </summary>
        /// <value></value>
        [JsonPropertyName("id")]
        public ulong Id { get; set; }

        [JsonIgnore]
        public string IdString => Id.ToString();
    }
}
