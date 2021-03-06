﻿using Newtonsoft.Json;

namespace NLog.Targets.MicrosoftTeams
{
    internal class MicrosoftTeamsMessageFact
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }
    }
}
