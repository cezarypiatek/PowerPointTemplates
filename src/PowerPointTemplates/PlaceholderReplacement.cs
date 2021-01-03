namespace PowerPointTemplates
{
    class PlaceholderReplacement
    {
        public string Key { get; set; }
        public string Value { get; set; }

        public PlaceholderReplacement()
        {
            
        }

        public PlaceholderReplacement(string key, string value)
        {
            Key = key;
            Value = value;
        }
    }
}