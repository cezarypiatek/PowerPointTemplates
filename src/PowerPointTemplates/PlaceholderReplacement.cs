namespace PowerPointTemplates
{
    class PlaceholderReplacement
    {
        public string Key { get; }
        public string Value { get; }

        public PlaceholderReplacement(string key, string value)
        {
            Key = key;
            Value = value;
        }
    }
}