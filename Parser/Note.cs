namespace Parser
{
    public class Note
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public string Threat { get; set; }
        public string IsNotConfidential { get; set; }
        public string IsComplete { get; set; }
        public string IsAccessible { get; set; }
        public string DateIn { get; set; }
        public string DateRewrite { get; set; }

        public Note(int id, string name, string description, string source, string threat, string isNotConfidential, string isComplete, string isAccessible, string dateIn, string dateRewrite)
        {
            Id = id;
            Name = name;
            Description = description;
            Source = source;
            Threat = threat;
            IsNotConfidential = isNotConfidential;
            IsComplete = isComplete;
            IsAccessible = isAccessible;
            DateIn = dateIn;
            DateRewrite = dateRewrite;
        }

        public override string ToString()
        {
            return $"{Id}, {Name}";
        }
    }
}
