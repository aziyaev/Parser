namespace Parser
{
    public class Note
    {
        public int Id { get; set; }
        public string IdInfo { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string Source { get; set; }
        public string Threat { get; set; }
        public string IsNotConfidential { get; set; }
        public string IsComplete { get; set; }
        public string IsAccessible { get; set; }
        public string DateIn { get; set; }
        public string DateRewrite { get; set; }

        public Note(int id, string idInfo, string name, string description, string source, string threat, string isNotConfidential, string isComplete, string isAccessible, string dateIn, string dateRewrite)
        {
            Id = id;
            IdInfo = idInfo;
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
            return $"{IdInfo}\n\nНаименование: {Name}\n\nОписание: {Description}\n\n" +
                $"Источник: {Source}\n\nОбъект воздействия: {Threat}\n\n" +
                $"Нарушение конфиденциальности: {IsNotConfidential}\n\nНарушение целостности: {IsComplete}\n\nНарушение доступности: {IsAccessible}\n\n" +
                $"Дата создания: {DateIn}\n\nДата изменения: {DateRewrite}";
        }
    }
}
