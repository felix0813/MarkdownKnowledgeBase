namespace MarkdownKnowledgeBase
{
    public record NoteItem(string Name, string Path);

    public record NoteOption(string DisplayName, string Path);

    public record Marker(string Id, string Name, string NotePath, int Position);

    public record Link(string Id, string SourceMarkerId, string TargetMarkerId);

    public record NavigationEntry(string NotePath, int Position);

    public record LineLink(string Id, string SourceNotePath, int SourceLineIndex, string TargetNotePath, int TargetPosition);
}
