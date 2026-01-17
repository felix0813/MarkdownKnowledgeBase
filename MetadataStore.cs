using System.IO;
using System.Text.Json;

namespace MarkdownKnowledgeBase
{
    public class MetadataStore
    {
        public List<Marker> Markers { get; set; } = new();
        public List<Link> Links { get; set; } = new();

        public static MetadataStore Load(string path)
        {
            if (!File.Exists(path))
            {
                return new MetadataStore();
            }

            var json = File.ReadAllText(path);
            return JsonSerializer.Deserialize<MetadataStore>(json) ?? new MetadataStore();
        }

        public void Save(string path)
        {
            var json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(path, json);
        }
    }
}
