namespace ImageToTextReader.Models
{
    public class OCRModel
    {
        public string DetectedText { get; set; }
        public IFormFile ImageFile { get; set; }
    }
}
