namespace PictureBoxReplace
{
    internal class DataModel
    {
        public string ResxPath { get; set; }
        public string PictureBoxName { get; set; }
        public string ImageData { get; set; }  //MUST
        public string ReplaceImagePath { get; set; }  //MUST
        public string TargetDirPath { get; set; }   // MUST
        public bool IsBackup { get; set; }

        public string ReplaceImageData { get; set; }
    }
}
