using DocumentFormat.OpenXml.Packaging;

namespace OpenXMLSDK.Engine.Word.Parts
{
    public static class PartsExtensions
    {
        /// <summary>
        /// Adds a new part to the document
        /// </summary>
        /// <typeparam name="T">Type of the object</typeparam>
        /// <param name="wordManager">Word document manager</param>
        /// <returns>Created part</returns>
        public static T AddNewPart<T>(this WordManager wordManager) where T : OpenXmlPart, IFixedContentTypePart
        {
            return wordManager.MainDocumentPart.AddNewPart<T>();
        }

        /// <summary>
        /// Returns the ID of a part in the document
        /// </summary>
        /// <typeparam name="T">Type of the part</typeparam>
        /// <param name="wordManager">Word document manager</param>
        /// <param name="part">Part</param>
        /// <returns>ID of the part in the document</returns>
        public static string GetIdOfPart<T>(this WordManager wordManager, T part) where T : OpenXmlPart
        {
            return wordManager.MainDocumentPart.GetIdOfPart(part);
        }
    }
}
