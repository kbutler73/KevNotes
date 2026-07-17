using System;
using System.Windows.Documents;

namespace KevNotes
{
    /// <summary>
    /// Convenience extensions for TextPointer to provide missing helper methods.
    /// </summary>
    internal static class TextPointerExtensions
    {
        /// <summary>
        /// Deletes the content between the start and end TextPointer instances.
        /// This mirrors the expected semantics of a DeleteContentTo helper.
        /// If either pointer is null or they are equal, the method returns without action.
        /// </summary>
        /// <param name="start">The start TextPointer.</param>
        /// <param name="end">The end TextPointer.</param>
        public static void DeleteContentTo(this TextPointer start, TextPointer end)
        {
            if (start == null || end == null)
            {
                return;
            }

            if (start.CompareTo(end) == 0)
            {
                return;
            }

            // Use TextRange to remove content. This will remove text and any inline formatting
            // inside the specified range.
            var range = new TextRange(start, end);
            range.Text = string.Empty;
        }
    }
}
