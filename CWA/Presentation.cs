using System;

namespace CWA
{
    internal class Presentation
    {
        public object Slides { get; internal set; }
        public object LayoutSlides { get; internal set; }
        public object SlideMaster { get; internal set; }

        internal void Save(string v, object pptx)
        {
            throw new NotImplementedException();
        }
    }
}